#!/usr/bin/env python3
"""営業定例MTG資料を Notion の各DB（顧客・受注・見積・目標）から集計し、
指定した Notion ページに「全社着地・ファネル・ヨミ・個別案件・原因分析」を生成する。

使い方:
  python3 generate_report.py --page-id <NotionページID> [--current-month 2026-07] [--dry-run]

トークンは既定で .tmp/notion_token.txt の「アクセストークン=...」行から読む。
3. 前回アクション棚卸し は議事録が一次情報のため既定で空欄（--minutes 未指定時）。
"""
import argparse, urllib.request, json, sys, os, time
from collections import defaultdict
from datetime import date, datetime

# ===== プロジェクト固定設定（VJ 営業DB）=====
DB = {
    'customer': 'cd305332579a8291ae2301d1744dcb0e',
    'quote':    '24305332-579a-8256-b325-0178de5b552c',
    'order':    '2f005332-579a-83bf-8b35-01fe09c84aad',
    'goal':     '37b05332-579a-81ec-aac2-d6fff5706222',
}
# 出力先（MTG DB）。新規ページを作成し、そのページ本文に資料を書き込む
OUTPUT_DB = '24f05332-579a-8047-a00d-f092eb1c6c2c'
AGENDA_NAME_PREFIX = '営業定例アジェンダ-'  # + YYYYMMDD（作成日）
AGENDA_PLACE = 'VISIONARY JAPAN本社'        # 場所(select)
AGENDA_CATEGORY = '営業定例'                # カテゴリ(select)
# 担当者名(formula) → 表示名。
# 登録データは実名が正（妻鹿一大 / 古屋喬士）。旧英名（Mega / furuyatakashi）も互換で同一人物として扱う。
PERSON_MAP = {'妻鹿一大': '妻鹿', '古屋喬士': '喬士', 'Mega': '妻鹿', 'furuyatakashi': '喬士'}
MAIN = ['妻鹿', '喬士']                 # 個人別に表示する担当
FISCAL_START = 3                        # 会計年度開始月（3月始まり）
PHASE_ROWS = ['初回再調整', '商談化', '初回商談済_お断り/保留', '初回商談済_先方社内確認',
    '初回商談済_実現性調査', '初回商談済_NDA締結+調査', '初回商談済_再訪調整中', '初回商談済_再訪設定済',
    '提案/見積り準備', '提案後検討', '稟議中', '契約手続き（発注書未締結）', '受注（初回締結完了）']
STAGES = [('1 初回再調整', ['初回再調整']), ('2 商談化', ['商談化']),
    ('3 初回商談済', ['初回商談済_お断り/保留', '初回商談済_先方社内確認', '初回商談済_実現性調査',
        '初回商談済_NDA締結+調査', '初回商談済_再訪調整中', '初回商談済_再訪設定済']),
    ('4 提案/見積り準備', ['提案/見積り準備']), ('5 提案後検討', ['提案後検討']),
    ('6 稟議中', ['稟議中']), ('7 契約手続き', ['契約手続き（発注書未締結）']),
    ('8 受注（初回締結完了）', ['受注（初回締結完了）', '継続延長交渉中'])]
WON = ['受注（初回締結完了）']
LEVELS = ['A：決裁者が前向き', 'B：決裁者と接点あり', 'C：キーマンと接点あり', 'D：担当者どまり']
LEVEL_W = {'A：決裁者が前向き': .9, 'B：決裁者と接点あり': .6, 'C：キーマンと接点あり': .3, 'D：担当者どまり': .1}
DEPTH = {'提案/見積り準備': 9, '初回商談済_NDA締結+調査': 7, '初回商談済_実現性調査': 6,
    '初回商談済_再訪調整中': 5, '初回商談済_再訪設定済': 5, '初回商談済_先方社内確認': 4,
    '商談化': 2, '初回再調整': 1}

# ===== Notion API =====
TOKEN = None
def api(url, method='GET', data=None):
    r = urllib.request.Request(url, method=method)
    r.add_header('Authorization', 'Bearer ' + TOKEN)
    r.add_header('Notion-Version', '2022-06-28')
    r.add_header('Content-Type', 'application/json')
    if data is not None:
        r.data = json.dumps(data).encode()
    try:
        with urllib.request.urlopen(r) as x:
            return x.status, json.loads(x.read().decode())
    except urllib.error.HTTPError as e:
        return e.code, json.loads(e.read().decode())
def qall(did):
    out = []; cur = None
    while True:
        p = {'page_size': 100}
        if cur: p['start_cursor'] = cur
        s, b = api('https://api.notion.com/v1/databases/' + did + '/query', 'POST', p)
        if s != 200:
            sys.exit('DB取得失敗 %s: %s' % (did, b.get('message')))
        out += b['results']; cur = b.get('next_cursor')
        if not b.get('has_more'): break
    return out
def gc(bid):
    out = []; cur = None
    while True:
        u = 'https://api.notion.com/v1/blocks/' + bid + '/children?page_size=100'
        if cur: u += '&start_cursor=' + cur
        s, b = api(u)
        out += b['results']; cur = b.get('next_cursor')
        if not b.get('has_more'): break
    return out

# ===== helpers =====
def fstr(p): return (p.get('formula', {}) or {}).get('string')
def man(n): return f"{round((n or 0)/10000):,}万"
def pct(a, b): return '―' if b == 0 else f'{round(a/b*100,1)}%'
def who(tanto): return PERSON_MAP.get(tanto, 'その他')
def rollup_status(pr):
    for it in pr.get('受注/close', {}).get('rollup', {}).get('array', []):
        if it.get('type') == 'select' and it.get('select'):
            return it['select']['name']
    return None

# ===== 会計月ユーティリティ =====
def fiscal_months(cur_ym):
    """会計年度(3月始まり)の 開始月〜当月+1 の 'YYYY-MM' リスト(昇順)を返す"""
    y, m = int(cur_ym[:4]), int(cur_ym[5:7])
    fy = y if m >= FISCAL_START else y - 1
    seq = [((FISCAL_START + i - 1) % 12) + 1 for i in range(12)]  # [3,4,..2]
    months = []
    yy = fy
    for mm in seq:
        if mm < FISCAL_START and len(months) > 0:  # 翌年に入った
            pass
        months.append((yy if mm >= FISCAL_START else yy + 1, mm))
    res = []
    for (yr, mm) in months:
        res.append(f'{yr:04d}-{mm:02d}')
    # 当月+1 まで
    end = f'{y:04d}-{m:02d}'
    idx = res.index(end) if end in res else len(res) - 1
    return res[:min(idx + 2, len(res))], fy
def quarter_of(ym):
    m = int(ym[5:7])
    pos = (m - FISCAL_START) % 12
    return pos // 3 + 1  # 1..4

# ===== 集計 =====
def collect(cur_ym, rvc_exclude=True):
    custs = qall(DB['customer'])
    cust = []; cmap = {}
    for pg in custs:
        pr = pg['properties']
        naf = pr.get('NA日', {}).get('formula', {}) or {}
        o = {'tanto': fstr(pr.get('担当者名', {})),
             'phase': (pr.get('商談フェーズ', {}).get('select') or {}).get('name'),
             'level': (pr.get('接触レベル', {}).get('select') or {}).get('name'),
             'st': (pr.get('受注/close/継続', {}).get('select') or {}).get('name'),
             'ryu': (pr.get('流入経路', {}).get('select') or {}).get('name'),
             'name': ''.join(t.get('plain_text', '') for t in (pr.get('顧客名', {}).get('title') or [])).strip(),
             'nad': naf.get('date', {}).get('start') if naf.get('type') == 'date' else naf.get('string'),
             'na': fstr(pr.get('NA', {}))}
        cust.append(o); cmap[pg['id']] = o
    def rows(w):
        if w == '組織': return cust
        return [c for c in cust if who(c['tanto']) == w]

    # フェーズ件数
    phase_cnt = {ph: {w: sum(c['phase'] == ph for c in rows(w)) for w in ['組織'] + MAIN} for ph in PHASE_ROWS}
    # 8ステージ移行率
    def stage_cum(rs):
        here = [sum(r['phase'] in phs for r in rs) for _, phs in STAGES]
        cum = [0]*len(STAGES); acc = 0
        for i in range(len(STAGES)-1, -1, -1):
            acc += here[i]; cum[i] = acc
        return cum
    cums = {w: stage_cum(rows(w)) for w in ['組織'] + MAIN}
    # 接触レベル
    level_cnt = {l: {w: sum(c['level'] == l for c in rows(w)) for w in ['組織']+MAIN} for l in LEVELS}
    level_none = {w: sum(not c['level'] for c in rows(w)) for w in ['組織']+MAIN}
    level_wt = {w: round(sum(LEVEL_W.get(c['level'], 0) for c in rows(w)), 1) for w in ['組織']+MAIN}
    # 歩留り(受注初回締結)
    yield_ = {w: (sum(r['phase'] in WON for r in rows(w)), len(rows(w))) for w in ['組織']+MAIN}

    # 受注DB
    orders = qall(DB['order'])
    conf = defaultdict(lambda: {'c': 0, 'a': 0})   # 月別確定(MAIN担当)
    mikomi = {'c': 0, 'a': 0}                       # 受注日未入力(MAIN担当)
    by_person = defaultdict(lambda: {'c': 0, 'a': 0})
    rvc = defaultdict(lambda: {'c': 0, 'a': 0})
    fy26 = defaultdict(lambda: {'c': 0, 'a': 0})
    fy_start_ym = cur_ym  # for FY range
    fyy = int(cur_ym[:4]) if int(cur_ym[5:7]) >= FISCAL_START else int(cur_ym[:4]) - 1
    fy_lo = f'{fyy:04d}-{FISCAL_START:02d}'
    fy_hi = f'{fyy+1:04d}-{FISCAL_START-1:02d}' if FISCAL_START > 1 else f'{fyy:04d}-12'
    for o in orders:
        pr = o['properties']; amt = pr.get('受注金額', {}).get('number') or 0
        d = (pr.get('受注日(契約締結日)', {}).get('date') or {}).get('start')
        rel = pr.get('DB_顧客情報', {}).get('relation') or []
        ci = cmap.get(rel[0]['id'], {}) if rel else {}
        w = who(ci.get('tanto'))
        by_person[w]['c'] += 1; by_person[w]['a'] += amt
        if ci.get('ryu') == 'RVC':
            rvc[w]['c'] += 1; rvc[w]['a'] += amt
        if d and fy_lo <= d[:7] <= fy_hi:
            fy26[w]['c'] += 1; fy26[w]['a'] += amt
        if w in MAIN:
            if d:
                conf[d[:7]]['c'] += 1; conf[d[:7]]['a'] += amt
            else:
                mikomi['c'] += 1; mikomi['a'] += amt

    # 見積DB(継続ヨミ, MAIN担当)
    quotes = qall(DB['quote'])
    yomi = defaultdict(lambda: {'c': 0, 'a': 0})
    yomi_person = defaultdict(lambda: {'c': 0, 'a': 0, 'items': []})
    for q in quotes:
        pr = q['properties']; amt = pr.get('見積金額', {}).get('number') or 0
        land = (pr.get('着地予想月(月末を記載)', {}).get('date') or {}).get('start')
        rel = pr.get('DB_顧客情報', {}).get('relation') or []
        ci = cmap.get(rel[0]['id'], {}) if rel else {}
        w = who(ci.get('tanto'))
        if rollup_status(pr) != '継続' or w not in MAIN:
            continue
        m = land[:7] if land else '未設定'
        yomi[m]['c'] += 1; yomi[m]['a'] += amt
        yomi_person[w]['c'] += 1; yomi_person[w]['a'] += amt
        yomi_person[w]['items'].append({'name': ci.get('name'), 'amt': amt, 'land': m})

    # 上位案件(妻鹿 生存)
    alive = [c for c in rows('妻鹿') if c['st'] not in ('close', 'close(熱)')]
    alive.sort(key=lambda c: -DEPTH.get(c['phase'], 0))

    return dict(rows=rows, phase_cnt=phase_cnt, cums=cums, level_cnt=level_cnt,
        level_none=level_none, level_wt=level_wt, yield_=yield_, conf=conf, mikomi=mikomi,
        by_person=by_person, rvc=rvc, fy26=fy26, yomi=yomi, yomi_person=yomi_person,
        alive=alive, cust=cust, rvc_exclude=rvc_exclude)

# ===== ブロック生成 =====
def rt(s, b=False): return [{"type": "text", "text": {"content": str(s)}, "annotations": {"bold": b}}]
def h1(s): return {"object": "block", "type": "heading_1", "heading_1": {"rich_text": rt(s, True)}}
def h2(s): return {"object": "block", "type": "heading_2", "heading_2": {"rich_text": rt(s, True)}}
def h3(s): return {"object": "block", "type": "heading_3", "heading_3": {"rich_text": rt(s, True)}}
def para(s): return {"object": "block", "type": "paragraph", "paragraph": {"rich_text": rt(s)}}
def bp(s): return {"object": "block", "type": "paragraph", "paragraph": {"rich_text": rt(s, True)}}
def bullet(s): return {"object": "block", "type": "bulleted_list_item", "bulleted_list_item": {"rich_text": rt(s)}}
def callout(s, e, c): return {"object": "block", "type": "callout", "callout": {"rich_text": rt(s), "icon": {"type": "emoji", "emoji": e}, "color": c}}
def div(): return {"object": "block", "type": "divider", "divider": {}}
def table(h, rows):
    tr = [{"object": "block", "type": "table_row", "table_row": {"cells": [rt(x) for x in h]}}]
    for r in rows:
        tr.append({"object": "block", "type": "table_row", "table_row": {"cells": [rt(x) for x in r]}})
    return {"object": "block", "type": "table", "table": {"table_width": len(h), "has_column_header": True, "has_row_header": False, "children": tr}}

def build_blocks(D, cur_ym):
    B = []
    conf, mikomi, yomi = D['conf'], D['mikomi'], D['yomi']
    months, fy = fiscal_months(cur_ym)
    # 当月・四半期
    cur_m = cur_ym
    def q_sum_conf(qn):
        return sum(conf[m]['a'] for m in conf if quarter_of(m) == qn and m[:4] <= cur_ym[:4])
    def q_sum_yomi(qn):
        return sum(yomi[m]['a'] for m in yomi if m != '未設定' and quarter_of(m) == qn)
    all_conf = sum(v['a'] for v in conf.values())
    all_yomi = sum(v['a'] for m, v in yomi.items() if m != '未設定')

    # --- 1 ---
    B += [h1("1. 全社着地と進捗")]
    mk = man(mikomi['a']); cm = conf.get(cur_m, {'a': 0, 'c': 0})
    B += [callout(f"当月{int(cur_ym[5:7])}月の受注見込は{mk}（受注日未入力{mikomi['c']}件）＋確定{man(cm['a'])}。担当は{ '/'.join(MAIN) }、流入経路RVCは除外集計。", "🎯", "blue_background")]
    B += [h3("1-1. 目標・着地見込・ギャップ・率（組織）")]
    trows = []
    trows.append(["26年 年間", "―（PL未集約）", f"{man(all_conf+mikomi['a'])}（内見込{man(mikomi['a'])}）", man(all_yomi), "―", "0%", "集計未確定"])
    done_q = sorted({quarter_of(m) for m in months})
    for qn in done_q:
        qc, qy = q_sum_conf(qn), q_sum_yomi(qn)
        lbl = {1: "第1Q（3-5月）", 2: "第2Q（6-8月）", 3: "第3Q（9-11月）", 4: "第4Q（12-2月）"}[qn]
        trows.append([lbl, "―（PL未集約）", man(qc) if qc else "―", f"{man(qy)}" if qy else "―", "―", "0%", "四半期集計"])
    for m in sorted(months, reverse=True):
        mm = int(m[5:7])
        rc = conf.get(m, {'a': 0, 'c': 0}); ry = yomi.get(m, {'a': 0, 'c': 0})
        if m == cur_m:
            oc = f"見込{man(mikomi['a'])}（{mikomi['c']}件）＋確定{man(rc['a'])}" if mikomi['a'] else (man(rc['a'])+f"（確定{rc['c']}件）" if rc['a'] else "―")
        else:
            oc = f"{man(rc['a'])}（確定{rc['c']}件）" if rc['a'] else "―"
        yc = f"{man(ry['a'])}（{ry['c']}件）" if ry['a'] else "―"
        trows.append([f"{mm}月", "―（PL未集約）", oc, yc, "―", "―", ""])
    B += [table(["期間", "売上目標", "受注(確定/見込)", "ご提案済みヨミ", "GAP", "進捗率(実績)", "備考"], trows)]
    B += [para("受注はDB_受注情報の実データ。月別＝受注日が該当月の確定受注、見込＝受注日未入力分（当月に計上）。ご提案済みヨミはDB_見積情報管理で受注/close＝継続・着地予想月別。年間・四半期は集計未確定のため目標は空欄。")]
    # 1-2 個人別
    B += [h3("1-2. 実績（個人別）")]
    note = "目標は事業部単位で設定。" + ("実績は流入経路RVC（外部パートナー経由）を除外して集計。" if D['rvc_exclude'] else "")
    B += [para(note)]
    p_rows = []
    org_ord = 0
    for w in MAIN + ['その他']:
        bp_ = D['by_person'][w]; rv = D['rvc'][w]; fy = D['fy26'][w]
        a = bp_['a'] - (rv['a'] if D['rvc_exclude'] else 0)
        c = bp_['c'] - (rv['c'] if D['rvc_exclude'] else 0)
        org_ord += a
        yp = D['yomi_person'].get(w, {'a': 0, 'c': 0})
        if w in MAIN:
            p_rows.append([w, f"{man(a)} / {c}件", f"{man(fy['a'])} / {fy['c']}件", f"{man(yp['a'])} / {yp['c']}件"])
        else:
            p_rows.append([w + "（社長案件等）", f"{man(a)} / {c}件", "―", "―"])
    B += [table(["担当", "受注（累計/件）", "うちFY26内", "継続ヨミ（見積/件）"], p_rows)]
    rvc_org = sum(D['rvc'][w]['a'] for w in D['rvc'])
    rvc_orgc = sum(D['rvc'][w]['c'] for w in D['rvc'])
    if D['rvc_exclude']:
        B += [bullet(f"RVC除外後の組織 受注累計は{man(org_ord)}（RVC流入{rvc_orgc}件・{man(rvc_org)}を除外）。")]
    B += [div()]

    # --- 2 ---
    B += [h1("2. ファネル＋ヨミ精度")]
    B += [h3("2-1. フェーズ別 件数（組織 / " + " / ".join(MAIN) + "）")]
    stage_idx = {}
    for i, (nm, phs) in enumerate(STAGES):
        for ph in phs: stage_idx[ph] = nm.split()[0]
    crows = []
    for ph in PHASE_ROWS:
        crows.append([f"{stage_idx.get(ph,'')} {ph}"] + [str(D['phase_cnt'][ph][w]) for w in ['組織']+MAIN])
    crows.append(["合計"] + [str(len(D['rows'](w))) for w in ['組織']+MAIN])
    B += [table(["商談フェーズ", "組織"] + MAIN, crows)]
    B += [bp("ファネル移行率（8ステージ集約・到達累計ベース）")]
    rrows = []
    for i, (nm, _) in enumerate(STAGES):
        row = [nm]
        for w in ['組織'] + MAIN:
            cum = D['cums'][w]
            row += [str(cum[i]), '―' if i == 0 else pct(cum[i], cum[i-1])]
        rrows.append(row)
    trow = ["通算（1→8）"]
    for w in ['組織'] + MAIN:
        cum = D['cums'][w]; trow += [str(cum[-1]), pct(cum[-1], cum[0])]
    rrows.append(trow)
    hdr = ["ステージ"]
    for w in ['組織'] + MAIN: hdr += [f"{w} 到達", f"{w} 移行率"]
    B += [table(hdr, rrows)]
    # 妻鹿 提案転換に着目（MAINの1人目）
    m0 = MAIN[0]
    B += [callout(f"{m0}の失速点は「3 初回商談済 → 4 提案」の {pct(D['cums'][m0][3], D['cums'][m0][2])}（組織{pct(D['cums']['組織'][3], D['cums']['組織'][2])}）。初回商談は通過するが提案フェーズに引き上げられていない。", "🔎", "yellow_background")]
    B += [para("移行率＝そのステージの到達件数 ÷ ひとつ手前のステージの到達件数（到達累計ベース）。ステージ3は初回商談済の6フェーズを集約。時系列の移行履歴ではなくスナップショット近似。")]
    B += [h3("2-2. 全体歩留り（リード→受注）")]
    yrow = ["歩留り"]
    for w in ['組織'] + MAIN:
        wn, tot = D['yield_'][w]; yrow.append(f"{pct(wn,tot)}（{wn}/{tot}）")
    B += [table(["区分", "組織"] + MAIN, [yrow])]
    B += [para("受注＝「受注（初回締結完了）」フェーズ到達件数。")]
    B += [h3("2-3. ヨミ加重（接触レベル A/B/C/D）")]
    lrows = []
    for l in LEVELS:
        lrows.append([l] + [str(D['level_cnt'][l][w]) for w in ['組織']+MAIN])
    lrows.append(["未設定"] + [str(D['level_none'][w]) for w in ['組織']+MAIN])
    B += [table(["接触レベル", "組織"] + MAIN, lrows)]
    wt = D['level_wt']
    B += [callout(f"件数加重（A0.9/B0.6/C0.3/D0.1）は 組織{wt['組織']}・" + "・".join(f"{w}{wt[w]}" for w in MAIN) + "。決裁者接点(A+B)の薄さが受注率に直結。", "⚠️", "red_background")]
    B += [div()]

    # --- 3 空欄 ---
    B += [h1("3. 前回アクション棚卸し")]
    B += [para("一次情報（営業部MTG議事録）が未取得のため、本セクションは空欄。議事録が取得でき次第、アクションリストを記載する。")]
    B += [div()]

    # --- 4 個別案件 ---
    B += [h1("4. 個別案件 深掘り（上位案件）")]
    B += [h2(MAIN[0])]
    arows = []
    for c in D['alive'][:6]:
        nad = c['nad'][:10] if c['nad'] else '未設定'
        arows.append([c['name'], "―", c['phase'], (c['na'] or '')[:22], nad])
    B += [table(["顧客名", "金額", "ヨミ（フェーズ）", "停滞理由/状況", "次アクション(NA日)"], arows)]
    if len(MAIN) > 1:
        B += [h2(MAIN[1])]
        yp = sorted(D['yomi_person'].get(MAIN[1], {'items': []})['items'], key=lambda x: -x['amt'])
        qrows = [[it['name'], man(it['amt']), f"{int(it['land'][5:7])}月" if it['land'] != '未設定' else '未設定', "ご提案済みヨミ"] for it in yp[:8]]
        B += [table(["顧客名", "金額", "着地予想", "ヨミ区分"], qrows)]
    B += [div()]

    # --- 5 原因とアクション ---
    B += [h1("5. 達成に向けた原因とアクションの切り分け")]
    m0 = MAIN[0]
    n_teian = D['phase_cnt']['提案/見積り準備'][m0]
    n_shanai = D['phase_cnt']['初回商談済_先方社内確認'][m0]
    conv = pct(D['cums'][m0][3], D['cums'][m0][2])
    B += [h2(m0)]
    B += [table(["項目", "内容"], [
        ["問題", f"提案/見積り準備が{n_teian}件、先方社内確認が{n_shanai}件と滞留。受注実績が乏しい。"],
        ["原因", "初回商談で決裁者を巻き込めず「社内で確認します」預かりで止まり、状況ヒアリングを繰り返す受け身構造。"],
        ["課題", f"初回商談内での次回商談日確定と決裁者同席の打診ができていない。「初回商談→提案」の移行率{conv}が最大のボトルネック。"],
        ["打ち手", f"先方社内確認{n_shanai}件を棚卸しし、ヒアリング架電で終わらせず提案・見積り準備の再訪アポへ引き上げる。提案準備{n_teian}件は確実にクローズまで進める。"]]) ]
    if len(MAIN) > 1:
        m1 = MAIN[1]
        n_t1 = D['phase_cnt']['提案/見積り準備'][m1]
        B += [h2(m1)]
        B += [table(["項目", "内容"], [
            ["問題", "見積提案後のヨミ案件がcloseとなった場合の角度が高い提案準備案件数の不足。"],
            ["原因", f"現状の提案準備案件は{n_t1}件あるが、決裁者まで達していない。"],
            ["課題", "決裁者同席の打診が弱く、できていない案件がある。"],
            ["打ち手", "担当者と密に連絡をとり、温度感を高めた上で決裁者同席を打診する。"]]) ]
    return B

# ===== ページ書き込み =====
def append_blocks(page_id, blocks, start=0):
    for i in range(start, len(blocks), 12):
        s, b = api('https://api.notion.com/v1/blocks/' + page_id + '/children', 'PATCH', {'children': blocks[i:i+12]})
        if s != 200:
            sys.exit('append失敗: ' + json.dumps(b, ensure_ascii=False)[:400])
        time.sleep(0.35)

def rewrite_page(page_id, blocks):
    """既存ページの本文を全削除して再構築（--page-id 指定時）"""
    old = gc(page_id)
    for blk in old:
        if blk['type'] in ('child_page', 'transcription'):
            continue
        api('https://api.notion.com/v1/blocks/' + blk['id'], 'DELETE'); time.sleep(0.12)
    append_blocks(page_id, blocks)

def create_in_db(db_id, blocks, today):
    """MTG DB に新規ページを作成し、本文に資料を書き込む。作成したページIDを返す"""
    ymd = today.strftime('%Y%m%d')
    props = {
        '名前':     {'title': [{'text': {'content': AGENDA_NAME_PREFIX + ymd}}]},
        '日付':     {'date': {'start': today.strftime('%Y-%m-%d')}},
        '場所':     {'select': {'name': AGENDA_PLACE}},
        'カテゴリ': {'select': {'name': AGENDA_CATEGORY}},
    }
    s, b = api('https://api.notion.com/v1/pages', 'POST',
               {'parent': {'database_id': db_id}, 'properties': props, 'children': blocks[:20]})
    if s != 200:
        sys.exit('ページ作成失敗: ' + json.dumps(b, ensure_ascii=False)[:400])
    pid = b['id']
    append_blocks(pid, blocks, start=20)
    return pid, b.get('url')

def load_token(path):
    with open(path, encoding='utf-8') as f:
        for line in f:
            if 'アクセストークン' in line and '=' in line:
                return line.split('=', 1)[1].strip()
            if line.strip().startswith(('ntn_', 'secret_')):
                return line.strip()
    sys.exit('トークンが見つからない: ' + path)

def main():
    global TOKEN
    ap = argparse.ArgumentParser()
    ap.add_argument('--db-id', default=OUTPUT_DB, help='新規ページ作成先DB（既定=MTG DB）')
    ap.add_argument('--page-id', default=None, help='既存ページを上書きする場合に指定（DB作成の代わり）')
    ap.add_argument('--token-file', default='.tmp/notion_token.txt')
    ap.add_argument('--current-month', default=None, help='YYYY-MM（既定=今日）。集計の対象月')
    ap.add_argument('--no-rvc-exclude', action='store_true', help='RVC除外をしない')
    ap.add_argument('--dry-run', action='store_true', help='集計結果を表示しページは作成/更新しない')
    a = ap.parse_args()
    TOKEN = load_token(a.token_file)
    cur_ym = a.current_month or datetime.now().strftime('%Y-%m')
    D = collect(cur_ym, rvc_exclude=not a.no_rvc_exclude)
    blocks = build_blocks(D, cur_ym)
    print(f'集計完了: 顧客{len(D["cust"])}件 / 生成ブロック{len(blocks)} / 当月{cur_ym}')
    if a.dry_run:
        print('--dry-run: ページ未作成')
        return
    if a.page_id:
        rewrite_page(a.page_id.replace('-', ''), blocks)
        print('既存ページ更新完了:', a.page_id)
    else:
        pid, url = create_in_db(a.db_id.replace('-', ''), blocks, datetime.now())
        print('新規ページ作成完了:', pid)
        if url:
            print('URL:', url)

if __name__ == '__main__':
    main()
