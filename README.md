# 営業定例MTG資料 自動生成 — 変更手順書

このリポジトリは営業定例MTG資料を Notion に自動生成する定期実行ジョブ。
日曜17:00 / 火曜19:00 JST に GitHub Actions が `generate_report.py` を実行し、
MTG DB に「営業定例アジェンダ-YYYYMMDD」ページを新規作成する。

## ファイル構成

| ファイル | 役割 |
|---------|------|
| `generate_report.py` | 集計〜Notionページ生成の本体（**出力内容はここ**） |
| `.github/workflows/sales-mtg-report.yml` | 実行スケジュール・実行手順 |
| `runs.log` | 実行履歴（自動追記・触らない） |

---

## 「何を変えたいか」→「どこを直すか」

### A. 出力される文章・見出し・セクション構成を変える
`generate_report.py` の **`build_blocks()`（225行目〜）**。
セクションはコメントで区切られている:

- `--- 1 ---`（254行目付近）全社着地と進捗
- `1-2 個人別`（281行目付近）
- `--- 2 ---`（284行目付近）ファネル＋ヨミ精度
- `--- 3 空欄 ---`（330行目付近）前回アクション棚卸し（議事録が一次情報のため空欄）
- `--- 4 個別案件 ---`（335行目付近）
- `--- 5 原因とアクション ---`（350行目付近）

文章は `h1("...")`『見出し1』/ `h2` `h3` / `para("...")`『本文』/ `bullet("...")`『箇条書き』/
`callout("...", "絵文字", "色")` の引数を書き換える。表は `table(ヘッダ配列, 行配列)`。

### B. 集計ロジック・数値の出し方を変える
`generate_report.py` の **`collect()`（119行目〜）**。
フェーズ件数・8ステージ移行率・接触レベル加重・歩留り・受注/見積の集計はここ。

### C. 固定設定を変える（冒頭 15〜46行目）

| 変えたいもの | 定数 | 行 |
|-------------|------|----|
| 出力先の Notion DB | `OUTPUT_DB` | 23 |
| ページ名の接頭辞 | `AGENDA_NAME_PREFIX` | 24 |
| ページの場所 / カテゴリ | `AGENDA_PLACE` / `AGENDA_CATEGORY` | 25-26 |
| 担当者名マッピング | `PERSON_MAP` / `MAIN` | 29-30 |
| 会計年度の開始月 | `FISCAL_START` | 31 |
| フェーズ順・8ステージ定義 | `PHASE_ROWS` / `STAGES` | 32-40 |
| 受注扱いのフェーズ名 | `WON` | 41 |
| 接触レベルの加重係数 | `LEVELS` / `LEVEL_W` | 42-43 |

### D. 実行時刻を変える
`.github/workflows/sales-mtg-report.yml` の `schedule:` の `cron`。
**cron は UTC 表記**（JST − 9時間）。JST は夏時間なしなので固定。

- 現状: `0 8 * * 0`（日曜17:00 JST）/ `0 10 * * 2`（火曜19:00 JST）
- 例: 月曜9:00 JST にしたい → 月曜0:00 UTC → `0 0 * * 1`
- 曜日番号: 日=0, 月=1, 火=2 … 土=6

### E. Notion トークンを差し替える
GitHub の `Settings → Secrets and variables → Actions → NOTION_TOKEN` を更新する。
コードやこのリポジトリにトークンを直接書かない。

---

## 変更を反映する手順

### 方法1: GitHub 上で直接編集（軽微な文言・時刻変更向け）
1. GitHub でファイルを開き鉛筆アイコンで編集 → Commit
2. push 不要。次回スケジュールから反映

### 方法2: ローカルで編集（推奨・テストしてから反映）
```bash
cd "/Users/ichidai/Documents/Obsidian Vault/.tmp/sales-mtg-report-cron"

# 1) generate_report.py を編集

# 2) ローカルで集計だけ確認（ページは作らない）
mkdir -p .tmp
grep 'アクセストークン' "/Users/ichidai/Documents/Obsidian Vault/.tmp/notion_token.txt" > .tmp/notion_token.txt
python3 generate_report.py --dry-run
rm -f .tmp/notion_token.txt

# 3) 問題なければ反映
git add generate_report.py .github/workflows/sales-mtg-report.yml
git commit -m "chore: update report content"
git push
```

### 反映後の動作テスト
GitHub の `Actions → sales-mtg-report → Run workflow`（main）で手動実行し、
ログの dry-run 顧客件数・本実行成功・Notion のページ生成を確認する。

---

## 注意: 手動実行版との同期
このリポジトリの `generate_report.py` は、Vault 内スキルからの**コピー**。
Vault の `/aisales-regularmtg`（手動実行）は別ファイル
`.claude/skills/sales-mtg-report/scripts/generate_report.py` を使う。
**手動版と自動版の両方を揃えたい場合は両方を更新する**。片方だけ直すと出力がずれる。

同期する場合（Vault版 → このリポジトリ版へコピー）:
```bash
cp "/Users/ichidai/Documents/Obsidian Vault/.claude/skills/sales-mtg-report/scripts/generate_report.py" \
   "/Users/ichidai/Documents/Obsidian Vault/.tmp/sales-mtg-report-cron/generate_report.py"
```
