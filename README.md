# 宿泊旅行統計ダッシュボード (Tourism Analytics Dashboard)

観光庁・JNTO・インバウンド消費動向調査のデータを統合したダッシュボード。

## データソース
| データ | 出典 | 年次 |
|--------|------|------|
| 延べ宿泊者数・稼働率 | 観光庁「宿泊旅行統計調査」 | 2023年〜2026年（月次） |
| 訪日外国人数（国別） | 日本政府観光局（JNTO） | 2023年〜2026年（月次） |
| 消費単価・費目別・訪問率 | インバウンド消費動向調査 | 2025年（年次確報） |
| 国籍別×都道府県 宿泊者数 | 観光庁「宿泊旅行統計調査」 | 2024年（確定値） |

## ファイル構成
```
tourism-analytics/
├── index.html    ← ダッシュボード本体
├── data.json     ← 統合データ
└── README.md
```

## データ更新方法

### 1. 観光庁 宿泊旅行統計（月次更新）
1. [観光庁サイト](https://www.mlit.go.jp/kankocho/siryou/toukei/shukuhakutoukei.html)から推移表Excelをダウンロード
2. Python変換スクリプトで `data.json` を再生成
3. git push

### 2. JNTO 訪日外客数（月次更新）
1. [JNTO](https://www.jnto.go.jp/statistics/data/visitors-statistics/)からExcelをダウンロード
2. Python変換スクリプトで `data.json` を再生成
3. git push

### 3. 消費動向調査（年次更新）
1. [観光庁サイト](https://www.mlit.go.jp/kankocho/siryou/toukei/syouhityousa.html)からExcelをダウンロード
2. Python変換スクリプトで `data.json` を再生成
3. git push

## デプロイ
GitHub Pages で自動配信。`main` ブランチへの push で自動更新。
公開URL: https://naito1014-ux.github.io/tourism-analytics/

---

# 【保存版】月次更新の完全手順（次回の自分へ）

> データは `data.json` がマスター。**3系統すべて1行コマンドで実行でき**、
> スクリプトが `data.json` 更新と `index.html`（`var RAW = {...}`）の再埋め込みまで自動で行う。
> 作業場所は必ず `~/Desktop/tourism-analytics/`。

## 0. データ3系統と1行コマンド（早見表）
| # | データ | いつ | 入力Excel（目安サイズ） | 1行コマンド | 更新範囲 |
|---|--------|------|----------------|-------------|----------|
| A | 全国 月次速報（第1次速報） | 毎月 | `第1表(N月)`を含む速報（約26KB） | `python3 update_monthly.py 202604_宿泊統計.xlsx` | 全国 t/j/f/o のみ |
| B | 都道府県別 確報（第2次確報） | 確報公表時 | `第2表/第8表(N月)`を含む確報（約357KB） | `python3 update_data.py --kakuho 202603_宿泊統計.xlsx` | 全国＋都道府県47件 t/j/f/o |
| C | JNTO 訪日外客数 | 毎月 | 月別推計値（`YYYY.MM`シート、約18KB） | `python3 update_data.py --jnto-monthly 202604_JNTO.xlsx 202605_JNTO.xlsx` | jd/jg（国別実数・伸率） |

- A→Bの順序：速報Aで全国を先行更新し、後日確報Bが出たら**案A（後述）で全国を確報値に上書き**して都道府県と揃える。
- Cは独立。毎月AとCがセットで出る。**Cは複数月をスペース区切りでまとめて指定可**。
- B/C はどちらも `data.json` 更新後に `index.html` を自動で再埋め込みする（手動スニペット不要）。

## 1. 入手先URLとダウンロードするExcel
### A/B 観光庁 宿泊旅行統計
- 一覧: https://www.mlit.go.jp/kankocho/siryou/toukei/shukuhakutoukei.html
- 各月の報道発表（news02系ページ）から Excel を取得：
  - **A 全国速報（第1次速報）**：`第1表(N月)`〜`第5表(N月)` 程度の小さいファイル（**約26KB**）。全国の延べ・稼働率のみ。
  - **B 都道府県別 確報（第2次確報）**：`第1表`〜`第10表`＋`参考表`を含む大きいファイル（**約357KB**）。都道府県別・国籍別を含む。
### C JNTO 訪日外客数
- https://www.jnto.go.jp/statistics/data/visitors-statistics/
- 月別の**推計値**Excel（シート名が `YYYY.MM` 形式。例 `2026.04`）。

## 2. ファイル名の付け方と置き場所
`~/Desktop/tourism-analytics/` に、下記の名前で保存（**中身のシート名で自動判定するのでファイル名は目印**）：
| 系統 | ファイル名の例 |
|------|----------------|
| A 全国速報 | `202604_宿泊統計.xlsx` |
| B 都道府県確報 | `202603_宿泊統計.xlsx`（確報。速報と同月名になり得るので確報を優先保存） |
| C JNTO | `202604_JNTO.xlsx` / `202605_JNTO.xlsx` |
> Excel自体はリポジトリにコミットしない（入力素材。`git add` しない）。

## 3. 実行コマンド（系統別）
> どれも `~/Desktop/tourism-analytics/` で実行。まず `--deploy` なし＝ローカルのみ更新。

> 3系統とも `data.json` 更新＋`index.html` 再埋め込みまで自動。すべて push はしない（公開は5章）。

### A. 全国 月次速報
```bash
cd ~/Desktop/tourism-analytics
python3 update_monthly.py 202604_宿泊統計.xlsx
```

### B. 都道府県別 確報（案A：全国上書き＋都道府県47追加）
```bash
cd ~/Desktop/tourism-analytics
python3 update_data.py --kakuho 202603_宿泊統計.xlsx
```

### C. JNTO 月次（複数月まとめて指定可）
```bash
cd ~/Desktop/tourism-analytics
python3 update_data.py --jnto-monthly 202604_JNTO.xlsx 202605_JNTO.xlsx
# 1ファイルだけなら: python3 update_data.py --jnto-monthly 202604_JNTO.xlsx
```

> 内部ロジックは `update_data.py` の `parse_shukuhaku_kakuho` / `parse_jnto_monthly`。
> CLI は既存パーサを呼ぶ薄いラッパで、回帰テスト済み（pre状態から再構築して本番とバイト一致）。
> `--kakuho` と `--jnto-monthly` は1コマンドに併記も可能（例: `python3 update_data.py --kakuho 確報.xlsx --jnto-monthly J1.xlsx J2.xlsx`）。

## 4. 重要な仕様メモ（ハマりどころ）
- **案A採用**：確報(B)が出たら、その月は**全国も確報値で上書き**し、都道府県別と一致させる（都道府県47合計＝全国が成立する）。
- **確報/速報の混在は仕様**：確報を入れた月だけ全国が確報値、その前後の月は速報値のまま。系列上の段差は想定内。
- **JNTOは既存 `jc` の23件固定（ホワイトリスト）**：Excelに `北欧地域`/`中東地域`/`その他` があっても**追加しない**。逆に **ニュージーランドはこの月次様式に行が無く毎月欠測→スキップ**（歯抜けを許容。画面は `gv()` が null 返しで落ちない）。
- **当月値の列位置**（自動化済みだが確認用）：
  - JNTO：国名=`col2`、**当月値=`col5`（2026年）**、前年同月=`col4`。伸率は数式なので `(col5-col4)/col4*100` を自前計算。
  - 宿泊統計 確報：**第2表 `col1`=延べ(t)／`col19`=外国人(f)／日本人(j)=t−f**、**第8表 `col1`=客室稼働率(o)**。全国行=各表の7行目(r6, `令和N年M月`)。
  - 宿泊統計 速報(update_monthly.py)：第1表 r6 の A=日付／B=延べ／C=外国人、第5表 r6 B=稼働率。
- 県名キーは `　01北海道`→`01北海道` に全角空白除去して data.json キーと一致（47件一致確認済み）。全国は `全　国`（全角空白）。

## 5. 安全な更新フロー（毎回これに従う）
```bash
cd ~/Desktop/tourism-analytics
git checkout main && git pull                       # 最新化
git checkout -b feature/YYYYQn-update               # 作業ブランチ
cp data.json data.json.bak                          # 保険バックアップ

# 3章のA/B/Cを必要な分だけ実行（例。すべてローカルのみ更新、pushしない）
python3 update_monthly.py 202604_宿泊統計.xlsx                        # A 全国速報
python3 update_data.py --kakuho 202603_宿泊統計.xlsx                  # B 都道府県確報
python3 update_data.py --jnto-monthly 202604_JNTO.xlsx 202605_JNTO.xlsx  # C JNTO

open index.html                                     # ★ブラウザで目視確認★
#   - 全国系タブに最新月が出るか / 地域比較タブに都道府県別が出るか
#   - インバウンドタブにJNTO最新月が出るか / エラー・グラフ崩れが無いか

git add data.json index.html                        # 変更したファイルのみ（Excel/.bakは含めない）
git commit -m "YYYYMM更新: 内容を簡潔に"

git checkout main
git merge --squash feature/YYYYQn-update            # 1コミットに束ねる（reset不使用）
git commit -m "YYYYQn更新: 概要"
git push origin main                                # GitHub Pages へ公開

# 公開確認（数分後）
curl -s https://naito1014-ux.github.io/tourism-analytics/ | grep -c '"YYYYMM"'   # 最新月が反映されたか
```
- `feature` ブランチは消さずに残すと、やり直したいときの復元点になる。
- `data.json.bak` はローカルの最終保険（コミットしない）。
