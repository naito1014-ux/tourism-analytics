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
