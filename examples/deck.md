# md2pptx 最小PoC
HTML/CSSをレイアウトエンジンとして使い、座標抽出してPPTX化する。
- Markdown を注釈付きHTMLに変換
- Chromium で折り返し計算と座標抽出
- python-pptx はEMU換算して置くだけ
!! UNIX哲学 / 3ツール + JSON
---
# 表はブラウザで実測する
| 指標 | Q1 | Q2 | Q3 |
| --- | --- | --- | --- |
| MRR | 1200 | 1480 | 1690 |
| 解約率 | 3 | 2 | 2 |
| 商談数 | 42 | 55 | 61 |
セル幅と行高はgetBoundingClientRectの結果をそのまま使う。
---
# グラフは数値だけを渡す
```chart
type: column
categories: [Q1, Q2, Q3, Q4]
series:
  - {name: MRR, values: [1200, 1480, 1690, 2050]}
  - {name: 目標, values: [1300, 1500, 1800, 2100]}
```
HTML側のcanvasはプレビュー専用。PPTXはネイティブグラフとして再描画される。
