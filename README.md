# md2pptx PoC — Markdown → 注釈付きHTML → 座標JSON → PPTX

3つの単機能ツールを JSON で繋ぐ構成。各段の出力が独立して検査できるのが利点。

```bash
python3 md2html.py    deck.md    > deck.html     # ① md → 注釈付きHTML(プレビュー兼レイアウト定義)
python3 extract.py    deck.html  > layout.json   # ② Chromiumで描画 → 座標・スタイル抽出
python3 build_pptx.py layout.json out.pptx       # ③ EMU換算して python-pptx で配置
```

サンドボックス等でブラウザが使えない場合は ② を `extract_stub.py` に差し替える
（CSS定数決め打ちの近似。折り返しの正確さは保証しない）。

## 変換コントラクト

| HTML注釈 | PPTX上の姿 |
|---|---|
| `<section class="slide">` | スライド1枚 |
| `data-pptx="title"` | ネイティブテキストボックス（太字・大） |
| `data-pptx="text"` | ネイティブテキストボックス |
| `data-pptx="list"` | 箇条書き段落（`a:buChar` を直挿し） |
| `data-pptx="table"` | ネイティブ表（セル座標・列幅をブラウザ実測） |
| `data-pptx="chart"` + `data-chart` | ネイティブグラフ（数値は属性のJSONが正） |
| `data-pptx="raster"` | 要素スクショを画像として貼る |
| 注釈なし | 無視 |

Markdown 側の記法: `# ` = タイトル / `- ` = 箇条書き / `!!` = raster /
`| a | b |` = 表 / ```` ```chart ```` = グラフ / `---` = スライド区切り。

## 表とグラフの設計思想（真逆）

|  | 座標の出どころ | データの出どころ |
|---|---|---|
| **表** | ブラウザ実測（セル単位） | DOMのテキスト |
| **グラフ** | ブラウザ実測（外枠のみ） | `data-chart` のJSON |

表はセルの折り返しをブラウザに計算させる価値が大きいので徹底実測。
グラフはHTML側の描画（canvas）はプレビュー専用の使い捨てで、PowerPointが自前で
再描画するため、渡すべきは絵ではなく数値。

### 表の実装上の注意
- `tbl.first_row = False` / `tbl.horz_banding = False` — テーマ由来の青い縞スタイルを殺す
- 行高は「最小値」でPowerPointが勝手に伸ばす。フォントメトリクス差の保険に +8%
- `cell.text = ...` は書式が飛ぶので `run.text` に代入する
- `cell.margin_*` にCSSのpaddingを移植しないと座標が数mmずれる
- **罫線は python-pptx にAPIが無い** → `tcPr` に `a:lnL/lnR/lnT/lnB` を直挿しする。
  `tcPr` の子要素順序は `ln* → fill → headers` で、破るとPowerPointが修復を要求する。
  実装では逆順に `tcPr.insert(0, ln)` して L,R,T,B の順を作っている
- 線幅の下限は 0.25pt（3175 EMU）。CSS 1px = 9525 EMU = 0.75pt で描画される
- 結合セルは左上基準で走査して `cell.merge()`（未実装）

### グラフの実装上の注意
- `plot.vary_by_categories = False` — 単系列が虹色になるのを防ぐ（最頻出の罠）
- 積み上げ棒で `OUTSIDE_END` は不正値。`INSIDE_END` にする（ファイルが壊れる）
- `legend.include_in_layout = False` — 凡例がプロット領域を侵食するのを防ぐ
- 散布図/バブルは `XyChartData` / `BubbleChartData`（別クラス）
- 第2軸・複合グラフは python-pptx が実質未対応 → スライドを分けるのが現実解
- Sankey・ネットワーク図はPowerPointに無いので `raster` に逃がす

### 配色の二重管理を避ける
HTMLプレビュー（CSS）とPPTXネイティブは別レンダラなので、放置すると体裁がズレる。
`tokens.json` を単一の正として、CSS変数と `build_pptx.py` の両方へ注入している。

```
tokens.json ─┬→ CSS変数（プレビュー）
             └→ palette / number_format（PPTX）
```

## 設計上のポイント

- **キャンバス固定**: 1280×720px を 13.333×7.5in にマップ。`px→EMU = Inches(13.333)/1280`
- **フォントサイズ**: CSS px × 0.75 = pt（96dpi → 72pt）
- **テキストボックスの内側余白を0に**: しないと座標が数mmずれる
- **折り返しはブラウザ担当**: `getBoundingClientRect()` の結果をそのまま使うので、
  python-pptx 側でテキスト溢れの計算をしなくて済む（これがこの方式の最大の旨味）
- **CSSで表現できないものは raster に逃がす**: グラデ・box-shadow・擬似要素・SVG

## 検証ループ

```bash
python3 /path/to/validate.py out.pptx        # OOXMLスキーマ・関係チェック
soffice --headless --convert-to pdf out.pptx
pdftoppm -jpeg -r 100 out.pdf slide          # 画像化してHTMLプレビューと見比べる
```

## 未実装（次にやるなら）

- 結合セル（colspan/rowspan は抽出済みだがビルダーが未対応）
- 散布図・複合グラフ
- テンプレート .potx のレイアウト継承（現状は blank layout に絶対配置）
- 段組み（flex/grid）は座標抽出後に絶対配置化されるので実は既に動く
- フォント実測差の吸収（ブラウザと PowerPoint のメトリクス差、幅に10%の余裕を持たせる）

---

# Cloud Run API として動かす

## 構成

このサービスに**LLMは入っていない**。入力が決まれば出力は一意に決まる決定論的な変換器。
非構造な入力から md を書く役目は別サービス（Agent SDK側）に分離する。

```
job.py          夜間バッチのエントリポイント（Cloud Run Jobs / シャーディング・リトライ）
app.py          FastAPI ラッパ（アドホック要求用 / min-instances=0前提）
renderer.py     変換パイプラインの統合層。job.py と app.py の共通コア
theme.py        デザイントークンの読み込みとマージ（DEFAULTS → themes/*.json → リクエスト上書き）
md2html.py      md → 注釈付きHTML          （to_html）
extract.py      HTML → 座標JSON            （async / BrowserPool）
extract_stub.py 同上のブラウザ無し近似版    （CI・開発用）
build_pptx.py   座標JSON → pptx            （build）
```

## エンドポイント

| メソッド | パス | 用途 |
|---|---|---|
| GET | `/healthz` | ブラウザ生存確認。死んだら503でインスタンスを作り直させる |
| GET | `/themes` | 利用可能テーマ一覧 |
| POST | `/preview` | HTMLをそのまま返す。人間のレビューとエージェントのQAループ用 |
| POST | `/layout` | 中間表現を返す。座標ズレのデバッグ用 |
| POST | `/render` | pptxを返す（`output:"url"` でGCS署名URL） |

```jsonc
// POST /render
{
  "markdown": "# タイトル\n- 箇条書き",
  "theme_id": "default",
  "theme": { "palette": ["FF6B00","004E98"], "table": { "zebra": false } },  // 部分上書き
  "filename": "deck.pptx",
  "output": "file"   // または "url"（GCS_BUCKET設定時）
}
```

## デプロイ（2形態・同一イメージ）

### ① 夜間バッチ → Cloud Run Jobs（主用途）

常駐しないので待機コストはゼロ。Chromiumの起動コスト(3〜5秒)は
1回の起動で全件に償却されるため、件数が多いほど無視できる比率になる。

```bash
gcloud run jobs deploy md2pptx-batch \
  --source . \
  --region asia-northeast1 \
  --command python --args job.py \
  --memory 2Gi --cpu 2 \
  --tasks 4 \                  # タスク並列。job.py が自前でシャーディングする
  --task-timeout 30m \
  --max-retries 2 \            # 非ゼロ終了したタスクだけ再実行される
  --set-env-vars MANIFEST=gs://my-bucket/nightly/manifest.jsonl,\
OUT_PREFIX=gs://my-bucket/decks/2026-08-04,CONCURRENCY=4

# 毎晨2時に起動
gcloud scheduler jobs create http md2pptx-nightly \
  --schedule "0 2 * * *" --time-zone Asia/Tokyo \
  --uri "https://asia-northeast1-run.googleapis.com/apis/run.googleapis.com/v1/namespaces/PROJECT/jobs/md2pptx-batch:run" \
  --http-method POST --oauth-service-account-email SA@PROJECT.iam.gserviceaccount.com
```

### ② アドホック要求 → Cloud Run Service（任意・min-instances=0）

```bash
gcloud run deploy md2pptx \
  --source . \
  --region asia-northeast1 \
  --memory 2Gi --cpu 2 \
  --concurrency 4 \            # MAX_PAGES と揃える
  --min-instances 0 \          # 待機コストを払わない
  --cpu-boost \                # コールドスタートを短縮
  --timeout 300 \
  --no-allow-unauthenticated
```

## なぜ Service常駐ではなく Jobs なのか

| | Cloud Run Service (min-instances=1) | Cloud Run Jobs |
|---|---|---|
| 待機コスト | **常時課金される** | ゼロ |
| CPU割り当て | リクエスト処理中のみ（設定次第） | 実行中は常時 |
| リトライ | 自前で実装 | `--max-retries` で標準装備 |
| タスク並列 | 自前でスケール制御 | `--tasks` で分割 |
| 実行時間 | リクエストタイムアウト（最大60分）に縛られる | `--task-timeout` で独立 |
| コールドスタート | 避けたい | **N件で割れるので気にしない** |

夜間バッチで非同期・コールドスタート許容なら、Jobs 以外を選ぶ理由がない。

### コールドスタートの実体はイメージPull
Chromium入りイメージは2GB前後あり、起動の大半はプロセス起動ではなくPull時間。
これが気になる場合の順序は:
1. `--tasks` を増やして1タスクあたりの件数を減らす（Pullは並列に走る）
2. ベースを `python:3.12-slim` + `playwright install --with-deps chromium` に置き換えて痩せさせる
3. それでも足りなければ mdごとにJobを起動するのをやめて、1回の起動で束ねる件数を増やす

## Cloud Run 固有の注意

- **`--no-sandbox --disable-dev-shm-usage`**: gVisor環境では必須（`extract.py` の `LAUNCH_ARGS`）
- **`--workers 1` 固定**: ブラウザをプロセス内で共有するため。並列度はCloud Run側の
  concurrency で調整する。workers を増やすとブラウザがワーカー数分立ち上がってメモリが死ぬ
- **フォント**: `fonts-noto-cjk` をイメージに焼く。無いと日本語が全部豆腐になる。
  ただし**閲覧者のPCに無いフォントは開いた時に崩れる**ので、
  B2B用途では游ゴシック・Meiryo など標準搭載フォントに寄せるのが安全
- **バッチ側は枚数制限を気にしなくてよい**（`--task-timeout` が効くだけ）。
  Service側の `MAX_SLIDES` はリクエストタイムアウト対策
- **CPUバウンド処理**（python-pptx）は `run_in_threadpool` に投げてイベントループを塞がない

## バッチの入力形式

`MANIFEST` は JSONL。1行1デッキ。

```jsonc
{"id":"acme","markdown":"# ...","filename":"acme.pptx","theme":{"palette":["FF6B00"]}}
{"id":"globex","markdown_uri":"gs://bucket/src/globex.md","theme_id":"globex"}
```

- **1件の失敗は他を巻き込まない**（`asyncio.gather(return_exceptions=True)` で集計）
- 失敗が1件でもあればプロセスは非ゼロで終了し、Cloud Run Jobs がそのタスクを再実行する
- ログは Cloud Logging 向けに `{"severity":..., "message":...}` のJSON1行形式

## ローカル検証

```bash
EXTRACT_BACKEND=stub python3 smoke_test.py        # ブラウザ無しでAPI経路を通す
EXTRACT_BACKEND=stub MANIFEST=manifest.jsonl OUT_PREFIX=./batch-out python3 job.py
EXTRACT_BACKEND=stub uvicorn app:app --reload     # 手で叩く
docker build -t md2pptx . && docker run -p 8080:8080 md2pptx   # 本番同等
docker run --rm -e MANIFEST=... -e OUT_PREFIX=... md2pptx python job.py   # バッチ同等
```

`EXTRACT_BACKEND=stub` はCSS定数決め打ちの近似レイアウトなので、
**配線の確認には使えるが見た目の検証には使えない**。見た目は必ず browser バックエンドで確認する。
