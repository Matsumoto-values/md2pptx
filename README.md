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
`| a | b |` = 表 / ```` ```chart ```` = グラフ / ```` ```flow ```` = フローチャート /
`---` = スライド区切り。

### フローチャート記法

````markdown
```flow
direction: vertical        # 省略時は横並び
input: mdを受け取る
?レイアウト検証OK?          # 行頭"?"で菱形(判断)
output: pptxを出力
```
````

1行1ボックス。`見出し: 本文` と書くと見出しが小さい太字で乗る。
生成されるのは**ネイティブのオートシェイプと、図形に接続されたコネクタ**なので、
PowerPoint上でボックスをドラッグすると矢印が追従する。

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

## 会社テンプレートを適用する

既存の会社資料をベースに開き、マスター/レイアウトの装飾・ロゴ・フッタを継承する。
配色・フォント・余白はテンプレートから実測抽出する。

```bash
cp /path/to/corp.pptx templates/corp.pptx
python3 tools/pptx2theme.py templates/corp.pptx > themes/corp.json
# 出力をレビューして調整（特に font がLinux側で解決できるか）

python3 md2html.py deck.md corp > deck.html
python3 extract.py deck.html corp > layout.json     # テーマ指定を忘れないこと
python3 build_pptx.py layout.json out.pptx corp
```

API/バッチからは `theme_id: "corp"` を指定するだけ。

### 継承されるもの / されないもの

| | 状況 |
|---|---|
| 背景・ロゴ・装飾・フッタ・ページ番号 | ○ マスター/レイアウトから自動で乗る |
| スライドサイズ | ○ テンプレート側を正とする |
| 配色・フォント・余白 | △ 実測抽出してテーマJSONにする（人間のレビュー前提） |
| プレースホルダ | 削除する（全要素を絶対配置するため。`strip_placeholders`） |
| アウトライン表示・「リセット」 | × 絶対配置なのでプレースホルダを使わない |

### 抽出の設計（実資料3件で検証済み）

- **テーマ部品は読まない**。3件すべて Office 既定のまま放置されており、
  ブランドデザインは図形の直接書式に入っていた
- **配色はデザイン層（マスター＋使用中レイアウト）を優先**。スライド側から採ると
  注釈の赤やコンテンツの色を拾う
- **レイアウトは実スライドの使用頻度で選ぶ**。プレースホルダ数や並び順で選ぶと
  セクション表紙が当たる
- **全マスターを走査する**。`prs.slide_layouts` は先頭マスターの分しか返さない
- **フォントはCJK比率で判定**。`run.font.name` はラテン用指定を返すため、
  日本語資料でも `Arial` 等が最頻出になる

### フォントが用意できない場合

字幅互換の代替があれば `font_stack` に入れて計測に使い、`font`（pptxに書く名前）は
元のフォント名のままにする。これで**手元にフォントが無くても座標が信用できる**。

| テンプレートの指定 | 計測に使う代替 | 必要なパッケージ |
|---|---|---|
| Arial / Helvetica | Liberation Sans | `fonts-liberation` |
| Times New Roman | Liberation Serif | `fonts-liberation` |
| Calibri | Carlito | `fonts-crosextra-carlito` |
| **日本語フォント全般** | **互換なし** | — |

日本語フォントには字幅互換の代替が存在しない。`Meiryo UI` 等が指定されている場合、
`extract.py` が警告を出すので、フォントを導入するかテーマの `font` を
BIZ UDPGothic に変えるかを人間が判断する。

**顧客資料と、そこから生成したテーマはリポジトリに含めない**（`.gitignore` 済み）。
匿名サンプルは `themes/example-corporate.json`。

## テンプレートの登録（マルチテナント）

会社ごとにテンプレートを登録し、後日それを使って何度でも生成する運用向け。
**登録は稀・対話的・人間のレビューあり**、**生成は頻繁・自動**。性質が違うので処理を分ける。

```
POST /tenants/{tenant_id}/templates            登録（draftとして保存）
GET  /tenants/{tenant_id}/templates            一覧
GET  /tenants/{tenant_id}/templates/{id}       取得（?version=1 で固定）
POST /tenants/{tenant_id}/templates/{id}/publish?version=1   発行
POST /render {markdown, tenant_id, template_id[, template_version]}
```

```bash
curl -X POST http://localhost:8080/tenants/acme/templates \
  -F template_id=corp -F name=コーポレート -F file=@corp.pptx
```

### 登録時の処理

1. **受入検査**（サイズ・zip爆弾・マクロ・IDの形式）— 顧客のアップロードは信用できない入力
2. **元スライドからデザイントークンを実測抽出** — 色・フォント・余白はスライドから測るので
   ベース化より必ず先に行う
3. **ベース化: スライドを全削除** — マスターとレイアウトだけ残せばデザインは完全に再現できる。
   **顧客の機密スライドを保存しない**（実測で64〜96%の容量削減にもなる）
4. **フォント解決の判定を確定させる** — SaaSでは利用者がサーバーにフォントを入れられない。
   生成のたびに警告が出る状態にせず、登録時に人間へ提示する
5. draft として保存

### フォント判定の4状態

| status | 意味 | 対応 |
|---|---|---|
| `ok` | 宣言フォントがそのまま使える | そのまま |
| `substituted` | 字幅互換の代替で計測（Arial→Liberation Sans等） | そのまま。座標は信用できる |
| `approximated` | 代替はあるが字幅が違う | 折り返しがズレ得る。要レビュー |
| `unavailable` | 候補が全滅 | テーマの `font` を変更するか、フォントを導入する |

### バージョニング

**発行済みバージョンは書き換えない。** 過去に生成した資料を再現できるようにするため。
再登録は常に新しいバージョンになり、`template_version` で固定指定できる。

### 登録内容の確認（プレビュー）

```
GET  /tenants/{t}/templates/{id}/preview?version=1          JSON（data URLの配列）
GET  /tenants/{t}/templates/{id}/preview?version=1&page=1    単一PNG（imgタグ用）
POST /tenants/{t}/templates/{id}/preview/html                HTML
```

**HTMLプレビューだけでは不十分**。ロゴやバナーはベースpptxのマスター側にあり
HTMLには存在しないため、「ロゴと本文が一緒にどう見えるか」「本文が装飾に
被っていないか」は実物のpptxをラスタライズしないと確認できない。
そのため `libreoffice-impress` と `poppler-utils` をイメージに入れている
（pptx → pdf → png）。

- サンプルデッキは3枚（タイトル/本文/箇条書き/表/グラフ/フロー図を網羅）
- 所要時間の実測は **2.7〜3.6秒**（3枚・110dpi）。対話的な確認としては許容範囲
- **ベースの内容とテーマの内容でキャッシュする**。テーマを編集して再確認する
  使い方をするため、テーマもキーに含める（キャッシュヒットは即時）
- LibreOfficeが無い環境では **503** を返す。UI側はHTMLプレビューへ退避する

注意: soffice はpptx全体をPDFに変換するため、枚数が多いと遅い
（67枚の資料で23.5秒）。プレビューが3枚固定なのはこのため。

### ロゴ・画像・フォントの扱い

スライドを削除すると、python-pptx はパートグラフを辿って保存するため
**到達不能になった画像は自動的に破棄される**。つまり:

| 参照元 | ベースに残るか | 生成物での扱い |
|---|---|---|
| マスター/レイアウトが参照する画像（ロゴ・透かし・背景） | **残る** | 全スライドに自動で乗る |
| スライドだけが参照する画像（本文中の図・スクショ） | **破棄される** | 乗らない（顧客の中身なので当然） |
| 埋め込みフォント（`ppt/fonts/*.fntdata`） | 残る | 閲覧側の再現性に寄与する |

実測（7件）: アップロード 0.28〜33.9MB → ベース 0.05〜18.0MB。
破棄されるのはほぼスライド本文の画像。

**ロゴはベースpptxの中に埋め込まれた状態で保管される。** 画像を別管理せず、
生成のたびに貼り直すこともしない。テンプレートを1回登録すれば、以後の生成では
マスター/レイアウト経由で自動的に乗る。

**注意すべきケース**: ロゴが各スライドに直接置かれているだけの資料では、
ベースに何も残らず装飾が再現されない（配色・フォント・余白だけが効く）。
登録時に `design_warning` で通知する。実測7件のうち1件がこのケースだった。

**容量の実態**: 重いのは必ずしも画像ではない。ある資料ではベース11.35MBのうち
**10.53MBが埋め込みフォント22個**（`NotoSansJP-bold.fntdata` だけで2.5MB）で、
画像は0.64MBしかなかった。別の資料では逆に18.05MBすべてが画像
（`image6.png` が6.07MB）。登録レポートの `assets.largest` で内訳を返す。

**重複排除**: ベースの実体は内容ハッシュ（sha256）で1本化して保管し、
メタが参照する。テンプレートを再登録しても中身が同じなら容量は増えない
（実測: 同一テンプレートを3バージョン登録して 34MB → 11.35MB）。
そのため `pptx2base` は zip のタイムスタンプを固定して
**同じ入力なら常に同じバイト列**になるようにしている（現在時刻が入ると
秒をまたぐだけでハッシュが変わり、重複排除が効かない）。

**埋め込みフォントは計測には使えない**。`.fntdata` は難読化されていて
TrueType署名を持たないため、そのままブラウザに読ませることはできない。
描画側のフォント解決は別問題として `font_status` で扱う。

### 保管層の差し替え

`TemplateStore` インターフェースに `LocalTemplateStore`（開発用）と
`GcsTemplateStore` の2実装がある。`TEMPLATE_BUCKET` があればGCS、無ければローカル。

millet_ms に組み込む際は、**メタデータをCloudSQL/Firestoreに移した実装を注入する**のが素直
（一覧・検索・監査・会社テーブルとのFKが必要になるため）。ベースpptxはGCSのまま。
`TemplateStore` を実装したクラスを差し込めばよく、他のモジュールは変更不要。

ベースpptxは `materialize()` がプロセス内キャッシュ（内容ハッシュ）に展開する。
バッチで同じテンプレートを何度も使うため、ダウンロードは1回で済む。

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

### GitHub → Cloud Build で CI/CD する（推奨）

`cloudbuild.yaml` が `ビルド → 構造検証 → push → Jobs更新` を実行する。
**検証に失敗したらデプロイしない**。

#### 初回だけ必要な準備

```bash
PROJECT_ID=$(gcloud config get-value project)
REGION=asia-northeast1
BUCKET=YOUR_BUCKET

# 1. Artifact Registry
gcloud artifacts repositories create md2pptx \
  --repository-format=docker --location=$REGION

# 2. サービスアカウント（ビルドとジョブ実行を兼用する1つ）
gcloud iam service-accounts create md2pptx --display-name="md2pptx"
SA=md2pptx@$PROJECT_ID.iam.gserviceaccount.com

for ROLE in roles/logging.logWriter \
            roles/artifactregistry.writer \
            roles/run.admin; do
  gcloud projects add-iam-policy-binding $PROJECT_ID \
    --member="serviceAccount:$SA" --role="$ROLE"
done

# ジョブに自分自身をSAとして設定するため、自分に対する actAs 権限が必要。
# プロジェクト全体ではなくこのSA自身にだけ付けて範囲を絞る
gcloud iam service-accounts add-iam-policy-binding $SA \
  --member="serviceAccount:$SA" --role=roles/iam.serviceAccountUser

# GCSはバケット単位で付与（プロジェクト全体のストレージ権限は与えない）
gcloud storage buckets add-iam-policy-binding gs://$BUCKET \
  --member="serviceAccount:$SA" --role=roles/storage.objectAdmin

# 3. GitHubリポジトリを接続（初回はブラウザ認可が必要）
gcloud builds connections create github md2pptx-gh --region=$REGION
gcloud builds repositories create md2pptx \
  --remote-uri=https://github.com/Matsumoto-values/md2pptx.git \
  --connection=md2pptx-gh --region=$REGION

# 4. トリガー作成（--service-account の指定は必須）
gcloud builds triggers create github \
  --name=md2pptx-main \
  --region=$REGION \
  --repository=projects/$PROJECT_ID/locations/$REGION/connections/md2pptx-gh/repositories/md2pptx \
  --branch-pattern='^main$' \
  --build-config=cloudbuild.yaml \
  --service-account=projects/$PROJECT_ID/serviceAccounts/$SA \
  --substitutions=_MANIFEST=gs://$BUCKET/nightly/manifest.jsonl,_OUT_PREFIX=gs://$BUCKET/decks,_RUNTIME_SA=$SA
```

**`--service-account` を省略できない理由**: Cloud Build の既定SAが
Compute Engine の既定SAに変わっており、そこには Cloud Build Service Account
ロールが含まれていない。レガシーSAが有効な古いプロジェクトを除き、
トリガー作成/更新時にSAの明示指定が必要。
またユーザー指定SAで実行する場合、`options.logging: CLOUD_LOGGING_ONLY` が必須。

**1つのSAを兼用する場合の権限集中について**: このSAはビルド時の権限
（イメージのpush、Jobsのデプロイ）と実行時の権限（GCSの読み書き）の両方を持つ。
つまりジョブ側から自分を再デプロイすることも技術的には可能になる。
加えて Cloud Build トリガーの性質上、**リポジトリへの書き込み権限を持つ人は
このSAの権限で任意のコードを実行できる**（pushでトリガーが起動するため）。
権限を分離したい場合は、ビルド用SA（`artifactregistry.writer` / `run.admin` /
`logging.logWriter` / 対象SAへの `iam.serviceAccountUser`）と
ジョブ実行用SA（対象バケットへの `storage.objectAdmin` のみ）に分け、
トリガーの `_RUNTIME_SA` に後者を渡す。

#### CIが検証する内容（`tests/pptx_check.py`）

**イメージの中で** ブラウザバックエンドを使って実行するため、
コンテナ内のフォント解決も同時に検証される。

- フォントが解決できているか（できていなければ座標が全部信用できない）
- 全要素がスライド内に収まっているか（溢れ検出）
- 罫線の `tcPr` 子要素順序が `ln* → fill` になっているか
- コネクタが図形に接続されているか（`stCxn` / `endCxn`）
- 矢印の先端があるか、`a:ln` の子要素順序が正しいか
- 表・グラフの数、系列数、`vary_by_categories=False`

**XMLの要素順序違反は PowerPoint の「修復が必要です」に直結し、
顧客の信頼を直接損なう**。人間のレビューでは絶対に検出できないので、
ここを機械のゲートにしておく。

#### 手元での実行

```bash
uv run python tests/pptx_check.py            # 構造検証（要ブラウザ）
uv run python tests/smoke_test.py            # API配線（ブラウザ不要）
uv run python tests/template_flow_test.py    # 登録→発行→生成（ブラウザ不要）
uv run python tests/template_flow_test.py templates/corp.pptx   # 実資料で確認
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
- **フォント**: `fonts-morisawa-bizud-gothic` をイメージに焼く。
  BIZ UDGothic は Windows 標準搭載かつ SIL OFL なので、
  **描画側（Cloud Run）と閲覧側（顧客のPowerPoint）で同じ物理フォント**になり、
  ブラウザの実測値がそのまま通用する。
  ローカル検証時は `sudo apt install fonts-morisawa-bizud-gothic`。
  フォントが解決できない場合は `extract.py` が警告を出す（黙って座標が狂うのを防ぐため）
- **バッチ側は枚数制限を気にしなくてよい**（`--task-timeout` が効くだけ）。
  Service側の `MAX_SLIDES` はリクエストタイムアウト対策
- **CPUバウンド処理**（python-pptx）は `run_in_threadpool` に投げてイベントループを塞がない

## バッチの入力形式

`MANIFEST` は JSONL。1行1デッキ。

```jsonc
{"id":"acme","markdown":"# ...","filename":"acme.pptx","theme":{"palette":["FF6B00"]}}
{"id":"globex","markdown_uri":"gs://bucket/src/globex.md","theme_id":"globex"}
```

`examples/manifest.jsonl` は全件成功する正常系。
`examples/manifest-failures.jsonl` はデータ起因の失敗を混ぜた異常系テスト用。

- **1件の失敗は他を巻き込まない**（`asyncio.gather(return_exceptions=True)` で集計）
- ログは Cloud Logging 向けに `{"severity":..., "message":...}` のJSON1行形式

### 失敗の分類と終了コード

| 種別 | 例 | 終了コード | Cloud Run Jobs の挙動 |
|---|---|---|---|
| **データ起因** | テーマ名の誤り、参照先mdの不在、```chart のYAML構文エラー | **0** | リトライしない |
| **環境起因** | ブラウザの異常終了、GCSの書き込み失敗、OOM | **1** | `--max-retries` に従い再実行 |

リトライしても直らない失敗を非ゼロ終了で扱うと、**絶対に成功しない処理を
繰り返して実行全体が失敗扱いになる**（実測: 5件中1件のデータ不備で、
成功分が25秒で終わっているのに全体が73秒に延び、うち48秒が無駄なリトライだった）。

判定できない例外は **env に倒す**。環境起因を誤って data と判定すると、
デッキが黙って生成されないまま正常終了してしまうため。

失敗は `OUT_PREFIX/_failures/<task_index>.jsonl` に記録される。
監視はこのファイルか成果物の件数で組む。

```jsonc
{"id":"acme","kind":"data","error":"DataError","message":"theme: unknown theme: xxx"}
```

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
