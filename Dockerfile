# Chromiumと依存ライブラリが入った公式イメージ。自前で apt する地獄を避ける
FROM mcr.microsoft.com/playwright/python:v1.62.0-noble

# 日本語フォントは必須。無いと全部豆腐(□)になる。
# BIZ UDGothic は Windows 10 1809以降/11 に標準搭載 かつ SIL OFL なので、
# 描画側(ここ)と閲覧側(顧客のPowerPoint)で同じ物理フォントになる。
# これによりブラウザの実測値がPowerPointでもそのまま通用する。
# Noto CJK はフォールバック用（BIZ UDに無い字形の保険）。
# libreoffice-impress + poppler-utils はテンプレート登録時のプレビュー
# (pptx -> pdf -> png) に使う。HTMLプレビューではマスター側の装飾(ロゴ・バナー)が
# 見えないため、実物のラスタライズが必要。
# fonts-liberation はArial等の字幅互換代替(会社テンプレートのフォント対策)。
RUN apt-get update && apt-get install -y --no-install-recommends \
        fonts-morisawa-bizud-gothic fonts-noto-cjk fonts-liberation \
        libreoffice-impress poppler-utils \
    && rm -rf /var/lib/apt/lists/* && fc-cache -f

WORKDIR /app
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

# ベースイメージにChromiumは入っているが、pipで入れたplaywrightとの
# バージョン整合を取るため明示的に配置する
RUN playwright install chromium

COPY . .

ENV PYTHONUNBUFFERED=1 \
    PYTHONDONTWRITEBYTECODE=1 \
    EXTRACT_BACKEND=browser \
    MAX_PAGES=4 \
    PORT=8080

# 既定は HTTP サービス。Cloud Run Jobs では起動コマンドを上書きして使う:
#   gcloud run jobs deploy ... --command python --args job.py
# workers=1固定(ブラウザをプロセス内で共有するため。並列度はconcurrencyで調整)
CMD exec uvicorn app:app --host 0.0.0.0 --port ${PORT} --workers 1
