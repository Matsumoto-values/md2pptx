# Chromiumと依存ライブラリが入った公式イメージ。自前で apt する地獄を避ける
FROM mcr.microsoft.com/playwright/python:v1.56.0-noble

# 日本語フォントは必須。無いと全部豆腐(□)になる
# ここで焼き込むフォントは「閲覧者のPCにも入っているもの」を選ぶこと
RUN apt-get update && apt-get install -y --no-install-recommends \
        fonts-noto-cjk fonts-noto-cjk-extra \
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
