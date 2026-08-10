#!/usr/bin/env python3
"""app.py -- md -> pptx 変換API (Cloud Run Service向け / アドホック要求用).

夜間バッチは job.py + Cloud Run Jobs を使う。こちらは画面から即時生成したい時用。
min-instances=0 前提。コールドスタート(chromium起動3〜5秒)を許容する代わりに
アイドルコストをゼロにする。

このサービスにLLMは入っていない。入力が決まれば出力は一意に決まる決定論的な変換器。
非構造な入力から md を書く役目は別サービス(Agent SDK側)の仕事。

環境変数:
  EXTRACT_BACKEND  browser(既定) | stub      stubはブラウザ無し環境用の近似
  MAX_PAGES        同時に開くページ数 (既定4)。Cloud Runのconcurrencyと揃える
  MAX_SLIDES       1リクエストの上限枚数 (既定60)
  GCS_BUCKET       指定時は成果物をGCSへ置いて署名URLを返す
  SIGNED_URL_TTL   署名URLの有効秒数 (既定3600)
"""
import io
import logging
import os
import time
import uuid
from contextlib import asynccontextmanager

from fastapi import FastAPI, File, Form, HTTPException, UploadFile
from fastapi.concurrency import run_in_threadpool
from fastapi.responses import HTMLResponse, JSONResponse, StreamingResponse
from pydantic import BaseModel, Field

import md2html
import template_store
import theme as theme_mod
from renderer import Renderer
from template_register import register
from tools.pptx2base import RejectedFile

BACKEND = os.getenv("EXTRACT_BACKEND", "browser")
MAX_PAGES = int(os.getenv("MAX_PAGES", "4"))
MAX_SLIDES = int(os.getenv("MAX_SLIDES", "60"))
GCS_BUCKET = os.getenv("GCS_BUCKET")
SIGNED_URL_TTL = int(os.getenv("SIGNED_URL_TTL", "3600"))
PPTX_MIME = ("application/vnd.openxmlformats-officedocument"
             ".presentationml.presentation")

MAX_UPLOAD = int(os.getenv("MAX_UPLOAD_BYTES", str(80 * 1024 * 1024)))

logging.basicConfig(level=logging.INFO, format="%(levelname)s %(message)s")
log = logging.getLogger("md2pptx")
renderer = None
store = template_store.from_env(os.environ)


@asynccontextmanager
async def lifespan(app):
    """インスタンスの生存期間だけブラウザを持つ。min-instances=0 前提なので
    アイドル時のコストはゼロ。起動コストはそのインスタンスが処理する件数で割られる."""
    global renderer
    async with Renderer(BACKEND, MAX_PAGES) as r:
        renderer = r
        log.info("renderer ready backend=%s max_pages=%d", BACKEND, MAX_PAGES)
        yield
    renderer = None


app = FastAPI(title="md2pptx", version="0.1.0", lifespan=lifespan)


class RenderRequest(BaseModel):
    markdown: str = Field(min_length=1, max_length=500_000)
    theme_id: str = "default"
    theme: dict | None = None          # テーマの部分上書き(色だけ変える等)
    filename: str = "deck.pptx"
    output: str = "file"               # file | url
    # 登録済みテンプレートを使う場合。theme_id より優先される
    tenant_id: str | None = None
    template_id: str | None = None
    template_version: int | str | None = None   # 既定は最新のpublished


def resolve_theme(req):
    """テンプレート指定があれば保管層から、無ければ themes/*.json から解決する."""
    if req.template_id:
        if not req.tenant_id:
            raise HTTPException(400, "tenant_id is required with template_id")
        try:
            rec = store.load(req.tenant_id, req.template_id, req.template_version)
        except template_store.TemplateNotFound as e:
            raise HTTPException(404, str(e))
        # ベースpptxをローカルへ展開し、テーマの base_pptx を実体に向ける
        base_theme = store.materialize(rec)
        return theme_mod.normalize(theme_mod.deep_merge(
            theme_mod.deep_merge(theme_mod.DEFAULTS, base_theme), req.theme))
    try:
        return theme_mod.load(req.theme_id, req.theme)
    except FileNotFoundError as e:
        raise HTTPException(400, str(e))


def require_renderer():
    if not (renderer and renderer.alive):
        raise HTTPException(503, "renderer unavailable")
    return renderer


def upload(blob_name, data):
    """成果物はレスポンスに載せずGCSへ。数MB×並列はメモリを食うため."""
    from datetime import timedelta

    from google.cloud import storage
    bucket = storage.Client().bucket(GCS_BUCKET)
    blob = bucket.blob(blob_name)
    blob.upload_from_string(data, content_type=PPTX_MIME)
    return blob.generate_signed_url(expiration=timedelta(seconds=SIGNED_URL_TTL),
                                    version="v4")


@app.get("/healthz")
async def healthz():
    ok = bool(renderer and renderer.alive)
    # ブラウザが死んだらインスタンスを落として作り直させる
    return JSONResponse({"ok": ok, "backend": BACKEND}, status_code=200 if ok else 503)


@app.post("/tenants/{tenant_id}/templates")
async def upload_template(tenant_id: str, template_id: str = Form(...),
                          name: str = Form(None), file: UploadFile = File(...),
                          keep_layout_only: bool = Form(False)):
    """会社テンプレートを登録する（draftとして保存）.

    アップロードは信用できない入力なので受入検査を通す。抽出は元スライドから行い、
    保存するベースはスライドを全削除したもの（顧客の機密を保存しないため）。
    """
    data = await file.read()
    if len(data) > MAX_UPLOAD:
        raise HTTPException(413, f"file too large (max {MAX_UPLOAD} bytes)")
    try:
        _, report = await run_in_threadpool(
            register, data, tenant_id, template_id, store,
            name=name, keep_layout_only=keep_layout_only)
    except RejectedFile as e:
        raise HTTPException(400, str(e))
    log.info("template registered tenant=%s id=%s v=%s font=%s", tenant_id,
             template_id, report["version"], report["font"]["status"])
    return report


@app.get("/tenants/{tenant_id}/templates")
async def list_templates(tenant_id: str):
    return {"templates": await run_in_threadpool(store.list, tenant_id)}


@app.get("/tenants/{tenant_id}/templates/{template_id}")
async def get_template(tenant_id: str, template_id: str, version: str | None = None):
    try:
        rec = await run_in_threadpool(store.load, tenant_id, template_id, version)
    except template_store.TemplateNotFound as e:
        raise HTTPException(404, str(e))
    rec.pop("base")                     # 実体は返さない
    return rec


@app.post("/tenants/{tenant_id}/templates/{template_id}/publish")
async def publish_template(tenant_id: str, template_id: str, version: int):
    """draft を published にする。発行済みバージョンは書き換えない（再現性のため）."""
    try:
        return await run_in_threadpool(store.publish, tenant_id, template_id, version)
    except template_store.TemplateNotFound as e:
        raise HTTPException(404, str(e))


@app.get("/themes")
async def themes():
    return {"themes": theme_mod.available()}


@app.post("/preview", response_class=HTMLResponse)
async def preview(req: RenderRequest):
    """HTMLをそのまま返す。人間のレビューとエージェントのQAループ用."""
    return md2html.to_html(req.markdown, resolve_theme(req))


@app.post("/layout")
async def layout(req: RenderRequest):
    """中間表現を返す。座標のズレを追う時のデバッグ用."""
    T = resolve_theme(req)
    _, lay = await require_renderer().layout(req.markdown, T)
    return lay


@app.post("/render")
async def render(req: RenderRequest):
    rid, t0 = uuid.uuid4().hex[:8], time.monotonic()
    T = resolve_theme(req)

    r = require_renderer()
    if r.count_slides(req.markdown, T) > MAX_SLIDES:
        raise HTTPException(413, f"too many slides (max {MAX_SLIDES})")

    data, n = await r.render(req.markdown, T)
    log.info("rid=%s slides=%d bytes=%d ms=%d", rid, n, len(data),
             (time.monotonic() - t0) * 1000)

    if req.output == "url":
        if not GCS_BUCKET:
            raise HTTPException(400, "GCS_BUCKET is not configured")
        url = await run_in_threadpool(upload, f"{rid}/{req.filename}", data)
        return {"url": url, "slides": n, "bytes": len(data)}

    return StreamingResponse(
        io.BytesIO(data), media_type=PPTX_MIME,
        headers={"Content-Disposition": f'attachment; filename="{req.filename}"',
                 "X-Request-Id": rid, "X-Slide-Count": str(n)})
