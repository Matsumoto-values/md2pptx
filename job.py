#!/usr/bin/env python3
"""job.py -- 夜間バッチ用エントリポイント (Cloud Run Jobs).

起動 -> 全件処理 -> 終了。常駐しないので待機コストはゼロ。
Chromiumの起動コスト(3〜5秒)は1回だけ払い、全件で使い回して償却する。

環境変数:
  MANIFEST         入力JSONL。gs://... またはローカルパス
  OUT_PREFIX       出力先。gs://bucket/prefix またはローカルディレクトリ
  EXTRACT_BACKEND  browser(既定) | stub
  MAX_PAGES        ブラウザ内の同時ページ数 (既定4)
  CONCURRENCY      同時に処理する件数 (既定4)
  CLOUD_RUN_TASK_INDEX / CLOUD_RUN_TASK_COUNT  Cloud Runが注入。自動でシャーディング

manifest(JSONL)の1行:
  {"id":"acme-2026-08","markdown":"# ...","theme_id":"default",
   "theme":{"palette":["FF6B00"]},"filename":"acme.pptx"}
  markdown の代わりに "markdown_uri":"gs://..." でも可
"""
import asyncio
import json
import logging
import os
import pathlib
import sys
import time

import theme as theme_mod
from renderer import Renderer

MANIFEST = os.getenv("MANIFEST", "manifest.jsonl")
OUT_PREFIX = os.getenv("OUT_PREFIX", "./out")
BACKEND = os.getenv("EXTRACT_BACKEND", "browser")
MAX_PAGES = int(os.getenv("MAX_PAGES", "4"))
CONCURRENCY = int(os.getenv("CONCURRENCY", "4"))
TASK_INDEX = int(os.getenv("CLOUD_RUN_TASK_INDEX", "0"))
TASK_COUNT = int(os.getenv("CLOUD_RUN_TASK_COUNT", "1"))

logging.basicConfig(
    level=logging.INFO,
    format='{"severity":"%(levelname)s","message":"%(message)s"}')  # Cloud Logging向け
log = logging.getLogger("job")

_gcs = None


def gcs():
    global _gcs
    if _gcs is None:
        from google.cloud import storage
        _gcs = storage.Client()
    return _gcs


def read(uri):
    if uri.startswith("gs://"):
        bucket, _, name = uri[5:].partition("/")
        return gcs().bucket(bucket).blob(name).download_as_bytes()
    return pathlib.Path(uri).read_bytes()


def write(uri, data, content_type=None):
    if uri.startswith("gs://"):
        bucket, _, name = uri[5:].partition("/")
        gcs().bucket(bucket).blob(name).upload_from_string(
            data, content_type=content_type)
    else:
        p = pathlib.Path(uri)
        p.parent.mkdir(parents=True, exist_ok=True)
        p.write_bytes(data)
    return uri


def load_manifest():
    lines = read(MANIFEST).decode("utf-8").splitlines()
    items = [json.loads(l) for l in lines if l.strip()]
    # Cloud Run Jobs のタスク並列に合わせて自前でシャーディングする
    return [it for i, it in enumerate(items) if i % TASK_COUNT == TASK_INDEX]


async def one(renderer, item, sem):
    async with sem:
        t0 = time.monotonic()
        ident = item.get("id") or item.get("filename") or "unnamed"
        md = item.get("markdown")
        if md is None:
            md = read(item["markdown_uri"]).decode("utf-8")
        T = theme_mod.load(item.get("theme_id", "default"), item.get("theme"))
        data, n = await renderer.render(md, T)
        name = item.get("filename") or f"{ident}.pptx"
        uri = write(f"{OUT_PREFIX.rstrip('/')}/{name}", data,
                    "application/vnd.openxmlformats-officedocument"
                    ".presentationml.presentation")
        log.info("done id=%s slides=%d bytes=%d ms=%d -> %s", ident, n, len(data),
                 (time.monotonic() - t0) * 1000, uri)
        return ident


async def main():
    items = load_manifest()
    log.info("start task=%d/%d items=%d backend=%s", TASK_INDEX, TASK_COUNT,
             len(items), BACKEND)
    if not items:
        return 0

    t0 = time.monotonic()
    sem = asyncio.Semaphore(CONCURRENCY)
    async with Renderer(BACKEND, MAX_PAGES) as r:
        # 1件の失敗で全体を落とさない。return_exceptionsで拾って最後に集計する
        results = await asyncio.gather(
            *(one(r, it, sem) for it in items), return_exceptions=True)

    failed = []
    for item, res in zip(items, results):
        if isinstance(res, Exception):
            ident = item.get("id") or item.get("filename") or "unnamed"
            failed.append(ident)
            log.error("failed id=%s %s: %s", ident, type(res).__name__, res)

    log.info("finish ok=%d failed=%d elapsed=%.1fs",
             len(items) - len(failed), len(failed), time.monotonic() - t0)
    # 非ゼロで終了するとCloud Run Jobsがタスクをリトライする
    return 1 if failed else 0


if __name__ == "__main__":
    sys.exit(asyncio.run(main()))
