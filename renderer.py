#!/usr/bin/env python3
"""renderer.py -- 変換パイプラインの統合層.

HTTPサービス(app.py)とバッチジョブ(job.py)の両方がここを呼ぶ。
ブラウザの寿命だけが両者で違う:
  - HTTP : リクエストを跨いで使い回す
  - バッチ: 1回の起動で全件使い回して終了時に閉じる（コールドスタートをN件で割る）
"""
import io

import build_pptx
import md2html
from extract import BrowserPool
from extract_stub import extract_stub


class Renderer:
    def __init__(self, backend="browser", max_pages=4):
        self.backend = backend
        self.pool = BrowserPool(max_pages=max_pages) if backend == "browser" else None

    async def __aenter__(self):
        if self.pool:
            await self.pool.start()
        return self

    async def __aexit__(self, *exc):
        if self.pool:
            await self.pool.stop()

    @property
    def alive(self):
        return self.backend == "stub" or bool(self.pool and self.pool.alive)

    async def layout(self, markdown, T):
        html = md2html.to_html(markdown, T)
        canvas = (T["canvas"]["w"], T["canvas"]["h"])
        if self.pool:
            return html, await self.pool.extract(html, canvas)
        return html, extract_stub(html, canvas, T)

    async def render(self, markdown, T):
        """-> (pptx bytes, slide count)"""
        _, lay = await self.layout(markdown, T)
        buf = io.BytesIO()
        build_pptx.build(lay, buf, T)
        return buf.getvalue(), len(lay["slides"])

    def count_slides(self, markdown, T):
        return md2html.to_html(markdown, T).count('<section class="slide">')
