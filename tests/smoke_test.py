#!/usr/bin/env python3
"""smoke_test.py -- ブラウザ無しでAPI経路とバッチ経路を通すスモークテスト.

python3 tests/smoke_test.py
"""
import os
import pathlib
import sys

ROOT = pathlib.Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))
os.environ.setdefault("EXTRACT_BACKEND", "stub")

from fastapi.testclient import TestClient   # noqa: E402

import app as service                        # noqa: E402

MD = (ROOT / "examples/deck.md").read_text(encoding="utf-8")


def main():
    fails = []
    with TestClient(service.app) as c:
        r = c.get("/healthz")
        print("healthz  ", r.status_code, r.json())
        fails += [] if r.status_code == 200 else ["healthz"]

        r = c.get("/themes")
        print("themes   ", r.status_code, r.json())

        r = c.post("/preview", json={"markdown": MD})
        ok = r.status_code == 200 and 'data-pptx="table"' in r.text
        print("preview  ", r.status_code, f"{len(r.text)} bytes", "OK" if ok else "NG")
        fails += [] if ok else ["preview"]

        r = c.post("/layout", json={"markdown": MD})
        n = len(r.json()["slides"])
        print("layout   ", r.status_code, f"{n} slides")
        fails += [] if n == 3 else ["layout"]

        # テーマの部分上書きが効くか（ブランド色だけ差し替える想定）
        r = c.post("/render", json={
            "markdown": MD, "filename": "themed.pptx",
            "theme": {"palette": ["FF6B00", "004E98"],
                      "table": {"zebra": False, "header_bg": "004E98"}}})
        ct = r.headers.get("content-type", "")
        ok = r.status_code == 200 and "presentation" in ct and len(r.content) > 20000
        print("render   ", r.status_code, f"{len(r.content)} bytes",
              f"slides={r.headers.get('X-Slide-Count')}", "OK" if ok else "NG")
        fails += [] if ok else ["render"]

        r = c.post("/render", json={"markdown": MD, "theme_id": "nope"})
        print("bad theme", r.status_code, "(400期待)")
        fails += [] if r.status_code == 400 else ["bad-theme"]

        r = c.post("/render", json={"markdown": ""})
        print("empty md ", r.status_code, "(422期待)")
        fails += [] if r.status_code == 422 else ["empty"]

    fails += batch_check()
    print("\n" + ("ALL PASSED" if not fails else f"FAILED: {fails}"))
    return 1 if fails else 0


def batch_check():
    """job.py を別プロセスで走らせ、1件失敗しても他が完走することを確認する."""
    import shutil
    import subprocess
    import tempfile
    out = tempfile.mkdtemp()
    try:
        r = subprocess.run(
            [sys.executable, str(ROOT / "job.py")], cwd=ROOT, capture_output=True,
            env={**os.environ, "EXTRACT_BACKEND": "stub",
                 "MANIFEST": str(ROOT / "examples/manifest.jsonl"),
                 "OUT_PREFIX": out})
        made = sorted(p.name for p in pathlib.Path(out).glob("*.pptx"))
        # manifestには意図的に失敗する1件が入っているので exit=1 が正常
        ok = r.returncode == 1 and len(made) == 4
        print("batch    ", f"exit={r.returncode}", f"{len(made)} files", made,
              "OK" if ok else "NG")
        return [] if ok else ["batch"]
    finally:
        shutil.rmtree(out, ignore_errors=True)


if __name__ == "__main__":
    sys.exit(main())
