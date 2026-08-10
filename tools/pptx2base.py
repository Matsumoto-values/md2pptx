#!/usr/bin/env python3
"""pptx2base.py -- アップロードされた資料を「テンプレート専用ベース」に変換する.

CLI : pptx2base.py corp.pptx base.pptx [keep_layout_name]
API : to_base(data: bytes, keep_layout=None) -> (bytes, info: dict)

やること:
  1. 受入検査（サイズ・zip爆弾・マクロ）— 顧客がアップロードする信用できない入力なので
  2. .potx のコンテンツタイプを .pptx に書き換え（python-pptxが.potxを開けないため）
  3. **スライドを全削除** — マスターとレイアウトだけ残せばデザインは完全に再現できる。
     顧客の機密スライドを保存しないため（実測で64〜96%の容量削減にもなる）
  4. 未使用レイアウトの削除（任意）— バッチで毎回読み込むので軽い方がよい

注意: デザイントークンの抽出は**この変換より先に**行う必要がある。
配色・フォント・余白は元のスライドから実測するため。
"""
import io
import sys
import zipfile

from pptx import Presentation

MAX_BYTES = 80 * 1024 * 1024        # 圧縮後の上限
MAX_UNCOMPRESSED = 600 * 1024 * 1024  # 展開後の上限（zip爆弾対策）
MAX_MEMBERS = 5000
PML_TEMPLATE = ("application/vnd.openxmlformats-officedocument"
                ".presentationml.template.main+xml")
PML_PRESENTATION = ("application/vnd.openxmlformats-officedocument"
                    ".presentationml.presentation.main+xml")


class RejectedFile(Exception):
    """受入検査で弾いた入力。呼び出し側が4xxで返すことを想定する."""


def validate(data):
    if len(data) > MAX_BYTES:
        raise RejectedFile(f"file too large: {len(data)} bytes")
    try:
        zf = zipfile.ZipFile(io.BytesIO(data))
    except zipfile.BadZipFile as e:
        raise RejectedFile(f"not a valid Office file: {e}") from e
    infos = zf.infolist()
    if len(infos) > MAX_MEMBERS:
        raise RejectedFile(f"too many members: {len(infos)}")
    total = sum(i.file_size for i in infos)
    if total > MAX_UNCOMPRESSED:
        raise RejectedFile(f"uncompressed size too large: {total} bytes")
    names = zf.namelist()
    if not any(n.startswith("ppt/") for n in names):
        raise RejectedFile("not a PowerPoint package")
    # マクロを含むものは受け付けない（保存も配布もしたくない）
    if any(n.endswith("vbaProject.bin") for n in names):
        raise RejectedFile("macro-enabled file is not accepted")
    return zf, names


def as_presentation_package(data, names):
    """.potx を .pptx として開けるようにコンテンツタイプを書き換える.

    python-pptx の valid_content_types にテンプレートが含まれないため。
    中身は同一で、宣言だけの違い。
    """
    src = zipfile.ZipFile(io.BytesIO(data))
    ct = src.read("[Content_Types].xml").decode("utf-8")
    if PML_TEMPLATE not in ct:
        return data, False
    ct = ct.replace(PML_TEMPLATE, PML_PRESENTATION)
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as out:
        for n in names:
            out.writestr(n, ct.encode("utf-8") if n == "[Content_Types].xml"
                         else src.read(n))
    return buf.getvalue(), True


def strip_slides(prs):
    """スライドを全削除する（python-pptxにAPIが無いのでsldIdLstを直接操作）."""
    lst = prs.slides._sldIdLst
    n = len(lst)
    for sldId in list(lst):
        prs.part.drop_rel(sldId.rId)
        lst.remove(sldId)
    return n


def prune_layouts(prs, keep):
    """指定レイアウト以外を削除する。バッチで毎回読むので軽くしたい場合に使う.

    マスターごとに最低1件は残す（レイアウトが0件のマスターは不正になるため）。
    keep が無いマスターでは先頭だけを残す。
    """
    removed = 0
    for master in prs.slide_masters:
        lst = master.slide_layouts._sldLayoutIdLst
        entries = []
        for lid in list(lst):
            part = master.part.rels[lid.rId].target_part
            entries.append((lid, part.slide_layout.name))
        keepers = [e for e in entries if e[1] == keep] or entries[:1]
        keep_ids = {id(e[0]) for e in keepers}
        for lid, _name in entries:
            if id(lid) in keep_ids:
                continue
            lst.remove(lid)
            master.part.drop_rel(lid.rId)
            removed += 1
    return removed


def to_base(data, keep_layout=None):
    """-> (base_pptx_bytes, info)"""
    _, names = validate(data)
    data, converted = as_presentation_package(data, names)
    prs = Presentation(io.BytesIO(data))
    info = {"was_template": converted,
            "masters": len(prs.slide_masters),
            "layouts_before": sum(len(m.slide_layouts) for m in prs.slide_masters),
            "slides_removed": strip_slides(prs),
            "bytes_before": len(data)}
    if keep_layout:
        try:
            info["layouts_removed"] = prune_layouts(prs, keep_layout)
        except Exception as e:            # noqa: BLE001 レイアウト削除は best effort
            info["layouts_removed"] = 0
            info["prune_error"] = f"{type(e).__name__}: {e}"
    out = io.BytesIO()
    prs.save(out)
    info["bytes_after"] = out.tell()
    info["layouts_after"] = sum(len(m.slide_layouts)
                                for m in Presentation(io.BytesIO(out.getvalue())).slide_masters)
    return out.getvalue(), info


def main():
    data = open(sys.argv[1], "rb").read()
    base, info = to_base(data, sys.argv[3] if len(sys.argv) > 3 else None)
    open(sys.argv[2], "wb").write(base)
    for k, v in info.items():
        print(f"{k}: {v}", file=sys.stderr)


if __name__ == "__main__":
    main()
