#!/usr/bin/env python3
"""template_register.py -- アップロードされた資料をテンプレートとして登録する.

登録は「稀・対話的・人間のレビューあり」、生成は「頻繁・自動」。
性質が違うので処理も分ける。

流れ:
  1. 受入検査（サイズ・zip爆弾・マクロ）
  2. **元のスライドからデザイントークンを実測抽出**（ベース化より先。色や余白は
     スライドから測るため）
  3. ベース化（スライド全削除 → 顧客の機密を保存しない・64〜96%軽量化）
  4. フォント解決の判定 → font_status を確定させる
  5. draft として保存し、プレビュー用の情報を返す

フォントについて: SaaSでは利用者がサーバーにフォントを入れられない。
そのため「このテンプレートのフォントは再現できない」ことを登録時に確定させ、
代替を使うかどうかを人間に選ばせる。生成のたびに警告が出る状態にはしない。
"""
import io
import pathlib
import re
import sys

sys.path.insert(0, str(pathlib.Path(__file__).resolve().parent))

from tools import pptx2base, pptx2theme   # noqa: E402

SAFE_ID = re.compile(r"^[A-Za-z0-9][A-Za-z0-9_-]{0,62}$")


def check_id(value, label):
    if not SAFE_ID.match(value or ""):
        raise pptx2base.RejectedFile(f"invalid {label}: {value!r}")
    return value


def font_status(theme, available):
    """フォントの再現可否を確定させる.

    available: 与えられた名前が描画側で使えるかを返す関数
    -> (status, 実際に計測に使うフォント名, 説明)
    """
    declared = theme["font"]
    stack = theme.get("font_stack") or [declared]
    if available(declared):
        return "ok", declared, f"'{declared}' はそのまま計測できる"
    # font_stack[0] が字幅互換の代替なので先頭から見る（宣言フォント自身は飛ばす）
    for cand in stack:
        if cand == declared:
            continue
        if available(cand):
            same = cand == pptx2theme.METRIC_COMPATIBLE.get(declared.strip().lower())
            return ("substituted" if same else "approximated"), cand, (
                f"'{declared}' は使えないため '{cand}' で計測する"
                + ("（字幅互換）" if same else "（字幅は一致しないため折り返しがズレ得る）"))
    return "unavailable", None, (
        f"'{declared}' もフォールバック候補も使えない。"
        f"テーマの font を変更するか、フォントを導入する必要がある")


def fc_available(name):
    """fc-match でフォントが解決できるかを判定する（描画側の実環境を見る）."""
    import subprocess
    try:
        out = subprocess.run(["fc-match", "-f", "%{family}", name],
                             capture_output=True, text=True, timeout=10).stdout
    except (OSError, subprocess.SubprocessError):
        return False
    # フォールバックすると別名が返る
    return name.split()[0].lower() in out.lower()


def register(data, tenant_id, template_id, store, *, name=None,
             keep_layout_only=False, available=fc_available):
    """-> (レコード, 登録レポート)"""
    check_id(tenant_id, "tenant_id")
    check_id(template_id, "template_id")

    # 受入検査は最初に通す。信用できない入力を python-pptx に渡す前に弾く
    # （検査が to_base の中だけにあると、抽出の方が先に走って例外が漏れる）
    pptx2base.validate(data)

    # 抽出はベース化より先。色・フォント・余白は元のスライドから測る
    try:
        theme, dbg = pptx2theme.build(io.BytesIO(data))
    except pptx2base.RejectedFile:
        raise
    except Exception as e:              # noqa: BLE001 壊れた/未対応の資料
        raise pptx2base.RejectedFile(
            f"could not parse presentation: {type(e).__name__}: {e}") from e
    _, fonts, prs, design = dbg

    status, render_font, note = font_status(theme, available)
    theme["font_status"] = status
    theme["_font_warning"] = note
    if render_font and render_font != theme["font_stack"][0]:
        theme["font_stack"] = [render_font] + [f for f in theme["font_stack"]
                                               if f != render_font]

    base, info = pptx2base.to_base(
        data, theme["base_layout"] if keep_layout_only else None)
    theme["base_pptx"] = None          # 保管層が実体を管理するので固定パスは持たせない

    version = store.next_version(tenant_id, template_id)
    record = store.save(tenant_id, template_id, version, theme, base, status="draft")

    report = {
        "tenant_id": tenant_id, "template_id": template_id, "version": version,
        "status": "draft", "name": name or template_id,
        "font": {"declared": theme["font"], "render_with": render_font,
                 "status": status, "note": note},
        "palette": theme["palette"],
        "canvas": theme["canvas"],
        "layout": theme["base_layout"],
        "slides_removed": info["slides_removed"],
        "size": {"uploaded": info["bytes_before"], "base": info["bytes_after"]},
        "design_layer_colors": design,
        "fonts_found": fonts.most_common(4),
    }
    return record, report
