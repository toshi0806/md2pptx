#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""pptx → PDF 変換（生成後プレビュー用の土台．DESIGN.md §7 / Issue #39）．

「Markdown を編集しながら PDF を見る」運用の基礎として，生成した pptx を
そのまま PDF にする．**忠実度は保証しない**：LibreOffice の出力はテーマ
フォントの解決差などで実 PowerPoint と一致しない（README 参照）。見た目の
最終確認は実 PowerPoint で行う前提は変えない．

変換器は 3 系統:

- ``auto``（既定）: native PowerPoint → LibreOffice の順に，使えるものを試す．
- ``powerpoint`` / ``libreoffice``: その系統を名指し．
- 任意のコマンド行: ``ppt2pdf -o {output} {input}`` のように直接指定．
  プレースホルダ ``{input}`` / ``{output}`` / ``{outdir}`` を置換する．1 つも
  無ければ末尾に ``{input}`` を補う（出力パスを取らないツール向け）．その場合
  ツールは入力の隣に ``<basename>.pdf`` を書く想定で，期待パスと違えば移動する．

**macOS の native PowerPoint は非対応**：この PowerPoint ビルドの AppleScript 辞書には
``export`` コマンドが無く，標準 ``save … as save as PDF`` は宛先型により無反応または保存
ダイアログでハングする（実測）。ハングは実害なので試みず明示エラーにし，``auto`` は
LibreOffice へフォールバックする．忠実な mac 変換は VM 経由の ``ppt2pdf`` を
``--pdf-converter`` で使う．Windows の PowerPoint は COM（``SaveAs`` format 32）で対応．

このモジュールは cli 以外に依存しない（python-pptx 非依存）．外部プロセスの
起動と，どのバイナリを使うかの解決だけを担う．
"""
from __future__ import annotations

import os
import pathlib
import shlex
import shutil
import subprocess
import sys
import tempfile


class PdfError(Exception):
    """PDF 変換の失敗（原因メッセージ付き）．cli が警告表示に使う．"""


# 環境変数名（CLI 引数 --pdf-converter が優先）．
ENV_CONVERTER = "MD2PPTX_PDF_CONVERTER"


def _which_libreoffice() -> str | None:
    """LibreOffice の実行ファイルを探す．PATH 優先，無ければ OS 既知の場所．"""
    for name in ("soffice", "libreoffice"):
        found = shutil.which(name)
        if found:
            return found
    candidates: list[str] = []
    if sys.platform == "darwin":
        candidates.append("/Applications/LibreOffice.app/Contents/MacOS/soffice")
    elif sys.platform.startswith("win"):
        for env in ("ProgramFiles", "ProgramFiles(x86)"):
            base = os.environ.get(env)
            if base:
                candidates.append(
                    os.path.join(base, "LibreOffice", "program", "soffice.exe"))
    for path in candidates:
        if os.path.isfile(path):
            return path
    return None


def _run(cmd: list[str], what: str) -> None:
    """外部コマンドを実行し，失敗を PdfError に変換する．

    成功時の出力は捨てる．失敗時のみ stderr（無ければ stdout）の末尾 1 行を
    原因として拾う（cli が警告に整形する）．
    """
    try:
        proc = subprocess.run(cmd, capture_output=True, text=True)
    except FileNotFoundError:
        raise PdfError(f"{what}: command not found: {cmd[0]}")
    if proc.returncode != 0:
        detail = (proc.stderr or proc.stdout or "").strip().splitlines()
        tail = detail[-1] if detail else f"exit {proc.returncode}"
        raise PdfError(f"{what} failed: {tail}")


def _convert_libreoffice(src: str, dst: str) -> None:
    """LibreOffice で src(pptx) → dst(pdf)．--outdir 方式なので後で改名する．"""
    soffice = _which_libreoffice()
    if soffice is None:
        raise PdfError(
            "LibreOffice not found (looked for soffice/libreoffice on PATH "
            "and the default install location)")
    outdir = os.path.dirname(os.path.abspath(dst)) or "."
    # 同一プロファイルの多重起動は失敗しうるので，毎回使い捨てのプロファイルを渡す．
    with tempfile.TemporaryDirectory(prefix="md2pptx-lo-") as profile:
        # as_uri() は Windows のドライブレターも file:///C:/... と正しく組む
        # （手組みの "file://"+path だと file://C:/... になり不正）．
        uri = pathlib.Path(os.path.abspath(profile)).as_uri()
        _run([
            soffice, "--headless",
            f"-env:UserInstallation={uri}",
            "--convert-to", "pdf", "--outdir", outdir, src,
        ], "libreoffice")
    # soffice は <入力 basename>.pdf を outdir に書く．期待名と違えば移動する．
    produced = os.path.join(
        outdir, os.path.splitext(os.path.basename(src))[0] + ".pdf")
    _finish(produced, dst, "libreoffice")


def _convert_powerpoint(src: str, dst: str) -> None:
    """native PowerPoint（macOS: AppleScript / Windows: COM）で変換する．"""
    src_abs = os.path.abspath(src)
    dst_abs = os.path.abspath(dst)
    if sys.platform == "darwin":
        # macOS の PowerPoint（AppleScript）は無人 PDF 化が不安定．sdef に `export`
        # コマンドは無く，標準 `save … in <file> as save as PDF` は宛先の型次第で
        # 無反応（POSIX パス文字列）か保存ダイアログでハング（POSIX file）になる
        # （このビルドで実測）．ハングは実害なので試みず，明示エラーにして呼び出し側に
        # 委ねる：`auto` は LibreOffice へフォールバックし，忠実な変換は VM 経由の
        # `--pdf-converter 'ppt2pdf -o {output} {input}'` を使う．
        raise PdfError(
            "native PowerPoint on macOS is not supported for headless PDF "
            "(no reliable AppleScript path on this build); use LibreOffice or "
            "'--pdf-converter \"ppt2pdf -o {output} {input}\"'")
    if sys.platform.startswith("win"):
        # PowerShell + COM．32 = ppSaveAsPDF．パスは単一引用符文字列に埋めるので，
        # パス内の ' は '' にエスケープする（O'Brien 等でコマンドが壊れるのを防ぐ）．
        src_ps = src_abs.replace("'", "''")
        dst_ps = dst_abs.replace("'", "''")
        ps = (
            "$ppt = New-Object -ComObject PowerPoint.Application; "
            "$pres = $ppt.Presentations.Open("
            f"'{src_ps}', $true, $false, $false); "
            f"$pres.SaveAs('{dst_ps}', 32); "
            "$pres.Close(); $ppt.Quit()"
        )
        _run(["powershell", "-NoProfile", "-Command", ps], "powerpoint")
    else:
        raise PdfError("native PowerPoint is only available on macOS or Windows")
    if not os.path.isfile(dst_abs):
        raise PdfError("powerpoint did not produce a PDF")


def _convert_custom(command: str, src: str, dst: str) -> None:
    """任意のコマンド行で変換する．プレースホルダを置換して実行する．"""
    outdir = os.path.dirname(os.path.abspath(dst)) or "."
    parts = shlex.split(command)
    if not parts:
        raise PdfError(f"empty {ENV_CONVERTER}/--pdf-converter command")
    # 判定はすべて分割後のトークン（parts）で行い，元文字列との二重基準を避ける．
    has_output = any("{output}" in p for p in parts)
    has_ph = has_output or any(
        ("{input}" in p or "{outdir}" in p) for p in parts)
    if not has_ph:
        # 出力パスを取らないツール（例: ppt2pdf out.pptx）向け：入力を末尾に補う．
        parts.append("{input}")
    subst = {"input": src, "output": dst, "outdir": outdir}
    cmd = [p.format(**subst) for p in parts]
    _run(cmd, "converter")
    if has_output:
        # ツールが {output} をそのまま書いたはず．そこに無ければ失敗．
        if not os.path.isfile(dst):
            raise PdfError(f"converter did not write {dst}")
        return
    # {output} を渡していない場合，ツールは入力の隣に <basename>.pdf を書く想定．
    produced = os.path.splitext(os.path.abspath(src))[0] + ".pdf"
    _finish(produced, dst, "converter")


def _finish(produced: str, dst: str, what: str) -> None:
    """ツールが書いた PDF(produced) を期待パス(dst) へ収める．"""
    produced = os.path.abspath(produced)
    dst = os.path.abspath(dst)
    if not os.path.isfile(produced):
        raise PdfError(f"{what} did not produce a PDF (expected {produced})")
    if produced != dst:
        os.replace(produced, dst)


def default_pdf_path(output_pptx: str) -> str:
    """--pdf を PATH 無しで使ったときの既定 PDF パス（出力 pptx と同じ場所・basename）．"""
    return os.path.splitext(output_pptx)[0] + ".pdf"


def convert(src: str, dst: str, converter: str | None) -> None:
    """src(pptx) を dst(pdf) へ変換する．

    Args:
        src: 入力 pptx．
        dst: 出力 pdf．
        converter: 変換器の指定．None または "auto" で自動探索
            （PowerPoint → LibreOffice）．"powerpoint" / "libreoffice" で名指し．
            それ以外は任意のコマンド行として解釈する．

    Raises:
        PdfError: 変換に失敗したとき（cli が警告に整形する）．
    """
    if not os.path.isfile(src):
        raise PdfError(f"pptx not found: {src}")

    name = (converter or "auto").strip()

    if name == "auto":
        # 使える系統を順に試す．PowerPoint が無ければ LibreOffice へ．
        errors: list[str] = []
        if sys.platform == "darwin" or sys.platform.startswith("win"):
            try:
                _convert_powerpoint(src, dst)
                return
            except PdfError as e:
                errors.append(str(e))
        try:
            _convert_libreoffice(src, dst)
            return
        except PdfError as e:
            errors.append(str(e))
        raise PdfError(
            "no PDF converter available "
            "(tried PowerPoint / LibreOffice; use --pdf-converter or install "
            "LibreOffice)\n  - " + "\n  - ".join(errors))

    if name == "libreoffice":
        _convert_libreoffice(src, dst)
    elif name == "powerpoint":
        _convert_powerpoint(src, dst)
    else:
        _convert_custom(name, src, dst)
