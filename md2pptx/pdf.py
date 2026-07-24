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

このモジュールは cli 以外に依存しない（python-pptx 非依存）．外部プロセスの
起動と，どのバイナリを使うかの解決だけを担う．
"""
from __future__ import annotations

import os
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
    """外部コマンドを実行し，失敗を PdfError に変換する（stdout/stderr は捨てる）．"""
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
        uri = "file://" + os.path.abspath(profile).replace(os.sep, "/")
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
        # `export ... as save as PDF` を使う．`save ... as save as PDF` は保存
        # ダイアログが出て GUI でハングする（実測）が，export は無人で書き出せる．
        # in/to は辞書上 type="text" なので dst は POSIX パス文字列をそのまま渡す．
        # open の戻り値は束縛できないので active presentation を取る
        # （ppt2pdf.ps1 の COM 版が ActivePresentation へフォールバックするのと同じ）．
        script = (
            'on run {srcPath, dstPath}\n'
            '  tell application "Microsoft PowerPoint"\n'
            '    open srcPath\n'
            '    set pres to active presentation\n'
            '    export pres to dstPath as save as PDF\n'
            '    close pres saving no\n'
            '  end tell\n'
            'end run'
        )
        _run(["osascript", "-e", script, src_abs, dst_abs], "powerpoint")
    elif sys.platform.startswith("win"):
        # PowerShell + COM．32 = ppSaveAsPDF．
        ps = (
            "$ppt = New-Object -ComObject PowerPoint.Application; "
            "$pres = $ppt.Presentations.Open("
            f"'{src_abs}', $true, $false, $false); "
            f"$pres.SaveAs('{dst_abs}', 32); "
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
    has_ph = any(
        ("{input}" in p or "{output}" in p or "{outdir}" in p) for p in parts)
    if not has_ph:
        # 出力パスを取らないツール（例: ppt2pdf out.pptx）向け：入力を末尾に補う．
        parts.append("{input}")
    subst = {"input": src, "output": dst, "outdir": outdir}
    cmd = [p.format(**subst) for p in parts]
    _run(cmd, "converter")
    if has_ph and "{output}" in command:
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
        if sys.platform in ("darwin",) or sys.platform.startswith("win"):
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
