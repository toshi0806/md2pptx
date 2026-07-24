#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""md2pptx CLI（Phase 1 のエントリポイント．DESIGN.md §7）．

Markdown を入力に取り，フロントマター／CLI 引数で解決したテーマ（thmx/pptx）を
土台にして pptx を生成する．処理の流れは次のとおり（DESIGN.md §2 / §3.5）．

    parse_file(input.md) -> Deck
    テーマ・出力先を解決（CLI 引数 > フロントマター）
    load_base(theme)     -> base pptx パス（.thmx は変換，.pptx はそのまま）
    render.build(deck, base, out)
    "saved: <out> slides: <n>" を出力

使い方::

    md2pptx input.md --theme OfficeTheme.pptx -o out.pptx
    md2pptx input.md              # フロントマターの theme/output を使う
    md2pptx --version             # バージョンを表示して終了する
    python3 -m md2pptx input.md   # インストールせず開発中に実行する場合
"""
from __future__ import annotations

import argparse
import os
import sys

from . import __version__
from . import parser as md_parser  # 標準ライブラリ parser とは別物
from . import pdf as pdf_backend
from . import render
from .pdf import ENV_CONVERTER
from .thmx2pptx import ThmxError, thmx_to_pptx


def load_base(theme_path, keep_base=None):
    """テーマ（.thmx / .pptx）を base pptx のパスへ収束させる（DESIGN.md §3.5）．

    Args:
        theme_path: テーマファイルのパス（.thmx か .pptx）．
        keep_base: .thmx 変換時の base pptx 出力先（指定すれば破棄しない）．
            None なら一時ファイルへ書き出す（呼び出し側で破棄する）．

    Returns:
        (base_path, is_temp). is_temp が True のときは呼び出し側で削除する．
    """
    ext = os.path.splitext(theme_path)[1].lower()
    if ext == ".thmx":
        if keep_base:
            return thmx_to_pptx(theme_path, keep_base), False
        return thmx_to_pptx(theme_path), True
    if ext == ".pptx":
        # 既に base 形式なのでそのまま土台に使う（変換も一時ファイルも不要）．
        return theme_path, False
    raise SystemExit(
        f"unsupported theme format: {ext or '(none)'} "
        "(expected .thmx or .pptx)"
    )


def _as_path(value: object, key: str) -> str | None:
    """CLI 引数／フロントマター由来のパス値を str へ検証する．

    front matter は YAML なので `theme: 123` のように非文字列が来うる．素通しすると
    os.path が TypeError を投げてトレースバックが出るため，ここで原因の分かる
    メッセージに変える（未指定を表す None はそのまま返す）．
    """
    if value is None or isinstance(value, str):
        return value
    raise SystemExit(
        f"md2pptx: front matter '{key}' must be a string, got "
        f"{type(value).__name__} ({value!r})"
    )


def _parse_args(argv):
    ap = argparse.ArgumentParser(
        prog="md2pptx",
        description="Convert a Markdown deck into a themed .pptx (Phase 1).",
    )
    ap.add_argument(
        "--version",
        action="version",
        version=f"%(prog)s {__version__}",
    )
    ap.add_argument("input", help="input Markdown file")
    ap.add_argument(
        "--theme",
        help="theme file (.thmx or .pptx); overrides front matter 'theme'",
    )
    ap.add_argument(
        "-o", "--output",
        help="output .pptx; overrides front matter 'output'",
    )
    ap.add_argument(
        "--keep-base", metavar="PATH",
        help="keep the intermediate base pptx (from .thmx) at PATH",
    )
    ap.add_argument(
        "--pdf", nargs="?", const=True, default=None, metavar="PATH",
        help="also render a PDF after the pptx (preview only, not a faithful "
             "PowerPoint render); PATH defaults to the output with a .pdf suffix",
    )
    ap.add_argument(
        "--pdf-converter", metavar="NAME|COMMAND",
        help="PDF backend: 'auto' (default), 'powerpoint', 'libreoffice', or a "
             "command line with {input}/{output}/{outdir} placeholders; "
             f"overrides ${ENV_CONVERTER}",
    )
    return ap.parse_args(argv)


def main(argv=None):
    """CLI のエントリポイント．成功時は終了コード 0 を返し，失敗は SystemExit で終える．

    想定内の失敗（thmx 変換失敗・入力ファイルの不在／権限）は
    ``md2pptx: <理由>`` に整形して SystemExit する．想定外の例外はあえて
    握り潰さず（トレースバックのまま）伝播させ，バグを隠さない．
    """
    try:
        return _run(args=_parse_args(argv))
    except (ThmxError, FileNotFoundError, PermissionError) as e:
        # thmx 変換の失敗や，入力ファイルの不在・権限エラーはトレースバックでは
        # なく整形メッセージで失敗させる（§7）．OSError 全般には広げない
        # ——想定外の入出力エラー（例：ドライブ切断）はトレースバックを残す．
        raise SystemExit(f"md2pptx: {e}")


def _run(args):
    if not os.path.isfile(args.input):
        raise SystemExit(f"md2pptx: input not found: {args.input}")

    # 1) Markdown -> IR（Deck）
    try:
        deck = md_parser.parse_file(args.input)
    except Exception as e:  # パースエラーは原因を表示して失敗させる（§7）．
        raise SystemExit(f"md2pptx: failed to parse {args.input}: {e}")

    meta = deck.meta or {}

    # 2) テーマ・出力先を解決（CLI 引数 > フロントマター）．
    theme = _as_path(args.theme or meta.get("theme"), "theme")
    if not theme:
        raise SystemExit(
            "md2pptx: no theme specified (use --theme or front matter 'theme')"
        )
    # フロントマターの相対パスは Markdown ファイルからの相対として解決する．
    if not os.path.isabs(theme) and not os.path.isfile(theme):
        cand = os.path.join(os.path.dirname(os.path.abspath(args.input)), theme)
        if os.path.isfile(cand):
            theme = cand

    output = _as_path(args.output or meta.get("output"), "output")
    if not output:
        raise SystemExit(
            "md2pptx: no output specified (use -o or front matter 'output')"
        )

    # 3) base pptx へ収束 → レンダリング → 保存．
    # 画像などの相対パスは Markdown ファイルの置き場を基準に解決する．
    base_dir = os.path.dirname(os.path.abspath(args.input))
    base_path, is_temp = load_base(theme, keep_base=args.keep_base)
    try:
        render.build(deck, base_path, output, base_dir=base_dir)
    except Exception as e:  # 描画エラーも原因を表示して失敗させる（§7）．
        raise SystemExit(f"md2pptx: failed to render {args.input}: {e}")
    finally:
        if is_temp and os.path.exists(base_path):
            os.remove(base_path)

    n = len(deck.slides) + (1 if deck.title_slide is not None else 0)
    print(f"saved: {output} slides: {n}")

    # 4) 任意: PDF も生成（プレビュー用）．失敗しても pptx は成功なので終了コードは
    # 変えない——編集しながらのプレビュー運用を変換失敗で止めないため（Issue #39）．
    if args.pdf is not None:
        pdf_out = args.pdf if isinstance(args.pdf, str) \
            else pdf_backend.default_pdf_path(output)
        converter = args.pdf_converter or os.environ.get(ENV_CONVERTER)
        try:
            pdf_backend.convert(output, pdf_out, converter)
            print(f"saved: {pdf_out}")
        except pdf_backend.PdfError as e:
            sys.stderr.write(f"md2pptx: warning: PDF not generated: {e}\n")

    return 0


if __name__ == "__main__":
    sys.exit(main())
