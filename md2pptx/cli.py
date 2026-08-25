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
    md2pptx input.md --watch --pdf  # 保存のたびに作り直す（編集しながらのプレビュー）
    md2pptx --version             # バージョンを表示して終了する
    python3 -m md2pptx input.md   # インストールせず開発中に実行する場合

1 回ぶんのビルドは ``build_once`` に切り出してあり，一発実行と ``--watch`` が共有する
（``watch.py`` は「いつ作り直すか」だけを担い，何を作るかは知らない）．
"""
from __future__ import annotations

import argparse
import os
import sys
from dataclasses import dataclass
from typing import Iterable

import pptx2pdf as pdf_backend       # PDF 変換の実装（別パッケージ）

from . import __version__
from . import parser as md_parser  # 標準ライブラリ parser とは別物
from . import render
from . import watch
from .ir import Deck, Image
from .thmx2pptx import ThmxError, thmx_to_pptx

# PDF 変換の環境変数．**名前は md2pptx のもの**を保つ——変換の実装が別パッケージ
# （pptx2pdf）へ移っても，利用者が設定してきた名前が効かなくなる理由は無い．
# ここで解決した値を convert() へ明示的に渡す（pptx2pdf 自身の PPTX2PDF_* は，
# ここで何も見つからなかったときの下位の既定として効く）．
ENV_CONVERTER = "MD2PPTX_PDF_CONVERTER"
ENV_TIMEOUT = "MD2PPTX_PDF_TIMEOUT"


class BuildError(Exception):
    """1 回のビルドの想定内の失敗（原因つき）．

    一発実行では ``main`` が ``md2pptx: <理由>`` の SystemExit に整形し，``--watch``
    ではその場に表示して次の保存を待つ．**``build_once`` は SystemExit を投げない**
    ——投げると watch のループごと死に，直して保存しても誰も作り直さなくなる．

    Attributes:
        sources: この試行で判明した入力ファイル（watch の監視対象）．失敗しても
            そこまでに分かったものは返す——「画像が無い」で失敗したなら，その画像が
            置かれたときに作り直したいから．
    """

    def __init__(self, message: str, sources: Iterable[str] = ()) -> None:
        super().__init__(message)
        self.sources: frozenset[str] = frozenset(sources)


@dataclass(frozen=True)
class BuildResult:
    """1 回のビルドの成果．

    Attributes:
        output: 書いた pptx．
        pdf: 書いた PDF（作らなかった／変換に失敗したときは None）．
        sources: このビルドが読んだファイル（Markdown・テーマ・画像）．
    """

    output: str
    pdf: str | None
    sources: frozenset[str]


def load_base(theme_path: str, keep_base: str | None = None) -> tuple[str, bool]:
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
    raise BuildError(
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
    raise BuildError(
        f"front matter '{key}' must be a string, got "
        f"{type(value).__name__} ({value!r})"
    )


def _image_sources(deck: Deck, base_dir: str) -> list[str]:
    """Deck が参照する画像ファイルを列挙する（``--watch`` の監視対象）．

    単一カラム（``blocks``）と多カラム（``columns``）の両方を見る．解決規則は
    描画側と同じ ``render.resolve_image_path``——ここで独自に組み立てると，
    どちらかを直したときにもう片方が置き去りになる．
    """
    found: list[str] = []
    for slide in deck.slides:
        # 冗長に見えるが**このコピーを外してはいけない**．多カラムのスライドでは
        # parser が `columns = [blocks, []]` と組むので `columns[0] is blocks`．
        # `blocks = slide.blocks` にすると下の extend が slide.blocks 自身を継ぎ足し，
        # IR を書き換えてしまう（実測：ブロック 1 個のスライドが 3 個になる）．
        blocks = list(slide.blocks)
        for column in slide.columns:
            blocks.extend(column)
        for block in blocks:
            if isinstance(block, Image):
                found.append(
                    os.path.abspath(render.resolve_image_path(block.src, base_dir)))
    return found


def _parse_args(argv: list[str] | None) -> argparse.Namespace:
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
    # --pdf は値を取らない．値を省略できるオプション（nargs="?"）にすると
    # `md2pptx --pdf deck.md` で入力ファイルが --pdf の値として食われ，「input が無い」
    # という原因の分からないエラーになる（Issue #42）．出力先は --pdf-output で取る．
    ap.add_argument(
        "--pdf", action="store_true",
        help="also render a PDF after the pptx, next to the output pptx "
             "(fidelity depends on the converter: LibreOffice is a preview, "
             "PowerPoint is the real render)",
    )
    ap.add_argument(
        "--pdf-output", metavar="PATH",
        help="render the PDF to PATH; implies --pdf",
    )
    ap.add_argument(
        "--pdf-converter", metavar="NAME|COMMAND",
        help="PDF backend: 'auto' (default), 'powerpoint', 'libreoffice', or a "
             "command line with {input}/{output}/{outdir} placeholders; "
             f"overrides ${ENV_CONVERTER}",
    )
    ap.add_argument(
        "--pdf-timeout", metavar="SEC", type=float,
        help="give up on the converter after SEC seconds (0 = wait forever); "
             f"overrides ${ENV_TIMEOUT}. Without it, md2pptx waits forever when "
             "stderr is a terminal and gives up after 180s when it is not",
    )
    # --watch は --pdf を含意しない（#42 と同じ理由）．pptx だけを最新に保つ運用も
    # 正当で，「見張れ」と「PDF も作れ」は別の指示．編集しながらのプレビューには
    # --watch --pdf を組み合わせる（README 参照）．
    ap.add_argument(
        "--watch", action="store_true",
        help="keep running and rebuild whenever the Markdown (or its theme or "
             "images) changes; Ctrl-C to stop",
    )
    return ap.parse_args(argv)


def main(argv: list[str] | None = None) -> int:
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


def _run(args: argparse.Namespace) -> int:
    _check_invocation(args)
    if args.watch:
        return _run_watch(args)
    try:
        build_once(args)
    except BuildError as e:
        raise SystemExit(f"md2pptx: {e}")
    return 0


def _check_invocation(args: argparse.Namespace) -> None:
    """起動時に一度だけ行う引数の検査．

    **watch でもここは即座に失敗させる**（打ち間違いを見張り続けても仕方がない）．
    これ以降の失敗は原稿の側の問題なので，watch では表示するだけで待ち続ける．
    """
    if not os.path.isfile(args.input):
        raise SystemExit(f"md2pptx: input not found: {args.input}")
    # 空の --pdf-output は「指定なし」と区別が付かないまま黙って PDF 生成を落とす
    # （`--pdf-output "$PDF_OUT"` で変数が未設定のときに起こる）．黙って何もしないより，
    # ここで落とす．pptx を書く前に検査するので、失敗しても成果物が中途半端に残らない．
    if args.pdf_output is not None and not args.pdf_output.strip():
        raise SystemExit("md2pptx: --pdf-output requires a path")


def _run_watch(args: argparse.Namespace) -> int:
    """`--watch`：入力とその依存を見張り，変わるたびに作り直す（止まらない）．"""
    previous: frozenset[str] = frozenset()

    def build() -> frozenset[str]:
        nonlocal previous
        try:
            # 端末は tty でも，人が見ているのはエディタと PDF であってこの端末では
            # ない．無制限に待つと以後のプレビューが全部止まるので，打ち切って次の
            # 保存で作り直す方に賭ける（前面化もしない．pdf.convert 参照）．
            result = build_once(args, unattended=True)
        except BuildError as e:
            sys.stderr.write(f"md2pptx: {e}\n")
            # 失敗しても監視は続ける．前回の依存も残すのは，足りない画像を置いた／
            # theme を直した，というときに作り直したいから．
            previous = previous | e.sources
        else:
            # **自分の出力は見張らない**．theme に出力 pptx を指されると，作る →
            # 変わった → また作る，の無限ループになる．
            made = {os.path.abspath(result.output)}
            if result.pdf:
                made.add(os.path.abspath(result.pdf))
            previous = result.sources - made
        return previous

    # 入力 Markdown だけはビルド前から分かっている．初回ビルド（実 PowerPoint なら
    # 数秒）の最中に保存されたぶんを取りこぼさないよう，先に見張り始める．
    return watch.run(build, label=args.input,
                     seed=[os.path.abspath(args.input)])


def _pdf_timeout(explicit: float | None) -> float | None:
    """``--pdf-timeout`` → ``MD2PPTX_PDF_TIMEOUT`` の順に上限（秒）を解決する．

    どちらも無ければ ``None``（＝指定なし）を返し，決め方は pptx2pdf に委ねる
    （``PPTX2PDF_TIMEOUT``，それも無ければ stderr が tty かどうか）．値の妥当性
    （負値・``nan``・``inf``）を見るのも向こう側で，ここでやるのは **md2pptx 固有の
    名前から値を取り出すところまで**．

    Raises:
        pdf_backend.PdfError: 環境変数が秒数として読めないとき．呼び出し側が
            警告に整形するので，終了コードは変わらない．
    """
    if explicit is not None:
        return explicit
    raw = (os.environ.get(ENV_TIMEOUT) or "").strip()
    if not raw:
        return None
    try:
        return float(raw)
    except ValueError:
        raise pdf_backend.PdfError(
            f"invalid {ENV_TIMEOUT}: {raw!r} (seconds, 0 = no limit)")


def build_once(args: argparse.Namespace, *,
               unattended: bool = False) -> BuildResult:
    """引数 1 セットぶんを 1 回ビルドする（parse → 解決 → render → 任意で PDF）．

    一発実行と ``--watch`` の共通部分．``saved:`` 等の表示もここで行う——標準出力と
    標準エラーに出る順序を，一発実行と watch で同一に保つため．

    Args:
        unattended: 端末が tty でも「人は見ていない」として PDF 変換に伝える．

    Raises:
        BuildError: 想定内の失敗．**SystemExit は投げない**（watch が死ぬ）．
    """
    # 失敗しても，そこまでに判明した入力は watch へ返す（BuildError に載せる）．
    sources: set[str] = {os.path.abspath(args.input)}
    try:
        return _build(args, sources, unattended)
    except BuildError as e:
        raise BuildError(str(e), sources | e.sources)
    except (ThmxError, FileNotFoundError, PermissionError) as e:
        # thmx 変換の失敗や，入力ファイルの不在・権限エラーはトレースバックでは
        # なく整形メッセージで失敗させる（§7）．OSError 全般には広げない
        # ——想定外の入出力エラー（例：ドライブ切断）はトレースバックを残す．
        raise BuildError(str(e), sources)


def _build(args: argparse.Namespace, sources: set[str],
           unattended: bool) -> BuildResult:
    """``build_once`` の本体．``sources`` に読んだファイルを足しながら進む．"""
    # 1) Markdown -> IR（Deck）
    try:
        deck = md_parser.parse_file(args.input)
    except Exception as e:  # パースエラーは原因を表示して失敗させる（§7）．
        raise BuildError(f"failed to parse {args.input}: {e}")

    meta = deck.meta or {}

    # 2) テーマ・出力先を解決（CLI 引数 > フロントマター）．
    theme = _as_path(args.theme or meta.get("theme"), "theme")
    if not theme:
        raise BuildError(
            "no theme specified (use --theme or front matter 'theme')"
        )
    # フロントマターの相対パスは Markdown ファイルからの相対として解決する．
    if not os.path.isabs(theme) and not os.path.isfile(theme):
        cand = os.path.join(os.path.dirname(os.path.abspath(args.input)), theme)
        if os.path.isfile(cand):
            theme = cand
    sources.add(os.path.abspath(theme))

    output = _as_path(args.output or meta.get("output"), "output")
    if not output:
        raise BuildError(
            "no output specified (use -o or front matter 'output')"
        )

    # 3) base pptx へ収束 → レンダリング → 保存．
    # 画像などの相対パスは Markdown ファイルの置き場を基準に解決する．
    base_dir = os.path.dirname(os.path.abspath(args.input))
    sources.update(_image_sources(deck, base_dir))
    base_path, is_temp = load_base(theme, keep_base=args.keep_base)
    try:
        render.build(deck, base_path, output, base_dir=base_dir)
    except Exception as e:  # 描画エラーも原因を表示して失敗させる（§7）．
        raise BuildError(f"failed to render {args.input}: {e}")
    finally:
        if is_temp and os.path.exists(base_path):
            os.remove(base_path)

    n = len(deck.slides) + (1 if deck.title_slide is not None else 0)
    # watch では出力をパイプへ流されることがある（ブロックバッファで無音に見える）．
    print(f"saved: {output} slides: {n}", flush=True)

    # 4) 任意: PDF も生成（プレビュー用）．失敗しても pptx は成功なので終了コードは
    # 変えない——編集しながらのプレビュー運用を変換失敗で止めないため（Issue #39）．
    # --pdf-output は単体で生成を有効にする（出力先を書いた人が「作るな」を意図する
    # ことはない）．--pdf-converter と --pdf-timeout は「どう作るか」の指定なので
    # 有効化しない——環境変数を export しただけで全実行が PDF を作り始めてしまうため
    # （Issue #42）．
    pdf_out: str | None = None
    if args.pdf or args.pdf_output:
        pdf_out = args.pdf_output or pdf_backend.default_pdf_path(output)
        converter = args.pdf_converter or os.environ.get(ENV_CONVERTER)
        # 変換は数秒かかる（LibreOffice は例で ~4 秒）．無音で止まって見えないよう
        # 開始を stderr に出す（stdout の saved: 行は汚さない）．
        sys.stderr.write(f"md2pptx: converting to PDF: {pdf_out}\n")
        try:
            # 上限の解決も try の中で行う．環境変数の書き間違いは PdfError になり，
            # 下の except が警告に整形する——「PDF が作れなくても pptx は成功」
            # （Issue #39）を，値が壊れていた場合にも保つ．
            timeout = _pdf_timeout(args.pdf_timeout)
            pdf_backend.convert(output, pdf_out, converter, timeout,
                                unattended=unattended)
            print(f"saved: {pdf_out}", flush=True)
        except pdf_backend.PdfError as e:
            sys.stderr.write(f"md2pptx: warning: PDF not generated: {e}\n")
            pdf_out = None      # 出来ていないものを「作った」とは報告しない

    return BuildResult(output=output, pdf=pdf_out, sources=frozenset(sources))


if __name__ == "__main__":
    sys.exit(main())
