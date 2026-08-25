# -*- coding: utf-8 -*-
"""md2pptx — Markdown と PowerPoint テーマ（thmx/pptx）から発表スライドを生成する．

パッケージの公開 API はコンソールスクリプト（cli.main）が中心．ライブラリとして
使う場合は parser.parse_file() で Deck を得て render.build() で描画できる．
"""
import pptx2pdf

__version__ = "1.2.0"

__all__ = ["__version__"]

# PDF 変換は pptx2pdf（別パッケージ）が行う．そちらが出す警告や案内は **md2pptx として
# 起動された以上 md2pptx の名前で**出さなければならない——`md2pptx deck.md --pdf` の
# 途中で `pptx2pdf: warning: …` と出ても，利用者はそんなコマンドを打っていない．
# 上限や変換器の指定方法も、こちらのフラグ名で案内する．
# import 時に一度だけ行う（cli 以外の入口——ライブラリとしての利用や -m 実行——でも
# 同じ名前になるように）．
pptx2pdf.set_program_name("md2pptx")
pptx2pdf.set_hints(timeout="--pdf-timeout / MD2PPTX_PDF_TIMEOUT",
                   converter="--pdf-converter")
