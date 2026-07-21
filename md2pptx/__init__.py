# -*- coding: utf-8 -*-
"""md2pptx — Markdown と PowerPoint テーマ（thmx/pptx）から発表スライドを生成する．

パッケージの公開 API はコンソールスクリプト（cli.main）が中心．ライブラリとして
使う場合は parser.parse_file() で Deck を得て render.build() で描画できる．
"""

__version__ = "0.9.0"

__all__ = ["__version__"]
