#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""md2pptx の中間表現（IR）定義．

パーサ（parser.py）とレンダラ（render.py）の契約を担う純 Python の
データクラス群．Markdown の方言や DSL の詳細をレンダラから隠蔽し，
レンダラは IR の型だけを見て描画する．

DESIGN.md §4 に対応．外部依存を持たない（python-pptx 等は import しない）．
色・フォントはここでは扱わず，テーマに委ねる（採番記号色などのテーマ色名
だけを文字列で保持する）．
"""
from __future__ import annotations

from dataclasses import dataclass, field


@dataclass
class Line:
    """本文の 1 段落（箇条書き・自動採番・記号なしのいずれか）．

    Markdown の行頭マーカー（DESIGN.md §5.3）を解釈した結果を保持する．

    Attributes:
        text: 段落の表示テキスト（行頭マーカー記号は除去済み）．
        level: 箇条書きの深さ．0 が最上位，2 スペースのインデントごとに +1．
        kind: 段落種別．
            - "bullet"  : テーマ既定の箇条書き記号（add_bullets 相当）．
            - "autonum" : 自動採番（set_autonum 相当）．num_style で形式を指定．
            - "plain"   : 行頭記号なし（no_bullet 相当．結論・補足行など）．
        num_style: kind=="autonum" のときの採番形式．python-pptx の
            buAutoNum type 値をそのまま使う．
            "arabicPeriod"（1. 2. 3.）/ "circleNumDbPlain"（丸数字 ①②③）/
            "arabicParenBoth"（丸括弧 (1) (2)）など．kind!="autonum" のときは None．
        num_color: 採番記号の色をテーマ色名で指定（例 "tx1"）．
            None ならテーマ任せ．kind=="autonum" のときのみ意味を持つ．
    """

    text: str
    level: int = 0
    kind: str = "bullet"
    num_style: str | None = None
    num_color: str | None = None


@dataclass
class Table:
    """表ブロック（Markdown 標準のテーブル記法由来．DESIGN.md §5.4）．

    Attributes:
        header: ヘッダ行のセル文字列リスト（アクセント色で着色する想定）．
        rows: 本体行のリスト．各行はヘッダと同じ列数のセル文字列リスト．
    """

    header: list[str] = field(default_factory=list)
    rows: list[list[str]] = field(default_factory=list)


@dataclass
class Flow:
    """フロー図ブロック（```flow フェンス由来．DESIGN.md §5.5）．

    box / arrow による横並び（または縦並び）のフロー図を宣言的に表す．
    Phase 3 で本格対応するが，IR としては先に型を定義しておく．

    Attributes:
        nodes: ノード列．各ノードは box（[ラベル | サブラベル]）または
            省略記号 note（"…" 単独）を表す．要素表現はパーサ／レイアウタ
            (flow.py) の取り決めに従う．
        edges: エッジ列．矢印（->）と任意のラベル（-PR-> 等）を表す．
        caption: 図下キャプション．無ければ None．
    """

    nodes: list = field(default_factory=list)
    edges: list = field(default_factory=list)
    caption: str | None = None


@dataclass
class Slide:
    """1 枚のコンテンツスライド．

    Attributes:
        title: スライドタイトル（"## 見出し" 由来）．タイトルなしなら None．
        layout: 使用するスライドレイアウト番号．既定 1（タイトルとコンテンツ）．
        blocks: スライド本文を構成するブロック列．Line / Table / Flow を
            出現順に保持する（混在可）．
        directives: スライド単位の上書き指示（DESIGN.md §5.6）．
            例: {"autonum_color": "tx1", "layout": 5, "autofit": 90}．
    """

    title: str | None = None
    layout: int = 1
    blocks: list = field(default_factory=list)
    directives: dict = field(default_factory=dict)


@dataclass
class TitleSlide:
    """タイトルスライド（front matter 由来．あれば 1 枚目に生成）．

    DESIGN.md §5.1 のタイトルスライド情報に対応．

    Attributes:
        title: 主タイトル．改行を含む場合は段落分け（多段タイトル）．
        subtitle: 副題段落．無ければ None．
        author: 発表者名．無ければ None．
        affiliation: 所属・日付などの行リスト（著者欄に複数行で並べる）．
    """

    title: str | None = None
    subtitle: str | None = None
    author: str | None = None
    affiliation: list[str] = field(default_factory=list)


@dataclass
class Deck:
    """1 つの発表（pptx）全体に対応する最上位 IR．

    Attributes:
        meta: front matter 全体（theme / output / slide_number /
            default_autofit などを含む生の dict）．
        title_slide: タイトルスライド．無ければ None．
        slides: コンテンツスライドの列（出現順）．
    """

    meta: dict = field(default_factory=dict)
    title_slide: TitleSlide | None = None
    slides: list = field(default_factory=list)
