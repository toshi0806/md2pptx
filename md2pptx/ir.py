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
        size_delta: 相対フォントサイズの段数（行頭 "{+1}"/"{-2}" 由来）．
            その行が level から得るテーマ既定サイズを基点に，1 段ごとに
            ×1.125（拡大）/ ÷1.125（縮小）する（render が実サイズへ換算）．
            None ならスライド既定（@body-size）に従う＝未指定．0 で「テーマ既定
            に固定（スライド既定を無効化）」を表す．絶対 pt は持たない（テーマ委譲）．
    """

    text: str
    level: int = 0
    kind: str = "bullet"
    num_style: str | None = None
    num_color: str | None = None
    size_delta: int | None = None


@dataclass
class Table:
    """表ブロック（Markdown 標準のテーブル記法由来．DESIGN.md §5.4）．

    Attributes:
        header: ヘッダ行のセル文字列リスト（アクセント色で着色する想定）．
        rows: 本体行のリスト．各行はヘッダと同じ列数のセル文字列リスト．
        aligns: 各列の水平寄せ（区切り行のコロン由来）．要素は
            "left" / "center" / "right"．空リストは「指定なし＝すべて左」を
            意味する既定．列数に満たない場合，未指定の列は左寄せとして扱う
            （render 側で添字が範囲外なら "left"）．
    """

    header: list[str] = field(default_factory=list)
    rows: list[list[str]] = field(default_factory=list)
    aligns: list[str] = field(default_factory=list)


@dataclass
class FlowNode:
    """フロー図のノード（box または省略記号）．

    Attributes:
        label: 主ラベル（[ラベル | サブラベル] の前半）．
        sublabel: 副ラベル（後半）．無ければ None．
        kind: "box"（角丸四角）または "ellipsis"（"…" 単独の省略記号）．
        color: テーマ色名の個別指定（例 "accent6"）．None なら自動割当．
    """

    label: str = ""
    sublabel: str | None = None
    kind: str = "box"
    color: str | None = None


@dataclass
class FlowEdge:
    """フロー図のエッジ（ノード間の矢印）．

    Attributes:
        src: 始点ノードの index（Flow.nodes 内）．
        dst: 終点ノードの index．
        label: 矢印上のラベル（-PR-> の "PR"）．無ければ None．
    """

    src: int = 0
    dst: int = 0
    label: str | None = None


@dataclass
class Flow:
    """フロー図ブロック（```flow フェンス由来．DESIGN.md §5.5）．

    box / arrow による横並び（lr）または縦並び（tb）のフロー図を宣言的に表す．

    Attributes:
        direction: 並び方向．"lr"（左→右，既定）/ "tb"（上→下）．
        nodes: FlowNode の列（出現順）．
        edges: FlowEdge の列（隣接ノードを結ぶ）．
        caption: 図下キャプション．無ければ None．
        note_top: 図の上に置く注記．無ければ None．
        note_bottom: 図の下に置く注記．無ければ None．
    """

    direction: str = "lr"
    nodes: list = field(default_factory=list)
    edges: list = field(default_factory=list)
    caption: str | None = None
    note_top: str | None = None
    note_bottom: str | None = None


@dataclass
class Length:
    """埋め込みサイズの 1 次元（画像の width / height）．

    unit で解釈が変わる：
        - "percent": セグメント（帯）矩形に対する割合（value は 0..100）．render で解決．
        - "emu":     絶対サイズ（value は EMU 整数相当．parser が cm/pt/in/px から換算）．
    """

    unit: str
    value: float


@dataclass
class Crop:
    """画像トリミングの「残す矩形」（左上原点）．

    unit で x/y/w/h の単位が変わる：
        - "px":      ソース画像のピクセル座標（解像度依存）．
        - "percent": ソース画像サイズに対する割合（0..100．解像度非依存）．
    render がソース画像の実ピクセル寸法を読み，PowerPoint のクロップ割合
    （各辺 0..1）へ換算する．
    """

    unit: str
    x: float
    y: float
    w: float
    h: float


@dataclass
class Image:
    """画像ブロック（jpg / png）．表・フロー図と同じ「オブジェクト」として中央帯に配置する．

    Markdown の `![cap](src){opts}` ショートハンド，または ```image フェンス由来
    （DESIGN.md §5.9）．描画（python-pptx の add_picture）は render の責務で，
    ここ（IR）はパス・サイズ・トリミング等の宣言のみ保持する（外部依存なし）．

    Attributes:
        src: 画像ファイルのパス（Markdown ファイルからの相対 or 絶対）．
        width: 埋め込み幅（Length）．None なら height かアスペクトから決める．
        height: 埋め込み高（Length）．None なら width かアスペクトから決める．
        crop: トリミングの残す矩形（Crop）．None ならトリミングなし．
        align: セグメント内の水平寄せ．"left" / "center"（既定）/ "right"．
        fit: width/height 両指定時の収め方．"contain"（既定・比維持で内接）/
            "fill"（歪ませて充填）．片方のみ・省略時は常にアスペクト維持．
        caption: 図下キャプション（省略可．ショートハンドは alt を採用）．
        overflow: True なら帯（セグメント）に収める最終クランプを行わず，
            width/height で明示したサイズのままはみ出しを許可する．
            はみ出す方向は下（結論文・罫線側）のみで，タイトル側へは重ねない．
            None は「ブロックでの指定なし」＝スライドの @overflow ディレクティブ
            （既定 False）に従う．ブロック指定（True/False）はスライド指定に優先する．
    """

    src: str
    width: Length | None = None
    height: Length | None = None
    crop: Crop | None = None
    align: str = "center"
    fit: str = "contain"
    caption: str | None = None
    overflow: bool | None = None


@dataclass
class Slide:
    """1 枚のコンテンツスライド．

    Attributes:
        title: スライドタイトル（"## 見出し" 由来）．タイトルなしなら None．
        layout: 使用するスライドレイアウト番号．既定 1（タイトルとコンテンツ）．
        blocks: スライド本文を構成するブロック列．Line / Table / Flow を
            出現順に保持する（混在可）．単一カラム時に使用．
        directives: スライド単位の上書き指示（DESIGN.md §5.6）．
            例: {"autonum_color": "tx1", "layout": 5, "autofit": 90}．
        columns: 多カラム（「2つのコンテンツ」レイアウト）時の各カラムのブロック列．
            空なら単一カラム（blocks を使用）．非空なら columns[i] が i 番目の
            カラム内容で，レイアウトは 3 を既定とする（DESIGN.md §5.7）．
    """

    title: str | None = None
    layout: int = 1
    blocks: list = field(default_factory=list)
    directives: dict = field(default_factory=dict)
    columns: list = field(default_factory=list)


@dataclass
class TitleSlide:
    """タイトルスライド（front matter 由来．あれば 1 枚目に生成）．

    DESIGN.md §5.1 のタイトルスライド情報に対応．

    Attributes:
        title: 主タイトル．改行を含む場合は段落分け（多段タイトル）．
        subtitle: 副題段落．無ければ None．
        author: 発表者名．無ければ None．
        affiliation: 所属・日付などの行リスト（著者欄に複数行で並べる）．
        subtitle_delta: 副題の相対フォントサイズ段数（先頭 "{-1}" 由来．None＝未指定）．
        author_delta: 著者名の相対フォントサイズ段数（同上）．
        affiliation_deltas: affiliation 各行と 1 対 1 対応する相対サイズ段数リスト
            （各要素 int｜None．None＝未指定）．本文の Line.size_delta と同じ意味で，
            render がテーマ既定サイズを基点に実サイズへ換算する．
    """

    title: str | None = None
    subtitle: str | None = None
    author: str | None = None
    affiliation: list[str] = field(default_factory=list)
    subtitle_delta: int | None = None
    author_delta: int | None = None
    affiliation_deltas: list[int | None] = field(default_factory=list)

    def __post_init__(self):
        # 不変条件：affiliation_deltas は affiliation と同じ長さ（各行 1 対 1）．
        # 直接構築（テスト等）で長さがずれても None 詰め／切り詰めで揃え，
        # render 側が添字で安全に対応付けられるようにする．
        # 揃えるのは構築時のみ．IR は parser が一度構築し render が消費する契約で，
        # 構築後に affiliation を破壊的変更する運用は想定しない（同期はしない）．
        n = len(self.affiliation)
        d = self.affiliation_deltas
        if len(d) < n:
            self.affiliation_deltas = list(d) + [None] * (n - len(d))
        elif len(d) > n:
            self.affiliation_deltas = list(d[:n])


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
