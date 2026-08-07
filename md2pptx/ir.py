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
from typing import Literal, TypeGuard, get_args

# 水平寄せ．表の列・画像の配置で共通に使う．
Align = Literal["left", "center", "right"]

# レイアウト番号．OOXML では**先頭から「タイトル スライド」「タイトルとコンテンツ」
# 「セクション見出し」**という並びがテーマの慣行で，thmx / pptx のどちらから作った
# base でもここに来る．parser（見出しレベルの割り当て）と render（番号を付けない
# 判定・title_layout）が同じ値を見るので，契約であるこのモジュールに置く．
TITLE_LAYOUT = 0
CONTENT_LAYOUT = 1
SECTION_LAYOUT = 2


@dataclass
class Span:
    """段落の中の 1 続きの文字（run に 1 対 1 で対応する．DESIGN.md §5.13）．

    行内の装飾（``**強調**`` / ``[色]{red}`` / `` `等幅` `` / ``2^32^`` /
    ``[表示](url)``）を解釈した結果．装飾の無い行では ``Line.spans`` は空で，
    render は ``Line.text`` を 1 つの run として書く（従来どおり）．

    Attributes:
        text: この run の文字（装飾の記号は除去済み）．
        segment: この run が属する ``\\v`` 区切りセグメントの添字．
            **1 セグメントが複数 run に割れる**ので，相対サイズ（``Line.seg_deltas``）を
            付ける相手を位置ではなくこの値で決める．
        bold: 太字（``**…**``）．
        mono: 等幅（`` `…` ``）．フォント名は render が持つ（``mono_font``）．
        color: 文字色．テーマ色名／CSS の色名／16進のいずれか（``colors.parse_color``）．
        link: ハイパーリンクの URL（``[表示](url)`` 由来）．
        script: 上付き ``"sup"`` / 下付き ``"sub"``（``2^32^`` / ``H~2~O``）．
    """

    text: str
    segment: int = 0
    bold: bool = False
    mono: bool = False
    color: str | None = None
    link: str | None = None
    script: Literal["sup", "sub"] | None = None


@dataclass
class Line:
    """本文の 1 段落（箇条書き・自動採番・記号なしのいずれか）．

    Markdown の行頭マーカー（DESIGN.md §5.3）を解釈した結果を保持する．

    Attributes:
        text: 段落の表示テキスト（行頭マーカー記号は除去済み）．
        level: 箇条書きの深さ．0 が最上位，2 スペースのインデントごとに +1．
        kind: 段落種別．
            - "bullet"  : テーマ既定の箇条書き記号（記号は指定せず任せる）．
            - "autonum" : 自動採番（set_autonum 相当）．num_style で形式を指定．
            - "plain"   : 行頭記号なし（no_bullet 相当．結論・補足行など）．
            - "code"    : コードブロックの 1 行（フェンス由来．DESIGN.md §5.12）．
                行頭記号を消し等幅フォントで描く．text は原稿のまま
                （行頭マーカーもサイズトークンも <br> も解釈しない）．
            取りうる値は注釈（Literal）が正．
        num_style: kind=="autonum" のときの採番形式．python-pptx の
            buAutoNum type 値をそのまま使う．
            "arabicPeriod"（1. 2. 3.）/ "circleNumDbPlain"（丸数字 ①②③）/
            "arabicParenBoth"（丸括弧 (1) (2)）など．kind!="autonum" のときは None．
        num_color: 採番記号の色をテーマ色名で指定（例 "tx1"）．
            None ならテーマ任せ．kind=="autonum" のときのみ意味を持つ．
        num_start: 原稿に書かれていた番号（"8." なら 8，"⑤" なら 5）．
            **効くのはリストの先頭の行だけ**で，以降の行の値は使わない
            （CommonMark と同じ規則．"1. 1. 1." と書けば 1・2・3 になる）．
            番号を数えるのは render で，**全ての採番段落に buAutoNum の startAt を
            明示的に書く**——PowerPoint は startAt の付いた段落の次から数え直すため，
            先頭にだけ書くと "8. 1. 2. 3. …" になる（DESIGN.md §5.3）．
            kind!="autonum" では None．
        size_delta: 相対フォントサイズの段数（行頭 "{+1}"/"{-2}" 由来）．
            その行が level から得るテーマ既定サイズを基点に，1 段ごとに
            ×1.125（拡大）/ ÷1.125（縮小）する（render が実サイズへ換算）．
            None ならスライド既定（@body-size）に従う＝未指定．0 で「テーマ既定
            に固定（スライド既定を無効化）」を表す．絶対 pt は持たない（テーマ委譲）．
        spans: 行内装飾を解釈した run の列（DESIGN.md §5.13）．**空なら装飾なし**で，
            render は text を 1 つの run として書く（従来の経路）．非空なら
            連結した文字列が text と一致する（text は装飾記号を除いた素のテキスト）．
        seg_deltas: text を "\\v"（行内改行）で割った各セグメントの相対段数
            （2 番目以降のセグメント先頭 "{+1}"/"{-2}" 由来）．セグメントと同じ長さで，
            要素は int｜None（None＝未指定）．**[0] は常に None**——先頭セグメントの
            段数は size_delta が持つ．分けてあるのは書き込む先が違うからで，
            size_delta は段落の既定文字書式（defRPr）へ書く．そうしないと
            **ビュレットや採番記号のサイズが本文とずれる**．2 番目以降は
            段落を分けずに run だけを変えたいので run へ書く（DESIGN.md §5.8）．
    """

    text: str
    level: int = 0
    kind: Literal["bullet", "autonum", "plain", "code"] = "bullet"
    num_style: str | None = None
    num_color: str | None = None
    num_start: int | None = None
    size_delta: int | None = None
    seg_deltas: list[int | None] = field(default_factory=list)
    spans: list[Span] = field(default_factory=list)

    def __post_init__(self) -> None:
        # 不変条件：seg_deltas は text のセグメント数と同じ長さで，[0] は None．
        # 直接構築（テスト等）で食い違っても揃える——render は添字で run に
        # 対応付けるので，長さがずれると**別のセグメントにサイズが付く**．
        # 揃えるのは構築時のみ（TitleSlide.affiliation_deltas と同じ契約）．
        # IR は parser が一度構築し render が消費するもので，構築後に text を
        # 破壊的変更する運用は想定しない（同期はしない）．
        n = len(self.text.split("\v"))
        d = list(self.seg_deltas[:n])
        d += [None] * (n - len(d))
        if d:
            # [0] は捨てる．先頭セグメントの段数は size_delta が持つ——**Slide と
            # 非対称なのは書き込む先が違うから**で，こちらは段落の既定文字書式へ
            # 書かないとビュレット・採番記号のサイズが本文とずれる．
            d[0] = None
        self.seg_deltas = d


@dataclass
class Table:
    """表ブロック（Markdown 標準のテーブル記法由来．DESIGN.md §5.4）．

    Attributes:
        header: ヘッダ行のセル文字列リスト（アクセント色で着色する想定）．
        rows: 本体行のリスト．各行はヘッダと同じ列数のセル文字列リスト．
        aligns: 各列の水平寄せ（区切り行のコロン由来）．空リストは
            「指定なし＝すべて左」を意味する既定．列数に満たない場合，
            未指定の列は左寄せとして扱う（render 側で添字が範囲外なら "left"）．
        fills: 本体行のセル背景色（``rows`` と同じ形．要素は色名か None）．
            **空なら色指定なし**で，従来どおりテーマ任せに描く（DESIGN.md §5.4）．
        header_fills: ヘッダ行のセル背景色（``header`` と同じ長さ）．
            既定のアクセント色を上書きする．空なら従来どおり．
    """

    header: list[str] = field(default_factory=list)
    rows: list[list[str]] = field(default_factory=list)
    aligns: list[Align] = field(default_factory=list)
    fills: list[list[str | None]] = field(default_factory=list)
    header_fills: list[str | None] = field(default_factory=list)


@dataclass
class FlowNode:
    """フロー図のノード（box または省略記号）．

    Attributes:
        label: 主ラベル（[ラベル | サブラベル] の前半）．
        sublabel: 副ラベル（後半）．無ければ None．
        kind: "box"（角丸四角）または "ellipsis"（"…" 単独の省略記号．
            入力記法 `[…]` に対応する活字の省略記号のこと）．
        color: テーマ色名の個別指定（例 "accent6"）．None なら自動割当．
        node_id: エッジから指すための名前（``[#pc PC]`` 由来）．無ければ None．
            **``#`` で書くのはラベル中のコロンと区別するため**——``[id: ラベル]``
            だと ``[HTTP: ハイパーテキスト転送プロトコル]`` を名前付きと読んでしまう．
    """

    label: str = ""
    sublabel: str | None = None
    kind: Literal["box", "ellipsis"] = "box"
    color: str | None = None
    node_id: str | None = None


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
        nodes: ノードの列（出現順）．
        edges: エッジの列（隣接ノードを結ぶ）．
        caption: 図下キャプション．無ければ None．
        note_top: 図の上に置く注記．無ければ None．
        note_bottom: 図の下に置く注記．無ければ None．
        rows: 段ごとのノード index（``--`` 区切り由来．DESIGN.md §5.5）．
            **空なら段の指定なし＝一列**で，従来の原稿はこちらを通る．
            非空なら ``rows[i]`` が i 段目に並ぶノードの index 列．
    """

    direction: Literal["lr", "tb"] = "lr"
    nodes: list[FlowNode] = field(default_factory=list)
    edges: list[FlowEdge] = field(default_factory=list)
    rows: list[list[int]] = field(default_factory=list)
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

    unit: Literal["percent", "emu"]
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

    unit: Literal["px", "percent"]
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
        align: セグメント内の水平寄せ（既定は中央）．
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
    align: Align = "center"
    fit: Literal["contain", "fill"] = "contain"
    caption: str | None = None
    overflow: bool | None = None


# 帯（中央領域）へ座標配置するブロック．地の文（Line）と違い，本文プレースホルダ
# ではなく矩形に直接置かれる．render の _stack_objects / _obj_weight はこの 3 種
# だけを扱う（Line を渡すと属性が無く落ちる）．
ObjectBlock = Table | Flow | Image

# 実行時の判定用（``isinstance`` に渡せる形）．**ObjectBlock を増やしたら
# ここだけ直せばよい**——render は 5 か所でこの判定をするので，型注釈と別に
# タプルを書き並べると必ずどこかが漏れる（Issue #108）．
OBJECT_BLOCKS: tuple[type, ...] = get_args(ObjectBlock)

# スライド本文を構成するブロック．parser が出現順に並べ，render が型で分岐する．
# Union は平坦化されるので Line | Table | Flow | Image と同一．
Block = Line | ObjectBlock


def is_object_block(b: "Block") -> TypeGuard[ObjectBlock]:
    """帯へ座標配置するブロックか（``Line`` ではないか）を判定する．

    ``isinstance`` を直に書くと ``(Table, Flow, Image)`` が render の 5 か所へ
    散らばり，``ObjectBlock`` を増やしたときに必ずどこかが漏れる（Issue #108）．
    ``TypeGuard`` にしてあるので型の絞り込みも効く．
    """
    return isinstance(b, OBJECT_BLOCKS)


@dataclass
class Slide:
    """1 枚のコンテンツスライド．

    Attributes:
        title: スライドタイトル（"## 見出し" 由来）．タイトルなしなら None．
        title_deltas: title を "\\v"（行内改行）で割った各セグメントの相対段数．
            セグメントと同じ長さで，要素は int｜None（None＝未指定）．
            基点はタイトル枠の実効既定サイズ（render の _frame_font_levels）．
            **Line.seg_deltas と違い [0] も有効**——タイトルにはビュレットも
            採番記号も無いので，段落の既定文字書式に書き分ける理由がなく，
            すべて run へ書けば足りる（DESIGN.md §5.8）．
        layout: 使用するスライドレイアウト番号．既定 1（タイトルとコンテンツ）．
        blocks: スライド本文を構成するブロック列．出現順に保持する（混在可）．
            単一カラム時に使用．
        directives: スライド単位の上書き指示（DESIGN.md §5.6）．
            キーは parser の _KNOWN_DIRECTIVES に正規化済み．値はキーごとに
            int / str / bool と異なるため，使う側で narrowing する
            （例: {"autonum_color": "tx1", "layout": 5, "overflow": True}）．
        columns: 多カラム（「2つのコンテンツ」レイアウト）時の各カラムのブロック列．
            空なら単一カラム（blocks を使用）．非空なら columns[i] が i 番目の
            カラム内容で，レイアウトは 3 を既定とする（DESIGN.md §5.7）．
        notes: 発表者ノート（```note フェンス由来．DESIGN.md §5.10）．
            スライド面には描画せず，notes slide のテキストになる．
            "\\n" は段落区切り．無ければ None．
    """

    title: str | None = None
    title_deltas: list[int | None] = field(default_factory=list)
    layout: int = 1
    blocks: list[Block] = field(default_factory=list)
    directives: dict[str, object] = field(default_factory=dict)
    columns: list[list[Block]] = field(default_factory=list)
    notes: str | None = None

    def __post_init__(self) -> None:
        # 不変条件：title_deltas は title のセグメント数と同じ長さ（Line と同じ理由．
        # 揃えるのは構築時のみ）．**[0] は Line と違い捨てない**——タイトルには
        # ビュレットも採番記号も無く，段落側へ書き分ける理由がないので，
        # 先頭セグメントも run へ書けば足りる．
        # title が None なら長さ 0．タイトルの無いスライドに段数だけ渡されても
        # 置き場所が無いので落とす（そもそも parser がそう作らない）．
        n = len(self.title.split("\v")) if self.title else 0
        d = list(self.title_deltas[:n])
        self.title_deltas = d + [None] * (n - len(d))


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
        notes: 発表者ノート（本文開始前の ```note フェンス由来．DESIGN.md §5.10）．
            Slide.notes と同じ意味．無ければ None．
    """

    title: str | None = None
    subtitle: str | None = None
    author: str | None = None
    affiliation: list[str] = field(default_factory=list)
    subtitle_delta: int | None = None
    author_delta: int | None = None
    affiliation_deltas: list[int | None] = field(default_factory=list)
    notes: str | None = None

    def __post_init__(self) -> None:
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
            default_autofit などを含む生の dict）．YAML 由来なので値の型は
            キーごとに異なる（使う側で narrowing する）．
        title_slide: タイトルスライド．無ければ None．
        slides: コンテンツスライドの列（出現順）．
    """

    meta: dict[str, object] = field(default_factory=dict)
    title_slide: TitleSlide | None = None
    slides: list[Slide] = field(default_factory=list)
