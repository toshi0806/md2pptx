#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Markdown → 中間表現（IR）パーサ（md2pptx Phase 1）．

DESIGN.md §5 の Markdown 記法仕様を解釈し，ir.py のデータクラス
（Deck / TitleSlide / Slide / Line）へ変換する純 Python モジュール．
python-pptx には依存しない（描画は render.py の責務）．

担当範囲:
    - フロントマター（YAML）分離 → Deck.meta / TitleSlide 生成（§5.1）
    - スライド分割（"## 見出し" / "---" 水平線 / "# 見出し"）（§5.2）
    - 行頭マーカー解釈（"-"/"*"/"1."/丸数字/"(n)"/"→"）（§5.3）
    - 表（§5.4）・flow 図（§5.5）・画像（§5.9）のブロック生成
    - スライド単位ディレクティブ（HTML コメント）の収集（§5.6）
    - 発表者ノート（```note フェンス）の収集（§5.10）

描画（python-pptx）は一切行わない．flow の座標計算は flow.py，画像の実寸読み取り・
配置は render.py が担う．
"""
from __future__ import annotations

import re
import sys
from dataclasses import dataclass, replace
from typing import Any, Literal, cast

import yaml

from .ir import (
    CONTENT_LAYOUT, SECTION_LAYOUT, TITLE_LAYOUT, Align, Block, Crop, Deck,
    ARROW_DIRECTIONS, Arrow, ArrowDirection, Flow, Image, Length, Line, Slide,
    Span, Table, TitleSlide,
)
from .colors import parse_color
from .flow import parse_flow as _parse_flow
from .seq import parse_seq as _parse_seq


# ---------------------------------------------------------------- 定数

# 画像オプションで受理する値の集合．型付きなので "not in で弾いた残り" が
# Literal に絞られる（検証と型の単一の情報源にもなる）．
_ALIGNS: tuple[Align, ...] = ("left", "center", "right")
_FITS: tuple[Literal["contain", "fill"], ...] = ("contain", "fill")

# 丸数字 ①(U+2460) 〜 ⑳(U+2473)．行頭にあれば circleNumDbPlain として採番する．
CIRCLED_DIGITS = "".join(chr(c) for c in range(0x2460, 0x2474))

# 矢印（結論・補足行の目印）．no_bullet 相当の plain 段落になる．
ARROW = "→"

# 行頭マーカーの正規表現（インデント除去後の文字列に対して評価する）．
_RE_HEADING = re.compile(r"^(#{1,6})\s+(.*?)\s*$")
_RE_ORDERED = re.compile(r"^(\d+)\.\s+(.*)$")        # 1. 2. 3. …（arabicPeriod）
_RE_PAREN = re.compile(r"^\(\s*(\d+)\s*\)\s+(.*)$")  # (1) (2) …（arabicParenBoth）
_RE_DIRECTIVE = re.compile(r"^<!--\s*@([\w-]+)\s*:\s*(.*?)\s*-->$")
# カラム区切り（「2つのコンテンツ」レイアウト）．値は arrow のみ．
# **_RE_DIRECTIVE より先に評価すること**——"@col: arrow" は汎用の
# "@キー: 値" にも当たるので、順序が入れ替わると _apply_directive へ落ちる．
_RE_COL = re.compile(r"^<!--\s*@col(?:\s*:\s*(\S+?))?\s*-->$")
# 段階の区切り（アニメーションの代替．§5.11）．同じく値を取らない．
_RE_STEP = re.compile(r"^<!--\s*@step\s*-->$")
# 表紙（テーマの「タイトル スライド」レイアウト）．同じく値を取らない．
_RE_TITLE_SLIDE = re.compile(r"^<!--\s*@title-slide\s*-->$")
# 1 行 HTML コメント（ディレクティブ以外のメモ等．無視する）．
_RE_COMMENT = re.compile(r"^<!--.*-->$")
# Markdown テーブルの区切り行（例 "| --- | :--: |"）．ヘッダ行の直後に現れる．
_RE_TABLE_SEP = re.compile(r"^\|?\s*:?-{1,}:?\s*(\|\s*:?-{1,}:?\s*)*\|?$")
# タイトル・本文行内の明示改行マーカー（<br> / <br/>）．"\v"（行内改行）へ変換する．
_RE_BR = re.compile(r"\s*<br\s*/?>\s*")

# 相対フォントサイズトークン．マーカー直後・本文直前の "{+1}"/"{-2}"/"{0}"．
# 符号は省略可（"{2}" は "+2" と同義）．render がテーマ基準で実サイズへ換算する．
_RE_SIZE = re.compile(r"^\{\s*([+-]?\d+)\s*\}\s*(.*)$")
_RE_BOX = re.compile(r"^\{\s*box\s*(?::\s*([^}\s]+)\s*)?\}\s*(.*)$",
                     re.IGNORECASE)

# 整数として解釈するディレクティブキー（正規化後の名前）．
# body_size はスライド既定の相対フォントサイズ段数（@body-size）．
_INT_DIRECTIVES = {"layout", "autofit", "body_size"}

# 受理するディレクティブキー（正規化後の名前）．未知のキーはタイポの可能性が
# 高いので黙殺せずエラーにする（§5.6）．@col と @title-slide は値を取らない
# 専用形式（_RE_COL / _RE_TITLE_SLIDE）で，ここへ来るのは値付きの誤りだけ．
_KNOWN_DIRECTIVES = {
    "layout", "autofit", "body_size", "autonum_color", "widths", "table_widths",
    "overflow",
}

# v0.7 で改名した旧ディレクティブ名 → 新名称（エラーメッセージで案内する）．
_RENAMED_DIRECTIVES = {
    "ph_widths": "@widths",
    "body_width": "@widths",
    "col_widths": "@table-widths",
}

# フロントマターの既知キー．未知のキーはエラー（ディレクティブと同方針）．
_KNOWN_META_KEYS = {
    "theme", "output", "slide_number", "default_autofit", "syntax", "mono_font",
    "title", "subtitle", "author", "affiliation",
}

# 見出しレベルの割り当て（Issue #99）．``syntax:`` で選ぶ．**既定は 1**．
# 0 は従来の割り当てで，旧原稿には ``syntax: 0`` を書き足して使う．
# 既定を 1 にしたので，**書き足していない旧原稿は 1 段ずれて読まれる**
# （``# 章の扉`` が表紙になる）．ずれてもエラーにならないため，
# 表紙が 2 枚以上できる形だけは ``_check_title_slides`` で止める．
_SYNTAX_HEADINGS: dict[int, dict[int, int]] = {
    0: {1: SECTION_LAYOUT, 2: CONTENT_LAYOUT},
    1: {1: TITLE_LAYOUT, 2: SECTION_LAYOUT, 3: CONTENT_LAYOUT},
}
_DEFAULT_SYNTAX = 1

# 非推奨のフロントマターキー（Issue #82）．表紙は本文記法で書く．
# **受理はやめない**——動く原稿を黙って壊さないため，警告だけ出して従来どおり描く．
_DEPRECATED_META_KEYS = ("title", "subtitle", "author", "affiliation")

# 画像ショートハンド（標準 Markdown 画像＋末尾 "{opts}"）．§5.9．
# 例: "![実験結果](fig.png){width=70% align=left}"．opts は省略可．
_RE_IMAGE = re.compile(r"^!\[(?P<alt>[^\]]*)\]\((?P<src>[^)]*)\)\s*(?:\{(?P<opts>[^}]*)\})?\s*$")

# 絶対長さの単位 → EMU 係数（1in=914400 / 1cm=360000 / 1pt=12700 / 1px=9525@96dpi）．
_EMU_PER = {"in": 914400, "cm": 360000, "pt": 12700, "px": 9525}


# ---------------------------------------------------------------- 公開 API

def parse(md_text: str) -> Deck:
    """Markdown 文字列を Deck（IR の最上位）へ変換する．

    Args:
        md_text: Markdown ソース全文．先頭に YAML フロントマターを持てる．

    Returns:
        Deck. meta（フロントマター生 dict）・title_slide（無ければ None）・
        slides（コンテンツスライド列）を保持する．
    """
    text = _normalize_newlines(md_text)
    meta, body, body_offset = _split_front_matter(text)
    syntax = _validate_syntax(meta)
    _warn_deprecated_meta(meta, syntax)
    deck = Deck(meta=meta)
    deck.title_slide = _build_title_slide(meta)
    deck.slides, title_notes = _parse_body(
        body, body_offset, has_title_slide=deck.title_slide is not None,
        syntax=syntax)
    # 不変条件：title_notes が非 None なのは has_title_slide=True のときだけ
    # （タイトルスライド無しの本文前 ```note は _parse_body が ValueError にする．
    # 空の ```note は捨てられ title_notes に積まれない）．よってここで
    # deck.title_slide は必ず存在する．
    if title_notes is not None:
        assert deck.title_slide is not None  # 上の不変条件（型チェッカ向け）
        deck.title_slide.notes = title_notes
    return deck


def parse_file(path: str) -> Deck:
    """Markdown ファイルを読み込んで parse() する利便関数．"""
    with open(path, encoding="utf-8") as f:
        return parse(f.read())


# ---------------------------------------------------------------- フロントマター

def _normalize_newlines(text: str) -> str:
    """改行コードを LF に正規化する（CRLF / CR 対策）．"""
    return text.replace("\r\n", "\n").replace("\r", "\n")


def _split_front_matter(text: str) -> tuple[dict, str, int]:
    """先頭の "---" 〜 "---" を YAML として切り出す．

    Returns:
        (meta, body, body_offset). フロントマターが無ければ ({}, text, 0)．
        body_offset は本文開始までに消費したファイル行数（本文行番号の換算用）．
    """
    lines = text.split("\n")
    if lines and lines[0].strip() == "---":
        for i in range(1, len(lines)):
            if lines[i].strip() == "---":
                fm = "\n".join(lines[1:i])
                body = "\n".join(lines[i + 1:])
                try:
                    meta = yaml.safe_load(fm)
                except yaml.YAMLError as e:
                    # 不正な YAML は握り潰さず，行番号付きで報告する（§7）．
                    # フロントマター本体はファイル 2 行目（先頭 "---" の次）から．
                    mark = getattr(e, "problem_mark", None)
                    if mark is not None:
                        lineno = mark.line + 2
                        problem = getattr(e, "problem", None) or "syntax error"
                        raise ValueError(
                            f"invalid YAML front matter at line {lineno}: {problem}"
                        )
                    raise ValueError(f"invalid YAML front matter: {e}")
                if not isinstance(meta, dict):
                    meta = {}
                unknown = [k for k in meta if k not in _KNOWN_META_KEYS]
                if unknown:
                    keys = ", ".join(repr(k) for k in unknown)
                    known = ", ".join(sorted(_KNOWN_META_KEYS))
                    raise ValueError(
                        f"unknown front matter key(s): {keys} (known keys: {known})"
                    )
                return meta, body, i + 1
    return {}, text, 0


def _validate_syntax(meta: dict) -> int:
    """フロントマターの ``syntax:`` を検証して返す（既定 1．Issue #99）．

    受けるのは ``_SYNTAX_HEADINGS`` にある値だけ．未知の値は，将来の版で
    増える記法を古い md2pptx が黙って既定として読むのを防ぐために止める．
    """
    value = meta.get("syntax", _DEFAULT_SYNTAX)
    if value not in _SYNTAX_HEADINGS:
        known = ", ".join(str(k) for k in sorted(_SYNTAX_HEADINGS))
        raise ValueError(
            f"invalid syntax value {value!r} in front matter (known: {known})")
    return int(value)


def _warn_deprecated_meta(meta: dict, syntax: int) -> None:
    """フロントマターの表紙記述に非推奨の警告を出す（Issue #82）．

    ``title`` / ``subtitle`` / ``author`` / ``affiliation`` は「文書のメタデータ」の
    名前を持つが，実体は表紙スライドの描画記述だった（コアプロパティへは 1 度も
    書いていない）．同じ絵を出す記法が 2 つある状態が続き，``<br>`` を
    片方だけ通し忘れる取りこぼし（Issue #79）まで起きた．

    **止めはしない．** 動作は従来のまま——非推奨の間は既存原稿が同じ見た目で
    動き続けるほうが価値がある（副題の基点を凍結したのと同じ理由．Issue #83）．
    """
    used = [k for k in _DEPRECATED_META_KEYS if meta.get(k)]
    if not used:
        return
    # 移行先は syntax で変わる．0 では "#" が章の扉なので @title-slide が要る．
    marker = "" if TITLE_LAYOUT in _SYNTAX_HEADINGS[syntax].values() \
        else "  <!-- @title-slide -->\n"
    sys.stderr.write(
        "md2pptx: warning: front matter " + " / ".join(used)
        + " is deprecated; write the title slide in body syntax instead:\n"
        "  # 主題<br>{-3} 副題\n" + marker
        + "\n"
        "  - 著者\n"
        "  - 所属\n"
    )


def _warn_deprecated_rule(lineno: int) -> None:
    """``---`` によるスライド分割に非推奨の警告を出す（Issue #92）．

    ``---`` が作るのは「タイトルとコンテンツ」のタイトル枠を空にしたスライドで，
    テーマが用意していない形になる．**止めはしない**（front matter の表紙記述と
    同じ扱い）．``---`` 1 件ごとに，行番号を添えて出す．
    """
    sys.stderr.write(
        f"md2pptx: warning: line {lineno}: '---' as a slide break is deprecated; "
        "give the slide a heading, or pick a layout without a title frame:\n"
        "  ### 見出し\n"
        "  <!-- @layout: 6 -->      # 白紙．図・表だけのスライドはこちらへ置ける\n"
    )


def _build_title_slide(meta: dict) -> TitleSlide | None:
    """フロントマターからタイトルスライドを構築する（title が無ければ None）．"""
    if not meta.get("title"):
        return None

    title = meta.get("title")
    if isinstance(title, str):
        # 複数行タイトル（YAML ブロックスカラー）の末尾改行を落とす．
        title = title.rstrip("\n")
        # <br> は見出し・本文と同じ規則で行内改行（\v）へ．front matter だけ
        # 素通しにすると "<br>" がそのまま画面に出る（Issue #79）．
        title = _RE_BR.sub("\v", title)

    # 副題・著者・所属も本文行と同じ相対サイズトークン "{-1}"/"{+1}" を先頭に置ける．
    # トークンは本文から剥がし，段数を IR の *_delta へ格納する（render が換算）．
    subtitle_delta, subtitle = _split_size_opt(meta.get("subtitle"))
    author_delta, author = _split_size_opt(meta.get("author"))

    affiliation_raw = meta.get("affiliation")
    if affiliation_raw is None:
        # 未指定のみ空扱い．`affiliation: 0` のような falsy な値まで捨てないよう
        # `or []` にはしない（非文字列スカラーは下で 1 行へ正規化する）．
        affiliation_raw = []
    if not isinstance(affiliation_raw, list):
        # スカラー（"affiliation: 所属" や YAML が数値として読んだ値）は 1 行扱い．
        affiliation_raw = [affiliation_raw]
    affiliation: list[str] = []
    affiliation_deltas: list[int | None] = []
    for line in affiliation_raw:
        delta, text = _split_size_opt(line)
        # YAML の空要素（"-" だけの行）は None になる．空行として残す
        # （None のままだと affiliation: list[str] を破り，render で落ちる）．
        affiliation.append(text if text is not None else "")
        affiliation_deltas.append(delta)

    return TitleSlide(
        title=title,
        subtitle=subtitle,
        author=author,
        affiliation=affiliation,
        subtitle_delta=subtitle_delta,
        author_delta=author_delta,
        affiliation_deltas=affiliation_deltas,
    )


def _split_size_opt(value: object) -> tuple[int | None, str | None]:
    """front matter 値（None 可）の先頭相対サイズトークンを剥がして (段数, 文字列) を返す．

    None はそのまま (None, None)．トークン判定は文字列のみ対象とし，数値等
    （YAML が int/float で読んだ値）は素直に文字列化して段数なしで返す．
    副題・著者・所属もタイトル同様 <br> を行内改行（\\v）へ変換する（Issue #79）．
    """
    if value is None:
        return None, None
    if not isinstance(value, str):
        return None, str(value)
    delta, text = _split_size(value)
    return delta, _RE_BR.sub("\v", text)


# ---------------------------------------------------------------- 本文

def _parse_body(body: str, body_offset: int = 0,
                has_title_slide: bool = False,
                syntax: int = _DEFAULT_SYNTAX,
                ) -> tuple[list[Slide], str | None]:
    """本文をスライド列へ分割し，各行を IR ブロックへ変換する．

    body_offset はフロントマターが消費したファイル行数（エラー報告の行番号を
    ファイル先頭基準へ換算するために使う）．has_title_slide はフロントマター
    由来のタイトルスライドの有無（本文開始前の ```note の宛先判定に使う）．
    syntax は見出しレベルの割り当て（``_SYNTAX_HEADINGS``．既定 ``_DEFAULT_SYNTAX``）．

    Returns:
        (slides, title_notes)．title_notes は本文開始前に現れた ```note の
        内容（タイトルスライドの発表者ノート）．無ければ None．
    """
    slides: list[Slide] = []
    current: Slide | None = None
    title_notes: list[str] = []
    # 現在のスライドの段階区切り（@step）．各要素は「その時点での各カラムの
    # ブロック数」．**切り出しはスライドを閉じるときにまとめて行う**——@step の
    # 位置で即座にスナップショットを取ると，その後に書いた @layout や ```note が
    # 前の段に入らず，「どこに書いたか」で結果が変わってしまう（§5.11）．
    # 各要素は (各カラムのブロック数, 途中まで見せる図の位置 or None)．
    # 後者は図の中の @step（Issue #125）用で ``(カラム番号, 単位数)``．
    # ブロック境界だけでは図の内部で切れないので、
    # 「このカラムの最後のブロックだけ途中まで」を表せるようにしてある．
    # **カラム番号を持つ**のは、左右どちらにも図があるとき「最初に見つかった図」
    # では取り違えるため（左の図が常に選ばれてしまう）．
    step_marks: list[tuple[list[int], tuple[int, int] | None]] = []

    def ensure_slide() -> Slide:
        """直前にスライド開始マーカーが無いまま本文が来た場合のフォールバック．"""
        nonlocal current
        if current is None:
            current = Slide()
        return current

    def add_block(b: Block) -> None:
        """ブロックを現在のカラム（多カラム時）または blocks へ追加する．

        図が**中に段階を持って**いれば（``Flow.steps`` / ``Seq.steps``）、
        その数だけスライドの段を刻む（Issue #125）．最後の段は図の全体なので
        刻まない——``_expand_steps`` が最終段としてスライド本体を使う．
        """
        s = ensure_slide()
        (s.columns[-1] if s.columns else s.blocks).append(b)
        steps = getattr(b, "steps", None)
        if steps:
            cols = s.columns if s.columns else [s.blocks]
            counts = [len(c) for c in cols]
            ci = len(s.columns) - 1 if s.columns else 0
            for n in steps:
                step_marks.append((list(counts), (ci, n)))

    def flush() -> None:
        """現在のスライドを（段があれば展開して）slides へ移す．

        ``slides`` に ``nonlocal`` が要らないのは ``extend`` するだけで
        束縛し直さないため（``current`` と ``step_marks`` は代入するので要る）．
        """
        nonlocal current, step_marks
        if current is not None:
            slides.extend(_expand_steps(current, step_marks))
        current = None
        step_marks = []

    lines = body.split("\n")
    n = len(lines)
    i = 0
    title_slides = 0
    while i < n:
        raw = lines[i]
        stripped = raw.strip()
        lineno = body_offset + i + 1  # ファイル先頭基準の 1 始まり行番号

        # --- スライド分割マーカー ---------------------------------
        m = _RE_HEADING.match(raw.lstrip())
        if m:
            hashes, htext = m.group(1), m.group(2)
            # 見出しレベル → レイアウト番号は syntax で決まる（Issue #99）．
            # 割り当ての外のレベルは未定義（将来のスライド内小見出し用に予約）．
            headings = _SYNTAX_HEADINGS[syntax]
            level = len(hashes)
            if level not in headings:
                usable = " / ".join(
                    "'" + "#" * lv + "'" for lv in sorted(headings))
                raise ValueError(
                    f"H{level} heading is not supported at line {lineno}: "
                    f"{stripped!r} (syntax {syntax} uses {usable})")
            # タイトル内の <br> を行内改行（\v）へ変換し，各セグメント先頭の
            # サイズトークンを剥がす．**先頭セグメントも段数を取れる**——
            # 本文行と違いタイトルにはビュレットも採番記号も無いので，
            # 段落の既定文字書式へ書き分ける理由がない（ir.Slide.title_deltas）．
            head_delta, htext = _split_size(htext)
            htext, title_deltas = _split_br(htext)
            title_deltas[0] = head_delta
            if headings[level] == TITLE_LAYOUT and syntax != 0:
                # 表紙は 1 枚．H1 が文書に 1 つという Markdown の慣習に沿う．
                # **``syntax: 0`` を書き忘れた旧原稿がここで止まる**——旧原稿の
                # ``# 章の扉`` は複数あるので 2 枚目で引っかかる．黙って 1 段
                # ずれたデッキが出るより、止めて書き足してもらうほうがよい．
                title_slides += 1
                if title_slides > 1:
                    raise ValueError(
                        f"a second title slide at line {lineno}: {stripped!r} "
                        f"(a deck has one; if this is a section divider, "
                        f"write 'syntax: 0' in the front matter or use '##')")
            flush()
            current = Slide(title=htext or None, title_deltas=title_deltas,
                            layout=headings[level])
            i += 1
            continue

        if stripped == "---":
            # 水平線 → タイトルなしスライドを明示的に開始．**非推奨**（Issue #92）．
            _warn_deprecated_rule(lineno)
            flush()
            current = Slide()
            i += 1
            continue

        # --- 表紙（テーマの「タイトル スライド」レイアウト）----------
        if _RE_TITLE_SLIDE.match(stripped):
            # "#" 自体が表紙になる割り当てでは，この指定は冗長．黙って受けると
            # 「表紙の書き方が 2 通りある」状態に戻る（@layout との併記と同じ理由）．
            # syntax 番号ではなく**割り当ての中身**で見る——「表紙を持つ割り当てなら
            # 冗長」が理由なので，将来 syntax が増えてもその判断がそのまま効く．
            if TITLE_LAYOUT in _SYNTAX_HEADINGS[syntax].values():
                raise ValueError(
                    f"@title-slide is not used with syntax {syntax} at line "
                    f"{lineno} ('#' is already the title slide)")
            s = ensure_slide()
            # @layout との併記はエラー．結果が同じ "@layout: 0" も含めて弾く——
            # 「同じ結果なら許す」を入れると，矛盾する組み合わせ（@layout: 5 等）を
            # どちらで解決するかを別に決めることになる．規則は 1 つで足りる．
            if "layout" in s.directives:
                raise ValueError(
                    f"@title-slide conflicts with @layout at line {lineno} "
                    f"(use one or the other)")
            s.directives["title_slide"] = True
            s.layout = TITLE_LAYOUT
            i += 1
            continue

        # --- カラム区切り（「2つのコンテンツ」）→ 多カラム化（§5.7）----
        m_col = _RE_COL.match(stripped)
        if m_col:
            s = ensure_slide()
            if m_col.group(1) is not None:
                # 区切りそのものを図形として描く指定．いまは arrow だけ．
                if m_col.group(1) != "arrow":
                    raise ValueError(
                        f"invalid @col value {m_col.group(1)!r} at line "
                        f"{lineno} (arrow)")
                s.directives["col_arrow"] = True
            if not s.columns:
                s.layout = 3                 # 2つのコンテンツ レイアウト
                s.columns = [s.blocks, []]   # 既存ブロックを左カラムへ
            else:
                s.columns.append([])
            i += 1
            continue

        # --- 段階の区切り（アニメーションの代替）→ 段を1つ刻む（§5.11）----
        if _RE_STEP.match(stripped):
            s = ensure_slide()
            cols = s.columns if s.columns else [s.blocks]
            step_marks.append(([len(c) for c in cols], None))
            i += 1
            continue

        # --- スライド単位ディレクティブ（HTML コメント）-----------
        md = _RE_DIRECTIVE.match(stripped)
        if md:
            slide = ensure_slide()
            _apply_directive(slide, md.group(1), md.group(2), lineno)
            i += 1
            continue

        # --- ディレクティブ以外の HTML コメントは無視（メモ等）-----
        if _RE_COMMENT.match(stripped):
            i += 1
            continue

        # --- フェンスドブロック（```flow / ```image / ```note / コード）------
        if stripped.startswith("```"):
            info = stripped[3:].strip().lower()
            j = i + 1
            buf: list[str] = []
            while j < n and lines[j].strip() != "```":
                buf.append(lines[j])
                j += 1
            if j >= n:
                # 閉じ忘れを黙って末尾まで飲み込むと，以降のスライドが丸ごと
                # 消えたデッキが出る．**行番号付きで止める**（§7 の方針）．
                raise ValueError(
                    f"unclosed code fence at line {lineno}: "
                    f"{stripped!r} (add a closing ```)")
            if info == "arrow":
                add_block(_parse_arrow_block("\n".join(buf)))
            elif info == "flow":
                add_block(_parse_flow("\n".join(buf)))
            elif info == "seq":
                add_block(_parse_seq("\n".join(buf)))
            elif info == "image":
                add_block(_parse_image_block("\n".join(buf)))
            elif info in ("note", "notes"):
                # 発表者ノート（§5.10）．スライド面には出さず notes へ蓄積する．
                # 本文開始前（スライドマーカーより先）ならタイトルスライド宛て．
                # strip("\n") はフェンス境界に接する空行の正規化（先頭・末尾のみ）．
                # ノート冒頭の空段落は表示上意味を持たないため意図的に落とす
                # （内部の空行＝段落区切りは保持される）．
                text = "\n".join(buf).strip("\n")
                if text:
                    if current is None and not slides:
                        if not has_title_slide:
                            raise ValueError(
                                f"```note block at line {lineno} appears before "
                                f"any slide, but there is no title slide "
                                f"(add 'title:' to the front matter or move the "
                                f"block after a slide heading)")
                        title_notes.append(text)
                    else:
                        s = ensure_slide()
                        s.notes = text if s.notes is None else s.notes + "\n" + text
            else:
                # それ以外はすべてコードブロック（§5.12）．**info string は
                # 自由で，md2pptx は読み飛ばす**——構文強調はしないので言語名に
                # 意味が無く，受理する名前の一覧を持つと維持する羽目になる．
                # 行は原稿のまま Line にする（行頭マーカーもサイズトークンも
                # <br> も解釈しない．解釈したらもうコードではない）．
                # 前後の空行だけ落とす（フェンス境界に接する空行の正規化）．
                body_lines = list(buf)
                while body_lines and not body_lines[0].strip():
                    body_lines.pop(0)
                while body_lines and not body_lines[-1].strip():
                    body_lines.pop()
                # ``color`` と書いたフェンスだけ ``[語]{色}`` を解釈する
                # （既定は「書いたまま」．Issue #162）．
                colored = "color" in info.split()
                for text in body_lines:
                    spans: list[Span] = []
                    if colored:
                        text, spans = _code_color_spans(text)
                    add_block(Line(text=text, kind="code", spans=spans))
            i = j + 1  # 閉じフェンスの次へ
            continue

        # --- 画像ショートハンド（![cap](src){opts}）→ Image（§5.9）--------
        mi = _RE_IMAGE.match(stripped)
        if mi:
            add_block(_parse_image_shorthand(
                mi.group("alt"), mi.group("src"), mi.group("opts")))
            i += 1
            continue

        # --- 表（ヘッダ行＋直後の区切り行）→ Table（§5.4）---------
        if "|" in stripped and i + 1 < n and _RE_TABLE_SEP.match(lines[i + 1].strip()):
            header = _split_row(stripped)
            aligns = _parse_aligns(lines[i + 1].strip())
            j = i + 2
            rows: list[list[str]] = []
            while j < n:
                rs = lines[j].strip()
                if not rs or "|" not in rs:
                    break  # 空行や非テーブル行で表は終わり
                if rs == "---" or _RE_HEADING.match(lines[j].lstrip()):
                    break  # 別ブロック開始
                rows.append(_split_row(rs))
                j += 1
            # セル末尾の {色} を剥がす．**色を書いていない表では fills を空の
            # まま**にして、従来どおりの経路（テーマ任せ）を通す．
            header_cells = [_split_cell_fill(c) for c in header]
            body_cells = [[_split_cell_fill(c) for c in r] for r in rows]
            header_fills = [f for _, f in header_cells]
            fills = [[f for _, f in r] for r in body_cells]
            has_fill = any(header_fills) or any(any(r) for r in fills)
            add_block(Table(
                header=[t for t, _ in header_cells],
                rows=[[t for t, _ in r] for r in body_cells],
                aligns=aligns,
                fills=fills if has_fill else [],
                header_fills=header_fills if has_fill else []))
            i = j
            continue

        # --- 空行は段落区切り（Line は作らない）-------------------
        if not stripped:
            i += 1
            continue

        # --- 本文行 → Line ---------------------------------------
        line = parse_content_line(raw)
        if line is not None:
            add_block(line)
        i += 1

    flush()

    return slides, ("\n".join(title_notes) if title_notes else None)


def _expand_steps(
        slide: Slide,
        marks: list[tuple[list[int], tuple[int, int] | None]]) -> list[Slide]:
    """段階の区切り（@step）を持つスライドを、積み上がる複数枚へ展開する．

    marks の各要素は (その区切りの時点での各カラムのブロック数,
    途中まで見せる図の位置 or None)．**最終段は slide そのもの**で、それより前の
    段は各カラムを先頭から切り出した写しになる．

    2 つめの値は**図の中の段階**（Issue #125）で ``(カラム番号, 単位数)``．
    ブロック境界だけでは図の内部で切れないので、そのカラムの最後のブロックを
    ``upto(n)`` で途中まで（矢印 n 本目まで・ノード n 個目まで）に差し替える．
    地の文は従来どおり累積したまま、**図だけがその段階の姿になる**．

    **カラム番号を持つ**のは、左右どちらにも図があるとき「最初に見つかった図」
    では取り違えるため（左の図が常に選ばれてしまう）．

    段はどれも**最終的なカラム構成**で描く．カラム区切りより前の段だけ単一カラムに
    すると、レイアウトが段ごとに変わって行頭の位置が動いてしまう．そのため
    marks が短い（＝その時点で存在しなかった）カラムは空として補う．

    タイトル・レイアウト・ディレクティブは全段で同じ．発表者ノートは最終段だけに
    残す——段の集まりで 1 つの話なので、発表者ビューに同じ原稿が何度も出ても
    読みにくいだけ．
    """
    if not marks:
        return [slide]
    # ``Slide.columns`` は**空リストが「単一カラム」の意味**（``None`` は取らない）．
    # parser の add_block も ``s.columns[-1] if s.columns else s.blocks`` と
    # 同じ判定をしており，ここだけ ``is not None`` にすると食い違う．
    cols = slide.columns if slide.columns else [slide.blocks]
    out: list[Slide] = []
    for mark, partial in marks:
        counts = list(mark) + [0] * (len(cols) - len(mark))
        sliced = [list(c[:k]) for c, k in zip(cols, counts)]
        if partial is not None:
            ci, n = partial
            if ci < len(sliced) and sliced[ci]:
                last = sliced[ci][-1]
                if hasattr(last, "upto"):
                    sliced[ci][-1] = last.upto(n)
        # 浅いコピーで足りる——``title_deltas`` は ``int | None``，
        # ``directives`` の値は ``int | str | bool`` で，どれも不変
        # （``ir.Slide`` の注釈が正）．``blocks`` は下でスライスした新しいリスト．
        step = Slide(title=slide.title,
                     title_deltas=list(slide.title_deltas),
                     layout=slide.layout,
                     directives=dict(slide.directives))
        if slide.columns:
            step.columns = sliced
        else:
            step.blocks = sliced[0]
        out.append(step)
    out.append(slide)
    return out


# セル末尾の色指定 "… {accent2}"．**色名らしい語だけ**を対象にする
# （"{n} 個" のような式まで拾うと、書けるものを勝手に減らすことになる）．
_RE_CELL_FILL = re.compile(r"^(.*?)\s*\{\s*(#?[A-Za-z][\w-]*|#[0-9A-Fa-f]{3,6})\s*\}$")


def _split_cell_fill(cell: str) -> tuple[str, str | None]:
    """セルの末尾に書かれた ``{色}`` を剥がして (中身, 色名) を返す（§5.4）．

    色として解釈できない ``{…}`` は**文字のまま残す**——式や記号を壊さないため．
    色名らしいのに解決できないものはタイポとみなして止める（``parse_color``）．
    """
    m = _RE_CELL_FILL.match(cell)
    if not m:
        return cell, None
    name = m.group(2)
    parse_color(name)          # 綴り違いはここで止まる
    return m.group(1).strip(), name


def _split_row(s: str) -> list[str]:
    """Markdown テーブル 1 行をセル列へ分割する（前後の "|" は除去）．"""
    s = s.strip()
    if s.startswith("|"):
        s = s[1:]
    if s.endswith("|"):
        s = s[:-1]
    return [c.strip() for c in s.split("|")]


def _parse_aligns(sep_row: str) -> list[Align]:
    """区切り行のコロンから各列の水平寄せを解析する．

    各セルの先頭・末尾のコロンで判定する：
        ":--:" → "center" / "--:" → "right" / ":--" または "---" → "left"．
    コロンが 1 つも無ければ「指定なし」として空リストを返し，既存テーブルの
    左寄せ挙動を回帰させない（render は空/範囲外を左寄せとして触らない）．
    """
    aligns: list[Align] = []
    has_colon = False
    for cell in _split_row(sep_row):
        c = cell.strip()
        left = c.startswith(":")
        right = c.endswith(":")
        if left or right:
            has_colon = True
        if left and right:
            aligns.append("center")
        elif right:
            aligns.append("right")
        else:
            aligns.append("left")
    return aligns if has_colon else []


# ---------------------------------------------------------------- 画像（§5.9）

def _parse_length(s: str) -> Length | None:
    """"70%" / "8cm" / "300pt" / "2in" / "150px" / 素の数値 を Length へ．

    "%" は割合（render がセグメント比で解決）．単位付きは EMU へ換算（Length("emu", …)）．
    単位無しの素の数値は px 扱い．空文字は None．不正値は ValueError．
    """
    s = (s or "").strip().lower()
    if not s:
        return None
    if s.endswith("%"):
        return Length("percent", _to_float(s[:-1], "size"))
    if s.endswith('"'):  # インチの別表記
        return Length("emu", _to_float(s[:-1], "size") * _EMU_PER["in"])
    for suf, factor in _EMU_PER.items():
        if s.endswith(suf):
            return Length("emu", _to_float(s[: -len(suf)], "size") * factor)
    return Length("emu", _to_float(s, "size") * _EMU_PER["px"])  # 単位無し＝px


def _parse_crop(s: str) -> Crop | None:
    """"x,y,w,h"（残す矩形）を Crop へ．既定 px，各値に "%" を付けると割合．

    4 値必須．"%" は全値に付けるか全く付けないか（混在は不可）．不正は ValueError．
    """
    s = (s or "").strip()
    if not s:
        return None
    parts = [p.strip() for p in s.split(",")]
    if len(parts) != 4:
        raise ValueError(f"crop expects 4 values 'x,y,w,h', got {s!r}")
    pct = [p.endswith("%") for p in parts]
    if any(pct) and not all(pct):
        raise ValueError(f"crop values must be all px or all %: {s!r}")
    unit: Literal["px", "percent"] = "percent" if all(pct) else "px"
    vals = [_to_float(p[:-1] if p.endswith("%") else p, "crop") for p in parts]
    return Crop(unit, *vals)


def _to_float(s: str, what: str) -> float:
    try:
        return float(s)
    except ValueError:
        raise ValueError(f"invalid {what} value: {s!r}")


def _apply_image_opt(img: Image, key: str, val: str) -> None:
    """画像オプション 1 件（key=val / key: val）を Image へ反映する．"""
    key = key.strip().lower()
    val = val.strip()
    if key == "src":
        img.src = val
    elif key == "width":
        img.width = _parse_length(val)
    elif key == "height":
        img.height = _parse_length(val)
    elif key == "crop":
        img.crop = _parse_crop(val)
    elif key == "align":
        v = val.lower()
        if v not in _ALIGNS:
            raise ValueError(f"invalid align: {val!r} (left|center|right)")
        img.align = v
    elif key == "fit":
        v = val.lower()
        if v not in _FITS:
            raise ValueError(f"invalid fit: {val!r} (contain|fill)")
        img.fit = v
    elif key == "caption":
        img.caption = val
    elif key == "overflow":
        v = val.lower()
        if v not in ("true", "false"):
            raise ValueError(f"invalid overflow: {val!r} (true|false)")
        img.overflow = (v == "true")
    else:
        raise ValueError(f"unknown image option: {key!r}")


def _validate_image(img: Image) -> None:
    """Image のキー間の組み合わせ制約を検証する（単一キーは _apply_image_opt）．"""
    if img.overflow and img.width is None and img.height is None:
        raise ValueError(
            "overflow: true requires an explicit width and/or height "
            "(without a size the image is inscribed in the band and never "
            "overflows)")


def _parse_image_shorthand(alt: str, src: str, opts: str | None) -> Image:
    """"![alt](src){opts}" ショートハンドを Image へ．

    opts は空白区切りの "key=value"（crop の値はカンマ区切りなので空白では割らない）．
    alt は caption に採用する（opts に caption があればそちらを優先）．
    """
    img = Image(src=src.strip())
    if alt and alt.strip():
        img.caption = alt.strip()
    for tok in (opts or "").split():
        if "=" not in tok:
            raise ValueError(
                f"invalid image option (expected key=value): {tok!r} — "
                "shorthand options are space-separated, so a caption cannot "
                "contain spaces; use the alt text ![caption](...) or the "
                "```image 'caption:' line instead")
        k, v = tok.split("=", 1)
        _apply_image_opt(img, k, v)
    if not img.src:
        raise ValueError("image requires a source path")
    _validate_image(img)
    return img


def _parse_image_block(text: str) -> Image:
    """```image フェンス（"key: value" 行）を Image へ（```flow と同じ発想）．"""
    img = Image(src="")
    for raw in text.split("\n"):
        line = raw.strip()
        if not line or line.startswith("#"):
            continue  # 空行・コメント行は無視
        if ":" not in line:
            raise ValueError(f"invalid image line (expected 'key: value'): {line!r}")
        k, v = line.split(":", 1)
        _apply_image_opt(img, k, v)
    if not img.src:
        raise ValueError("image block requires 'src:'")
    _validate_image(img)
    return img


def _parse_arrow_block(text: str) -> Arrow:
    """``` ```arrow ``` ブロックを Arrow へ解釈する（§5.16）．

    ``direction:`` は**必須**——向きの無い矢印は描きようが無い．
    ``width:`` / ``height:`` / ``color:`` は任意で、語彙は ``` ```image ``` および
    行内装飾と共通（``_parse_length`` / ``parse_color``）．新しい書き方を増やさない．
    知らないキー・知らない値はタイポとみなしてエラーで止める（他のフェンスと同じ）．

    同じキーを 2 回書いたら**後に書いたほうが残る**．``` ```image ``` と同じ扱いで、
    フェンスの中のキーはどれもそう動く（ここだけエラーにすると規則が二重になる）．
    """
    direction: str | None = None
    width = height = None
    color: str | None = None
    for raw in text.splitlines():
        line = raw.strip()
        if not line:
            continue
        if ":" not in line:
            raise ValueError(
                f"arrow block: expected 'key: value', got {line!r}")
        k, v = line.split(":", 1)
        k, v = k.strip().lower(), v.strip()
        if k == "direction":
            if v.lower() not in ARROW_DIRECTIONS:
                raise ValueError(
                    f"arrow block: unknown direction {v!r} "
                    f"({' | '.join(ARROW_DIRECTIONS)})")
            direction = v.lower()
        elif k == "width":
            width = _parse_length(v)
        elif k == "height":
            height = _parse_length(v)
        elif k == "color":
            # 色は行内装飾と同じ語彙．**検証して正規化してから** IR へ入れる．
            kind, value = parse_color(v)
            color = value if kind == "theme" else "#" + value
        else:
            raise ValueError(
                f"arrow block: unknown key {k!r} "
                f"(direction | width | height | color)")
    if direction is None:
        raise ValueError(
            "arrow block requires 'direction:' "
            f"({' | '.join(ARROW_DIRECTIONS)})")
    # direction は上で ARROW_DIRECTIONS に含まれることを確かめてあるが、
    # mypy は str から Literal への絞り込みを追えない．
    return Arrow(direction=cast(ArrowDirection, direction),
                 width=width, height=height, color=color)


def _apply_directive(slide: Slide, key: str, value: str, lineno: int) -> None:
    """HTML コメント由来のディレクティブを Slide へ反映する．

    キー名はハイフンをアンダースコアへ正規化する
    （@autonum-color → autonum_color）．未知のキーはタイポの可能性が高いので
    黙殺せずエラーにする（v0.7 で改名した旧名称は新名称を案内する）．
    """
    norm = key.replace("-", "_")
    if norm == "col":
        # "<!-- @col -->" と "<!-- @col: arrow -->" は _RE_COL が先に拾う．
        # ここへ来るのは値の綴りが違うときだけ．
        raise ValueError(
            f"invalid @col value at line {lineno} "
            f"(write '<!-- @col -->' or '<!-- @col: arrow -->')")
    if norm == "title_slide":
        raise ValueError(
            f"@title-slide takes no value at line {lineno} "
            f"(write '<!-- @title-slide -->')")
    if norm in _RENAMED_DIRECTIVES:
        raise ValueError(
            f"@{key} was renamed in v0.7; use {_RENAMED_DIRECTIVES[norm]} "
            f"(line {lineno})")
    if norm not in _KNOWN_DIRECTIVES:
        known = ", ".join("@" + k.replace("_", "-") for k in sorted(_KNOWN_DIRECTIVES))
        raise ValueError(
            f"unknown directive @{key} at line {lineno} "
            f"(known directives: @col, @title-slide, {known})")
    val: object = value
    if norm in _INT_DIRECTIVES:
        try:
            val = int(value)
        except ValueError:
            val = value  # 数値でなければ文字列のまま保持（堅牢性）．
    elif norm == "overflow":
        # スライド単位の overflow（表・画像共通）．画像ブロックの overflow: と同じく
        # true/false のみ受理し，それ以外は行番号付きでエラーにする．
        v = value.strip().lower()
        if v not in ("true", "false"):
            raise ValueError(
                f"invalid @overflow value {value!r} at line {lineno} (true|false)")
        val = (v == "true")

    slide.directives[norm] = val

    # @layout はスライドのレイアウト番号を直接上書きする．
    if norm == "layout" and isinstance(val, int):
        if slide.directives.get("title_slide"):
            raise ValueError(
                f"@layout conflicts with @title-slide at line {lineno} "
                f"(use one or the other)")
        slide.layout = val


# ---------------------------------------------------------------- 行頭マーカー

def _split_size(content: str) -> tuple[int | None, str]:
    """本文先頭の相対サイズトークン "{+1}" を剥がして (段数, 残りの本文) を返す．

    トークンが無ければ (None, content)．`None` は「未指定（スライド既定に従う）」を
    意味し，render 側でスライドの @body-size を継承する．

    符号は省略可（"{2}" ＝ "+2"）．"{+0}" / "{-0}" は int 化で 0 となり "{0}" と
    同義（render 側で「テーマ既定に固定」＝スライド既定を無効化）になる．
    """
    m = _RE_SIZE.match(content)
    if m:
        return int(m.group(1)), m.group(2).strip()
    return None, content


def _split_tokens(content: str) -> tuple[int | None, bool, str | None, str]:
    """本文先頭の行レベルトークンを剥がして (段数, 枠, 枠の色, 残りの本文) を返す．

    置き場は 1 か所（行頭マーカーの直後）で、``{+1}`` と ``{box}`` の**どちらが先でも
    受ける**——書く側に順序を覚えさせる理由が無い。同じトークンを 2 回書いても
    最後の値が残るだけで、エラーにはしない。

    知らない綴り（``{boxx}``）は**本文として残す**。``{+x}`` が今そうなっているのと
    同じ扱いで、行頭トークンだけディレクティブ流のエラーにすると規則が二重になる
    （SYNTAX.md にその旨を書いてある）。
    """
    delta: int | None = None
    boxed = False
    color: str | None = None
    while True:
        m = _RE_SIZE.match(content)
        if m:
            delta = int(m.group(1))
            content = m.group(2).strip()
            continue
        m = _RE_BOX.match(content)
        if m:
            boxed = True
            if m.group(1):
                # 色名はここで**検証して正規化する**——綴り違いは黙って既定色に
                # 落とさず止める．正規化まで済ませるのは Span.color と同じ理由で、
                # "#f00" と "#F00" を別物として IR に残さないため（§5.13）．
                kind, value = parse_color(m.group(1))
                color = value if kind == "theme" else "#" + value
            content = m.group(2).strip()
            continue
        return delta, boxed, color, content


def _split_br(text: str) -> tuple[str, list[int | None]]:
    """<br> を行内改行（\\v）へ変換し，各セグメント先頭のサイズトークンを剥がす．

    戻り値は (トークンを除いたテキスト, セグメントと同じ長さの段数リスト)．
    先頭セグメントの段数は **None 固定**——本文行では行頭トークンを
    ``_split_size`` が先に取っており，タイトルでは呼び出し側が入れ直す
    （DESIGN.md §5.8．格納先が違う理由は ir.Line.seg_deltas の説明）．

    ``_RE_BR`` は前後の空白ごと置換するので，セグメントの先頭に空白は残らない．
    """
    segs = _RE_BR.sub("\v", text).split("\v")
    deltas: list[int | None] = [None]
    out = [segs[0]]
    for seg in segs[1:]:
        delta, rest = _split_size(seg)
        deltas.append(delta)
        out.append(rest)
    return "\v".join(out), deltas


# 行内装飾（§5.13）．左から順に食い，マッチしない部分は素のテキストになる．
# ``[表示](url)`` は画像ショートハンド（``![…](…)``）より後で評価されるので衝突しない．
_RE_INLINE = re.compile(
    r"\*\*(?P<bold>.+?)\*\*"
    r"|`(?P<code>[^`]+)`"
    r"|\[(?P<ltext>[^\[\]]*)\]\((?P<url>[^()\s]*)\)"
    r"|\[(?P<ctext>[^\[\]]*)\]\{(?P<color>[^{}]*)\}"
    r"|\^(?P<sup>[^\^\s]+)\^"
    r"|~(?P<sub>[^~\s]+)~"
)


@dataclass(frozen=True)
class _Marks:
    """``[…]{色}`` や ``**…**`` の内側へ継承する装飾．

    ``dict`` で持つとキーの綴り違いを mypy が拾えない——``Span(**marks)`` は
    キーワード展開なので，間違えても実行時まで分からない．
    """

    bold: bool = False
    mono: bool = False
    color: str | None = None
    link: str | None = None
    script: Literal["sup", "sub"] | None = None


def _span(text: str, segment: int, m: _Marks) -> Span:
    """継承した装飾を載せた Span を作る（フィールドの対応はここ 1 か所）．"""
    return Span(text=text, segment=segment, bold=m.bold, mono=m.mono,
                color=m.color, link=m.link, script=m.script)


def _spans_in(text: str, segment: int, base: _Marks) -> list[Span]:
    """1 セグメントを Span 列へ分解する（``base`` は外側から継承する装飾）．

    ``[…]{色}`` と ``[…](url)`` の中身は**再帰的に解釈する**——
    ``[**赤い強調**]{red}`` のように重ねて書けるほうが自然で，
    「色を付けたら太字にできない」という説明を増やさずに済む．
    """
    out: list[Span] = []
    pos = 0
    for m in _RE_INLINE.finditer(text):
        if m.start() > pos:
            out.append(_span(text[pos:m.start()], segment, base))
        if m.group("bold") is not None:
            out += _spans_in(m.group("bold"), segment, replace(base, bold=True))
        elif m.group("code") is not None:
            out.append(_span(m.group("code"), segment, replace(base, mono=True)))
        elif m.group("url") is not None:
            out += _spans_in(m.group("ltext"), segment,
                             replace(base, link=m.group("url")))
        elif m.group("color") is not None:
            # 色名はここで**検証して正規化する**（綴り違いは黙って既定色にせず止める）．
            # 正規化しないと "#f00" と "#F00" が別物として IR に入り，render で
            # もう一度同じ文字列を解き直すことになる（Issue #105 のレビュー指摘）．
            kind, value = parse_color(m.group("color"))
            name = value if kind == "theme" else "#" + value
            out += _spans_in(m.group("ctext"), segment, replace(base, color=name))
        elif m.group("sup") is not None:
            out.append(_span(m.group("sup"), segment, replace(base, script="sup")))
        else:
            out.append(_span(m.group("sub"), segment, replace(base, script="sub")))
        pos = m.end()
    if pos < len(text):
        out.append(_span(text[pos:], segment, base))
    return out


# 等幅ブロックの中で受ける記法は ``[語]{色}`` **だけ**（Issue #162）．
_RE_CODE_COLOR = re.compile(r"\[(?P<ctext>[^\[\]]*)\]\{(?P<color>[^{}]+)\}")


def _code_color_spans(text: str) -> tuple[str, list[Span]]:
    """コード行の ``[語]{色}`` だけを解釈して (素のテキスト, Span 列) を返す．

    **色名として解決できない ``{…}`` は文字のまま残す**——``[133.69.130.4]{n}``
    のようなふつうの角括弧はコードにふつうに現れる．壊さないほうが大事．

    ``**強調**`` や `` `等幅` `` は解釈しない．もう等幅なので意味が無く、
    「書いたまま出る」という約束を色以外では守れる．
    """
    out: list[Span] = []
    pos = 0
    for m in _RE_CODE_COLOR.finditer(text):
        try:
            kind, value = parse_color(m.group("color"))
        except Exception:
            continue                    # 色名でないなら、ただの角括弧
        if m.start() > pos:
            out.append(Span(text=text[pos:m.start()], segment=0))
        name = value if kind == "theme" else "#" + value
        out.append(Span(text=m.group("ctext"), segment=0, color=name))
        pos = m.end()
    if not any(s.color for s in out):
        return text, []                 # 色がひとつも無ければ従来どおり
    if pos < len(text):
        out.append(Span(text=text[pos:], segment=0))
    return "".join(s.text for s in out), out


def _parse_spans(text: str) -> tuple[str, list[Span]]:
    """行内装飾を解釈し (装飾記号を除いたテキスト, Span 列) を返す（§5.13）．

    装飾がまったく無ければ **Span 列は空**で返す——従来どおり 1 run で書く経路に
    落ちるので，装飾を使わない原稿の出力はこの変更で 1 ビットも変わらない．

    セグメント（``\\v`` 区切り）ごとに解釈し，各 Span はどのセグメントの
    ものかを覚える．**1 セグメントが複数 run に割れる**ので，
    相対サイズを付ける相手を位置ではなくこの値で決める必要がある．
    """
    all_spans: list[Span] = []
    plain: list[str] = []
    for i, seg in enumerate(text.split("\v")):
        spans = _spans_in(seg, i, _Marks())
        plain.append("".join(s.text for s in spans))
        all_spans.extend(spans)
    decorated = any(s.bold or s.mono or s.color or s.link or s.script
                    for s in all_spans)
    return "\v".join(plain), (all_spans if decorated else [])


def parse_content_line(raw: str) -> Line | None:
    """1 行を行頭マーカー規則（DESIGN.md §5.3）に従って Line へ変換する．

    **公開関数**．図の note(top)/note(bottom) を地の文として解釈するため
    render からも呼ぶ（Issue #129）——本文行と同じ解釈でなければならず、
    行頭マーカーの判定を二重に持つと必ずずれる．

    インデント（半角スペース 2 つ＝1 レベル）でネスト深さを決める．

    **箇条書きマーカーだけの行は空の段落になる**（"- " だけの行＝1 行空ける．
    Issue #82）．他の行種はマーカー除去後に空なら None（＝行を作らない）．
    素の空行は呼び出し側が段落区切りとして先に読み飛ばすのでここへは来ない．

    各行種のマーカー直後・本文直前に相対サイズトークン "{+1}"/"{-2}" を置ける
    （DESIGN.md §5.8）．トークンは本文から除去し Line.size_delta へ格納する．
    """
    # インデント量からレベルを算出（タブは 1 スペース換算）．
    expanded = raw.replace("\t", " ")
    indent = len(expanded) - len(expanded.lstrip(" "))
    level = indent // 2
    s = expanded.strip()

    if not s:
        return None

    def _mk(text: str, **kw: Any) -> Line | None:
        """本文が空（マーカー／サイズトークンだけの行）なら Line を作らず None．

        空を段落として残すのは**箇条書きマーカーだけ**（``_mk_bullet``）．採番行を
        空で残すと番号を 1 つ消費するだけで，空けたい 1 行は得られない．

        本文中の <br> はタイトルと同じ規則で行内改行（\v）へ変換する．
        render 側は本文をそのまま段落 text へ渡すため，python-pptx が
        "\v" を段落内改行（<a:br/>）として出力する．"""
        text, seg_deltas = _split_br(text)
        text, spans = _parse_spans(text)
        return (Line(text=text, level=level, seg_deltas=seg_deltas,
                     spans=spans, **kw)
                if text else None)

    def _mk_bullet(text: str, size_delta: int | None,
                   boxed: bool = False, box_color: str | None = None) -> Line:
        """箇条書き行を Line にする．**本文が空でも段落を作る**（Issue #82）．

        マーカーだけの行は「1 行空ける」指示として使う．表紙の著者欄やセクション扉の
        ように記号の出ない枠で，行の塊を分けるのに要る．front matter の
        ``affiliation`` では ``- "{-2} "`` と書けば空段落が残っていたのに，本文では
        捨てられていた——同じことを二重に実装した結果のずれで，本文側だけが
        取りこぼしていた．

        代用として ``- <br>`` は残るが**1 行多く空く**（段落 1 行＋``a:br`` の 2 行目）．
        空けたいのは 1 行なので，空の段落そのものを作れる必要がある．
        """
        body, seg_deltas = _split_br(text)
        body, spans = _parse_spans(body)
        return Line(text=body, level=level, kind="bullet", boxed=boxed,
                    box_color=box_color, size_delta=size_delta,
                    seg_deltas=seg_deltas, spans=spans)

    # マーカーだけの行 → 空行（Issue #82）．**末尾に空白が無い形を必ず拾うこと**
    # ——多くのエディタは保存時に行末空白を除去するので，"- " しか受けないと空行
    # スペーサが保存した瞬間に壊れる（既定の箇条書きへ落ちて**文字の "-"** が出る）．
    # 下の "- " 判定と分けてあるのは s[2:] の意味を保つため．あちらは「マーカーと
    # 続く空白の 2 文字を除く」意図で，1 文字しかない行を同じ式に通すと意図が濁る．
    if s in ("-", "*"):
        return _mk_bullet("", None)

    # 通常箇条書き："- " / "* "
    if s.startswith("- ") or s.startswith("* "):
        delta, boxed, box_color, text = _split_tokens(s[2:].strip())
        return _mk_bullet(text, delta, boxed, box_color)

    # 採番行は**書かれていた番号を捨てない**（Issue #107）．render がリストの
    # 先頭の行だけ開始番号として使う——PowerPoint の自動採番はプレースホルダごとに
    # 1 から数え直すので，開始番号を渡せないと 2 カラムに割った採番リストの
    # 右カラムが 1. に戻る．先頭だけ効かせるのは CommonMark と同じ規則で，
    # "1. 1. 1." と書けば 1・2・3 になる従来の書き方もそのまま動く．

    # 連番："1. 2. 3." → arabicPeriod
    m = _RE_ORDERED.match(s)
    if m:
        delta, boxed, box_color, text = _split_tokens(m.group(2).strip())
        return _mk(text, kind="autonum", num_style="arabicPeriod",
                   num_start=int(m.group(1)), size_delta=delta, boxed=boxed,
                   box_color=box_color)

    # 丸括弧："(1) (2)" → arabicParenBoth（"(1)" 表記を忠実に再現）
    m = _RE_PAREN.match(s)
    if m:
        delta, boxed, box_color, text = _split_tokens(m.group(2).strip())
        return _mk(text, kind="autonum", num_style="arabicParenBoth",
                   num_start=int(m.group(1)), size_delta=delta, boxed=boxed,
                   box_color=box_color)

    # 丸数字："①②③ …" → circleNumDbPlain（番号文字は除去）
    if s[0] in CIRCLED_DIGITS:
        delta, boxed, box_color, text = _split_tokens(s[1:].lstrip())
        return _mk(text, kind="autonum", num_style="circleNumDbPlain",
                   num_start=CIRCLED_DIGITS.index(s[0]) + 1, size_delta=delta,
                   boxed=boxed, box_color=box_color)

    # 矢印："→ …" → 行頭記号なし（no_bullet 相当）．"→" は本文に残す
    # （結論・補足行の視覚的な導線として表示する）．トークンは "→" の後ろに置く．
    # 他の行種と同様，"→ 本文" へ空白を正規化する（トークン有無で挙動を変えない）．
    if s.startswith(ARROW):
        delta, boxed, box_color, rest = _split_tokens(s[len(ARROW):].lstrip())
        text = f"{ARROW} {rest}" if rest else ARROW
        text, seg_deltas = _split_br(text)
        text, spans = _parse_spans(text)
        return Line(text=text, level=level, kind="plain", size_delta=delta,
                    seg_deltas=seg_deltas, spans=spans, boxed=boxed,
                    box_color=box_color)

    # 上記以外 → 既定の箇条書き（インデントに応じたレベル）
    delta, boxed, box_color, text = _split_tokens(s)
    return _mk(text, kind="bullet", size_delta=delta, boxed=boxed,
               box_color=box_color)


# ---------------------------------------------------------------- 自己検証

# パッケージ内 import は相対のみなので `python3 -m md2pptx.parser` で実行する
# （cli.py / thmx2pptx.py の自己検証も同じ流儀）．
if __name__ == "__main__":
    import io

    sample = """---
theme: OfficeTheme.pptx
output: out.pptx
slide_number: true
default_autofit: true
title: |
  md2pptx
  Markdown でつくるスライド
subtitle: ― テーマ駆動のスライド生成 ―
author: md2pptx demo
affiliation:
  - Markdown ＋ PowerPoint テーマ → pptx
  - Python / python-pptx / PyYAML
---

## Background

- スライドは体裁に時間を取られがち
  - 配色・フォントをそろえるのが面倒
  - テキストとレイアウトが密結合

## Features
<!-- @autonum-color: tx1 -->

① テーマの配色・フォントを継承
② 表・フロー図・2カラムに対応
③ はみ出し防止の自動縮小

## モジュール構成
<!-- @autofit: 90 -->

1. parser.py
  - Markdown を中間表現（IR）へ変換
2. render.py
  - IR を pptx に描画

→ 色やフォントはコードに持たず、テーマに委ねる

## ブロック混在

導入：Line 以外のブロックもダンプ対象。

| 課題 | 対応 |
|---|---:|
| デザイン | テーマに委譲 |

```flow
[md] -変換-> [pptx]
```

→ 表・フロー図が Line 決め打ちで落ちないことの確認を兼ねる
"""

    def _dump_block(b: Block) -> str:
        """ブロックを 1 行で表す．blocks は Line 以外も持つので型で分岐する．"""
        if isinstance(b, Line):
            return (f"Line(kind={b.kind!r} level={b.level} "
                    f"num_style={b.num_style!r} text={b.text!r})")
        if isinstance(b, Table):
            return (f"Table(cols={len(b.header)} rows={len(b.rows)} "
                    f"aligns={b.aligns!r})")
        if isinstance(b, Flow):
            return (f"Flow(direction={b.direction!r} nodes={len(b.nodes)} "
                    f"edges={len(b.edges)} caption={b.caption!r})")
        if isinstance(b, Image):
            return f"Image(src={b.src!r} align={b.align!r} overflow={b.overflow!r})"
        return f"{type(b).__name__}({b!r})"

    deck = parse(sample)
    buf = io.StringIO()
    ts = deck.title_slide
    print("=== meta ===", file=buf)
    print(deck.meta, file=buf)
    print("=== title_slide ===", file=buf)
    if ts is None:
        print("(none)", file=buf)
    else:
        print(f"title={ts.title!r}", file=buf)
        print(f"subtitle={ts.subtitle!r}", file=buf)
        print(f"author={ts.author!r}", file=buf)
        print(f"affiliation={ts.affiliation!r}", file=buf)
    print(f"=== slides ({len(deck.slides)}) ===", file=buf)
    for si, sl in enumerate(deck.slides):
        print(f"[slide {si}] title={sl.title!r} layout={sl.layout} "
              f"directives={sl.directives}", file=buf)
        for b in sl.blocks:
            print(f"    {_dump_block(b)}", file=buf)

    with open("/tmp/parser_chk.txt", "w", encoding="utf-8") as f:
        f.write(buf.getvalue())
    print("wrote /tmp/parser_chk.txt")
