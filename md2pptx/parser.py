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
from typing import Literal

import yaml

from .ir import (
    Align, Crop, Deck, Flow, Image, Length, Line, Slide, Table, TitleSlide,
)
from .flow import parse_flow as _parse_flow


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
# カラム区切り（「2つのコンテンツ」レイアウト）．値を取らない指示．
_RE_COL = re.compile(r"^<!--\s*@col\s*-->$")
# 1 行 HTML コメント（ディレクティブ以外のメモ等．無視する）．
_RE_COMMENT = re.compile(r"^<!--.*-->$")
# Markdown テーブルの区切り行（例 "| --- | :--: |"）．ヘッダ行の直後に現れる．
_RE_TABLE_SEP = re.compile(r"^\|?\s*:?-{1,}:?\s*(\|\s*:?-{1,}:?\s*)*\|?$")
# タイトル・本文行内の明示改行マーカー（<br> / <br/>）．"\v"（行内改行）へ変換する．
_RE_BR = re.compile(r"\s*<br\s*/?>\s*")

# 相対フォントサイズトークン．マーカー直後・本文直前の "{+1}"/"{-2}"/"{0}"．
# 符号は省略可（"{2}" は "+2" と同義）．render がテーマ基準で実サイズへ換算する．
_RE_SIZE = re.compile(r"^\{\s*([+-]?\d+)\s*\}\s*(.*)$")

# 整数として解釈するディレクティブキー（正規化後の名前）．
# body_size はスライド既定の相対フォントサイズ段数（@body-size）．
_INT_DIRECTIVES = {"layout", "autofit", "body_size"}

# 受理するディレクティブキー（正規化後の名前）．未知のキーはタイポの可能性が
# 高いので黙殺せずエラーにする（§5.6）．@col は値を取らない専用形式（_RE_COL）．
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
    "theme", "output", "slide_number", "default_autofit",
    "title", "subtitle", "author", "affiliation",
}

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
    deck = Deck(meta=meta)
    deck.title_slide = _build_title_slide(meta)
    deck.slides, title_notes = _parse_body(
        body, body_offset, has_title_slide=deck.title_slide is not None)
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


def _build_title_slide(meta: dict) -> TitleSlide | None:
    """フロントマターからタイトルスライドを構築する（title が無ければ None）．"""
    if not meta.get("title"):
        return None

    title = meta.get("title")
    if isinstance(title, str):
        # 複数行タイトル（YAML ブロックスカラー）の末尾改行を落とす．
        title = title.rstrip("\n")

    # 副題・著者・所属も本文行と同じ相対サイズトークン "{-1}"/"{+1}" を先頭に置ける．
    # トークンは本文から剥がし，段数を IR の *_delta へ格納する（render が換算）．
    subtitle_delta, subtitle = _split_size_opt(meta.get("subtitle"))
    author_delta, author = _split_size_opt(meta.get("author"))

    affiliation_raw = meta.get("affiliation") or []
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


def _split_size_opt(value) -> tuple[int | None, str | None]:
    """front matter 値（None 可）の先頭相対サイズトークンを剥がして (段数, 文字列) を返す．

    None はそのまま (None, None)．トークン判定は文字列のみ対象とし，数値等
    （YAML が int/float で読んだ値）は素直に文字列化して段数なしで返す．
    """
    if value is None:
        return None, None
    if not isinstance(value, str):
        return None, str(value)
    return _split_size(value)


# ---------------------------------------------------------------- 本文

def _parse_body(body: str, body_offset: int = 0,
                has_title_slide: bool = False) -> tuple[list[Slide], str | None]:
    """本文をスライド列へ分割し，各行を IR ブロックへ変換する．

    body_offset はフロントマターが消費したファイル行数（エラー報告の行番号を
    ファイル先頭基準へ換算するために使う）．has_title_slide はフロントマター
    由来のタイトルスライドの有無（本文開始前の ```note の宛先判定に使う）．

    Returns:
        (slides, title_notes)．title_notes は本文開始前に現れた ```note の
        内容（タイトルスライドの発表者ノート）．無ければ None．
    """
    slides: list[Slide] = []
    current: Slide | None = None
    title_notes: list[str] = []

    def ensure_slide() -> Slide:
        """直前にスライド開始マーカーが無いまま本文が来た場合のフォールバック．"""
        nonlocal current
        if current is None:
            current = Slide()
        return current

    def add_block(b) -> None:
        """ブロックを現在のカラム（多カラム時）または blocks へ追加する．"""
        s = ensure_slide()
        (s.columns[-1] if s.columns else s.blocks).append(b)

    lines = body.split("\n")
    n = len(lines)
    i = 0
    while i < n:
        raw = lines[i]
        stripped = raw.strip()
        lineno = body_offset + i + 1  # ファイル先頭基準の 1 始まり行番号

        # --- スライド分割マーカー ---------------------------------
        m = _RE_HEADING.match(raw.lstrip())
        if m:
            hashes, htext = m.group(1), m.group(2)
            # "# 見出し"（H1）はセクションスライド（レイアウト2），
            # "## 見出し" はコンテンツスライド（レイアウト1）．
            # H3〜H6 は未定義（将来のスライド内小見出し用に予約）．
            if len(hashes) > 2:
                raise ValueError(
                    f"H{len(hashes)} heading is not supported at line {lineno}: "
                    f"{stripped!r} (use '#' for a section slide or '##' for a "
                    f"content slide)")
            # タイトル内の <br> を行内改行（\v）へ変換する．
            htext = _RE_BR.sub("\v", htext)
            if current is not None:
                slides.append(current)
            layout = 2 if len(hashes) == 1 else 1
            current = Slide(title=htext or None, layout=layout)
            i += 1
            continue

        if stripped == "---":
            # 水平線 → タイトルなしスライドを明示的に開始．
            if current is not None:
                slides.append(current)
            current = Slide()
            i += 1
            continue

        # --- カラム区切り（「2つのコンテンツ」）→ 多カラム化（§5.7）----
        if _RE_COL.match(stripped):
            s = ensure_slide()
            if not s.columns:
                s.layout = 3                 # 2つのコンテンツ レイアウト
                s.columns = [s.blocks, []]   # 既存ブロックを左カラムへ
            else:
                s.columns.append([])
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

        # --- フェンスドコードブロック（```flow … ```）→ Flow（§5.5）--
        if stripped.startswith("```"):
            info = stripped[3:].strip().lower()
            j = i + 1
            buf: list[str] = []
            while j < n and lines[j].strip() != "```":
                buf.append(lines[j])
                j += 1
            if info == "flow":
                add_block(_parse_flow("\n".join(buf)))
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
            # flow / image / note 以外のコードブロックは範囲外（無視）．
            i = j + 1  # 閉じフェンスの次へ（無い場合も末尾へ）
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
            add_block(Table(header=header, rows=rows, aligns=aligns))
            i = j
            continue

        # --- 空行は段落区切り（Line は作らない）-------------------
        if not stripped:
            i += 1
            continue

        # --- 本文行 → Line ---------------------------------------
        line = _parse_content_line(raw)
        if line is not None:
            add_block(line)
        i += 1

    if current is not None:
        slides.append(current)

    return slides, ("\n".join(title_notes) if title_notes else None)


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


def _apply_directive(slide: Slide, key: str, value: str, lineno: int) -> None:
    """HTML コメント由来のディレクティブを Slide へ反映する．

    キー名はハイフンをアンダースコアへ正規化する
    （@autonum-color → autonum_color）．未知のキーはタイポの可能性が高いので
    黙殺せずエラーにする（v0.7 で改名した旧名称は新名称を案内する）．
    """
    norm = key.replace("-", "_")
    if norm == "col":
        # 値なしの "<!-- @col -->" は _RE_COL が先に拾う．ここへ来るのは
        # "@col: 2" のような値付きで，カラム区切りとしては不正．
        raise ValueError(
            f"@col takes no value at line {lineno} (write '<!-- @col -->')")
    if norm in _RENAMED_DIRECTIVES:
        raise ValueError(
            f"@{key} was renamed in v0.7; use {_RENAMED_DIRECTIVES[norm]} "
            f"(line {lineno})")
    if norm not in _KNOWN_DIRECTIVES:
        known = ", ".join("@" + k.replace("_", "-") for k in sorted(_KNOWN_DIRECTIVES))
        raise ValueError(
            f"unknown directive @{key} at line {lineno} "
            f"(known directives: @col, {known})")
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


def _parse_content_line(raw: str) -> Line | None:
    """1 行を行頭マーカー規則（DESIGN.md §5.3）に従って Line へ変換する．

    インデント（半角スペース 2 つ＝1 レベル）でネスト深さを決める．
    空行（マーカー除去後に空）は None を返す．

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

    def _mk(text, **kw):
        """本文が空（マーカー／サイズトークンだけの行）なら Line を作らず None．
        マーカー除去後に空の行を IR に入れない（先頭の空行チェックと整合）．

        本文中の <br> はタイトルと同じ規則で行内改行（\v）へ変換する．
        render 側は本文をそのまま段落 text へ渡すため，python-pptx が
        "\v" を段落内改行（<a:br/>）として出力する．"""
        text = _RE_BR.sub("\v", text)
        return Line(text=text, level=level, **kw) if text else None

    # 通常箇条書き："- " / "* "
    if s.startswith("- ") or s.startswith("* "):
        delta, text = _split_size(s[2:].strip())
        return _mk(text, kind="bullet", size_delta=delta)

    # 連番："1. 2. 3." → arabicPeriod
    m = _RE_ORDERED.match(s)
    if m:
        delta, text = _split_size(m.group(2).strip())
        return _mk(text, kind="autonum", num_style="arabicPeriod", size_delta=delta)

    # 丸括弧："(1) (2)" → arabicParenBoth（"(1)" 表記を忠実に再現）
    m = _RE_PAREN.match(s)
    if m:
        delta, text = _split_size(m.group(2).strip())
        return _mk(text, kind="autonum", num_style="arabicParenBoth", size_delta=delta)

    # 丸数字："①②③ …" → circleNumDbPlain（番号文字は除去）
    if s[0] in CIRCLED_DIGITS:
        delta, text = _split_size(s[1:].lstrip())
        return _mk(text, kind="autonum", num_style="circleNumDbPlain", size_delta=delta)

    # 矢印："→ …" → 行頭記号なし（no_bullet 相当）．"→" は本文に残す
    # （結論・補足行の視覚的な導線として表示する）．トークンは "→" の後ろに置く．
    # 他の行種と同様，"→ 本文" へ空白を正規化する（トークン有無で挙動を変えない）．
    if s.startswith(ARROW):
        delta, rest = _split_size(s[len(ARROW):].lstrip())
        text = f"{ARROW} {rest}" if rest else ARROW
        text = _RE_BR.sub("\v", text)
        return Line(text=text, level=level, kind="plain", size_delta=delta)

    # 上記以外 → 既定の箇条書き（インデントに応じたレベル）
    delta, text = _split_size(s)
    return _mk(text, kind="bullet", size_delta=delta)


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

    def _dump_block(b) -> str:
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
