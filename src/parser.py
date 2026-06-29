#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Markdown → 中間表現（IR）パーサ（md2pptx Phase 1）．

DESIGN.md §5 の Markdown 記法仕様を解釈し，ir.py のデータクラス
（Deck / TitleSlide / Slide / Line）へ変換する純 Python モジュール．
python-pptx には依存しない（描画は render.py の責務）．

担当範囲（Phase 1）:
    - フロントマター（YAML）分離 → Deck.meta / TitleSlide 生成（§5.1）
    - スライド分割（"## 見出し" / "---" 水平線 / "# 見出し"）（§5.2）
    - 行頭マーカー解釈（"-"/"*"/"1."/丸数字/"(n)"/"→"）（§5.3）
    - スライド単位ディレクティブ（HTML コメント）の収集（§5.6）

表（§5.4）・flow 図（§5.5）は Phase 2/3 で対応するため，ここでは扱わない
（該当行は素朴に Line として扱われる）．
"""
from __future__ import annotations

import re

import yaml

try:  # パッケージ実行・単体実行のどちらでも import できるように
    from .ir import Deck, Line, Slide, Table, TitleSlide
    from .flow import parse_flow as _parse_flow
except ImportError:  # pragma: no cover - 単体実行時のフォールバック
    from ir import Deck, Line, Slide, Table, TitleSlide
    from flow import parse_flow as _parse_flow


# ---------------------------------------------------------------- 定数

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
# タイトル内の明示改行マーカー（<br> / <br/>）．"\v"（行内改行）へ変換する．
_RE_TITLE_BR = re.compile(r"\s*<br\s*/?>\s*")

# 整数として解釈するディレクティブキー（正規化後の名前）．
_INT_DIRECTIVES = {"layout", "autofit"}


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
    meta, body = _split_front_matter(text)
    deck = Deck(meta=meta)
    deck.title_slide = _build_title_slide(meta)
    deck.slides = _parse_body(body)
    return deck


def parse_file(path: str) -> Deck:
    """Markdown ファイルを読み込んで parse() する利便関数．"""
    with open(path, encoding="utf-8") as f:
        return parse(f.read())


# ---------------------------------------------------------------- フロントマター

def _normalize_newlines(text: str) -> str:
    """改行コードを LF に正規化する（CRLF / CR 対策）．"""
    return text.replace("\r\n", "\n").replace("\r", "\n")


def _split_front_matter(text: str) -> tuple[dict, str]:
    """先頭の "---" 〜 "---" を YAML として切り出す．

    Returns:
        (meta, body). フロントマターが無ければ ({}, text)．
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
                return meta, body
    return {}, text


def _build_title_slide(meta: dict) -> TitleSlide | None:
    """フロントマターからタイトルスライドを構築する（title が無ければ None）．"""
    if not meta.get("title"):
        return None

    title = meta.get("title")
    if isinstance(title, str):
        # 複数行タイトル（YAML ブロックスカラー）の末尾改行を落とす．
        title = title.rstrip("\n")

    subtitle = meta.get("subtitle")
    author = meta.get("author")

    affiliation = meta.get("affiliation") or []
    if isinstance(affiliation, str):
        affiliation = [affiliation]
    else:
        affiliation = list(affiliation)

    return TitleSlide(
        title=title,
        subtitle=subtitle,
        author=author,
        affiliation=affiliation,
    )


# ---------------------------------------------------------------- 本文

def _parse_body(body: str) -> list[Slide]:
    """本文をスライド列へ分割し，各行を IR ブロックへ変換する．"""
    slides: list[Slide] = []
    current: Slide | None = None

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

        # --- スライド分割マーカー ---------------------------------
        m = _RE_HEADING.match(raw.lstrip())
        if m:
            hashes, htext = m.group(1), m.group(2)
            # タイトル内の <br> を行内改行（\v）へ変換する．
            htext = _RE_TITLE_BR.sub("\v", htext)
            if current is not None:
                slides.append(current)
            # "# 見出し"（H1）はセクションスライド（レイアウト2），
            # "## 見出し" 以上はコンテンツスライド（レイアウト1）．
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
            _apply_directive(slide, md.group(1), md.group(2))
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
            # flow 以外のコードブロックは Phase 3 範囲外（無視）．
            i = j + 1  # 閉じフェンスの次へ（無い場合も末尾へ）
            continue

        # --- 表（ヘッダ行＋直後の区切り行）→ Table（§5.4）---------
        if "|" in stripped and i + 1 < n and _RE_TABLE_SEP.match(lines[i + 1].strip()):
            header = _split_row(stripped)
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
            add_block(Table(header=header, rows=rows))
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

    return slides


def _split_row(s: str) -> list[str]:
    """Markdown テーブル 1 行をセル列へ分割する（前後の "|" は除去）．"""
    s = s.strip()
    if s.startswith("|"):
        s = s[1:]
    if s.endswith("|"):
        s = s[:-1]
    return [c.strip() for c in s.split("|")]


def _apply_directive(slide: Slide, key: str, value: str) -> None:
    """HTML コメント由来のディレクティブを Slide へ反映する．

    キー名はハイフンをアンダースコアへ正規化する
    （@autonum-color → autonum_color）．未知のキーも素直に格納する．
    """
    norm = key.replace("-", "_")
    val: object = value
    if norm in _INT_DIRECTIVES:
        try:
            val = int(value)
        except ValueError:
            val = value  # 数値でなければ文字列のまま保持（堅牢性）．

    slide.directives[norm] = val

    # @layout はスライドのレイアウト番号を直接上書きする．
    if norm == "layout" and isinstance(val, int):
        slide.layout = val


# ---------------------------------------------------------------- 行頭マーカー

def _parse_content_line(raw: str) -> Line | None:
    """1 行を行頭マーカー規則（DESIGN.md §5.3）に従って Line へ変換する．

    インデント（半角スペース 2 つ＝1 レベル）でネスト深さを決める．
    空行（マーカー除去後に空）は None を返す．
    """
    # インデント量からレベルを算出（タブは 1 スペース換算）．
    expanded = raw.replace("\t", " ")
    indent = len(expanded) - len(expanded.lstrip(" "))
    level = indent // 2
    s = expanded.strip()

    if not s:
        return None

    # 通常箇条書き："- " / "* "
    if s.startswith("- ") or s.startswith("* "):
        return Line(text=s[2:].strip(), level=level, kind="bullet")

    # 連番："1. 2. 3." → arabicPeriod
    m = _RE_ORDERED.match(s)
    if m:
        return Line(text=m.group(2).strip(), level=level,
                    kind="autonum", num_style="arabicPeriod")

    # 丸括弧："(1) (2)" → arabicParenBoth（"(1)" 表記を忠実に再現）
    m = _RE_PAREN.match(s)
    if m:
        return Line(text=m.group(2).strip(), level=level,
                    kind="autonum", num_style="arabicParenBoth")

    # 丸数字："①②③ …" → circleNumDbPlain（番号文字は除去）
    if s[0] in CIRCLED_DIGITS:
        return Line(text=s[1:].lstrip(), level=level,
                    kind="autonum", num_style="circleNumDbPlain")

    # 矢印："→ …" → 行頭記号なし（no_bullet 相当）．"→" は本文に残す
    # （結論・補足行の視覚的な導線として表示する）．
    if s.startswith(ARROW):
        return Line(text=s, level=level, kind="plain")

    # 上記以外 → 既定の箇条書き（インデントに応じたレベル）
    return Line(text=s, level=level, kind="bullet")


# ---------------------------------------------------------------- 自己検証

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
  - MIT License
---

## Background

- スライドは体裁に時間を取られがち
  - 環境構築が困難
  - 版管理が煩雑

## Contributions
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
"""

    deck = parse(sample)
    buf = io.StringIO()
    ts = deck.title_slide
    print("=== meta ===", file=buf)
    print(deck.meta, file=buf)
    print("=== title_slide ===", file=buf)
    print(f"title={ts.title!r}", file=buf)
    print(f"subtitle={ts.subtitle!r}", file=buf)
    print(f"author={ts.author!r}", file=buf)
    print(f"affiliation={ts.affiliation!r}", file=buf)
    print(f"=== slides ({len(deck.slides)}) ===", file=buf)
    for si, sl in enumerate(deck.slides):
        print(f"[slide {si}] title={sl.title!r} layout={sl.layout} "
              f"directives={sl.directives}", file=buf)
        for b in sl.blocks:
            print(f"    Line(kind={b.kind!r} level={b.level} "
                  f"num_style={b.num_style!r} text={b.text!r})", file=buf)

    with open("/tmp/parser_chk.txt", "w", encoding="utf-8") as f:
        f.write(buf.getvalue())
    print("wrote /tmp/parser_chk.txt")
