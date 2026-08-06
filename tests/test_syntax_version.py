#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""``syntax:`` による見出しレベルの割り当てを固定するテスト（Issue #99）．

``#`` の意味が変わる変更なので、警告では守れない。旧原稿を新しい割り当てで読むと
**エラーにならず全体が 1 段ずれる**（``# 章の扉`` が表紙になり、``## スライド`` が
章の扉になる）。だから宣言で選ぶ。

既定は 1。旧原稿には ``syntax: 0`` を書き足して使う。既定を 1 にしたので
書き足していない旧原稿は 1 段ずれて読まれるが、``# 章の扉`` は複数あるのが普通で、
**表紙が 2 枚できた時点で止まる**。そこで ``syntax: 0`` を案内する。
"""
from __future__ import annotations

import pytest

from md2pptx.ir import CONTENT_LAYOUT, SECTION_LAYOUT, TITLE_LAYOUT
from md2pptx.parser import parse


def _layouts(src):
    return [s.layout for s in parse(src).slides]


def _fm(syntax=None):
    line = f"syntax: {syntax}\n" if syntax is not None else ""
    return f"---\ntheme: t.pptx\n{line}---\n\n"


_BODY_01 = "# H1\n\n## H2\n"
_BODY_1 = "# H1\n\n## H2\n\n### H3\n"


def test_the_default_is_one():
    """``syntax`` を書かない原稿は新しい割り当てで読む．"""
    assert _layouts(_fm() + _BODY_1) == [
        TITLE_LAYOUT, SECTION_LAYOUT, CONTENT_LAYOUT]


def test_syntax_zero_keeps_the_old_mapping():
    """``syntax: 0`` は従来の割り当て（`#` 章の扉 / `##` スライド）．"""
    assert _layouts(_fm(0) + _BODY_01) == [SECTION_LAYOUT, CONTENT_LAYOUT]


def test_syntax_one_shifts_every_level():
    """``syntax: 1`` は 1 段ずらす（`#` 表紙 / `##` 章の扉 / `###` スライド）．"""
    assert _layouts(_fm(1) + _BODY_1) == [
        TITLE_LAYOUT, SECTION_LAYOUT, CONTENT_LAYOUT]


def test_the_same_source_means_different_things():
    """同じ原稿が両者で別物になる——これが宣言で選ぶ理由．

    警告で済ませられないことを、ここで見えるようにしておく。
    """
    assert _layouts(_fm(0) + _BODY_01) == [SECTION_LAYOUT, CONTENT_LAYOUT]
    assert _layouts(_fm(1) + _BODY_01) == [TITLE_LAYOUT, SECTION_LAYOUT]


@pytest.mark.parametrize("syntax,body,unsupported", [
    (0, "### H3\n", "H3"),      # 従来は H3 が予約
    (None, "#### H4\n", "H4"),  # 既定（1）では 1 段下がる
    (1, "#### H4\n", "H4"),
])
def test_levels_outside_the_mapping_are_reserved(syntax, body, unsupported):
    """割り当ての外のレベルはエラー（将来のスライド内小見出し用に予約）．"""
    with pytest.raises(ValueError, match=f"{unsupported} heading is not supported"):
        parse(_fm(syntax) + body)


def test_the_error_names_the_usable_levels():
    """エラーはその syntax で使えるレベルを示す（何を書けばよいか分かるように）．"""
    with pytest.raises(ValueError, match=r"syntax 1 uses '#' / '##' / '###'"):
        parse(_fm(1) + "#### H4\n")


def test_title_slide_directive_is_rejected_under_syntax_one():
    """``syntax: 1`` では ``@title-slide`` はエラー（``#`` が表紙なので冗長）．

    黙って受けると「表紙の書き方が 2 通り」に戻る。
    """
    with pytest.raises(ValueError, match="not used with syntax 1"):
        parse(_fm(1) + "# 主題\n<!-- @title-slide -->\n")


def test_title_slide_directive_still_works_under_syntax_zero():
    """``syntax: 0`` では ``@title-slide`` は従来どおり効く．"""
    assert _layouts(_fm(0) + "# 主題\n<!-- @title-slide -->\n") == [TITLE_LAYOUT]


def test_a_second_title_slide_stops_and_points_at_syntax_zero():
    """表紙が 2 枚できたら止め、``syntax: 0`` を案内する．

    ``syntax: 0`` を書き忘れた旧原稿がここで捕まる——``# 章の扉`` は複数あるのが
    普通なので、2 枚目で引っかかる。黙って 1 段ずれたデッキを出すより良い。
    """
    old = _fm() + "# 第1部\n\n## スライドA\n\n# 第2部\n"
    with pytest.raises(ValueError) as e:
        parse(old)
    assert "second title slide" in str(e.value)
    assert "syntax: 0" in str(e.value)   # 何を書けばよいかを示すこと
    # syntax: 0 を足せば通る。
    assert _layouts(_fm(0) + "# 第1部\n\n## スライドA\n\n# 第2部\n") == [
        SECTION_LAYOUT, CONTENT_LAYOUT, SECTION_LAYOUT]


def test_syntax_zero_allows_more_than_one_title_slide():
    """``syntax: 0`` では ``@title-slide`` を複数書ける（従来の挙動）．"""
    src = _fm(0) + ("# A\n<!-- @title-slide -->\n\n"
                    "# B\n<!-- @title-slide -->\n")
    assert _layouts(src) == [TITLE_LAYOUT, TITLE_LAYOUT]


def test_an_unknown_syntax_value_stops():
    """未知の値は止める．

    将来の版で増える記法を、古い md2pptx が黙って既定として読むと、
    書き手の意図と違うデッキが出る。
    """
    with pytest.raises(ValueError, match="invalid syntax value"):
        parse(_fm(2) + "# x\n")
    with pytest.raises(ValueError, match="invalid syntax value"):
        parse(_fm("'v1'") + "# x\n")
