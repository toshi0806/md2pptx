#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""``<!-- @… -->`` の打ち間違いを止めることを固定するテスト（Issue #198）．

**値を取らない指示の打ち間違いは、黙って捨てられていた．** 値付きの
``<!-- @body-siz: -1 -->`` は止まるのに ``<!-- @stepp -->`` は素通りする——
``_RE_DIRECTIVE`` が**コロン必須**で、そこから外れた ``@…`` が
``_RE_COMMENT`` に拾われるため．

実害は「書いたつもりの指示が無いデッキが、**成功として**出てくる」こと．
``@stepp`` と打てば段階表示が消え、ビルドも検査も通るので、
気づくのは出力を開いたときになる．

``@`` で始まらないコメントは従来どおり黙って捨てる（``<!-- handout-skip -->``
や普通のメモ）．変えたのは ``@`` で始まるものだけ．
"""
from __future__ import annotations

import pytest

from md2pptx.parser import parse

_FM = "---\ntheme: t.pptx\n---\n\n"
# @title-slide は syntax 0 専用（1 では "#" 自体が表紙）．
_FM0 = "---\ntheme: t.pptx\nsyntax: 0\n---\n\n"


def _with(comment, fm=_FM):
    return fm + f"## 見出し\n\n- a\n{comment}\n- b\n"


@pytest.mark.parametrize("comment", [
    "<!-- @stepp -->",
    "<!-- @colx -->",
    "<!-- @title-slid -->",
    "<!-- @handout-skp -->",
])
def test_a_mistyped_valueless_directive_stops_the_build(comment):
    """値を取らない指示の綴り違いは、未知のディレクティブとして止まる．"""
    with pytest.raises(ValueError, match="unknown directive"):
        parse(_with(comment))


def test_the_known_valueless_directives_still_work():
    """``@step`` / ``@col`` / ``@title-slide`` は従来どおり通る．"""
    assert len(parse(_with("<!-- @step -->")).slides) == 2
    assert len(parse(_with("<!-- @col -->")).slides[0].columns) == 2
    assert parse(_FM0 + "# 主題\n<!-- @title-slide -->\n").slides[0].directives \
        == {"title_slide": True}


def test_a_comment_without_the_at_sign_is_still_dropped():
    """``@`` で始まらないコメントは黙って捨てる——**本文にも出さない**．

    ``<!-- handout-skip -->`` は IE2026 の配布版生成が使っていた印で、
    指示ではないものを止めると既存の原稿が通らなくなる．
    """
    slide = parse(_FM + ("## 見出し\n\n- a\n"
                         "<!-- handout-skip -->\n"
                         "<!-- ここはメモ -->\n"
                         "- b\n")).slides[0]
    assert [ln.text for ln in slide.blocks] == ["a", "b"]


def test_the_message_lists_step():
    """案内する一覧に ``@step`` が並ぶ．

    打ち間違いを一覧から探せることがこのエラーの役目で、**値を取らない指示が
    一覧から漏れていた**（値付きの一覧に手で 2 つ足していたため）．
    いちばん踏むのが ``@step`` の打ち間違いなので、漏れていては案内にならない．
    """
    with pytest.raises(ValueError, match=r"@step[,)]"):
        parse(_with("<!-- @stepp -->"))


def test_a_value_taking_directive_without_a_colon_asks_for_the_value():
    """名前が合っていてコロンが無い行は、綴りではなく書き方の誤りとして案内する．"""
    with pytest.raises(ValueError, match=r"@body-size needs a value"):
        parse(_with("<!-- @body-size 2 -->"))


def test_a_valueless_directive_with_junk_shows_the_right_form():
    """``@step`` に余りがある行は、正しい形を示して止まる．"""
    with pytest.raises(ValueError, match=r"invalid @step .*<!-- @step -->"):
        parse(_with("<!-- @step 2 -->"))


def test_a_renamed_directive_is_still_guided_without_a_value():
    """旧名称の案内は、コロンが無くても出る（値付きと同じ文言）．"""
    with pytest.raises(ValueError, match="renamed in v0.7; use @widths"):
        parse(_with("<!-- @ph-widths -->"))


def test_a_value_on_a_valueless_directive_shows_the_right_form():
    """``<!-- @step: 1 -->`` は未知のキー扱いにせず、正しい形を示す．"""
    with pytest.raises(ValueError, match=r"@step takes no value"):
        parse(_with("<!-- @step: 1 -->"))


def test_a_renamed_directive_with_a_value_is_guided_too():
    """旧名称に値を付けた従来の形も、同じ案内で止まる．

    名前の誤りを 1 箇所（``_directive_name_error``）に寄せたときに
    通る経路が変わったので、値付き側も固定しておく．
    """
    with pytest.raises(ValueError, match="renamed in v0.7; use @table-widths"):
        parse(_with("<!-- @col-widths: 30,70 -->"))
