#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""``<!-- @step -->`` による段階表示を固定するテスト（Issue #104）．

md2pptx は PowerPoint のアニメーションを書かない。代わりに **1 枚分の原稿から
複数枚を生成し、内容を積み上げて見せる**。講義スライドの実測では、既にアニメー
ションではなくスライド複製で段階表示している箇所が多数あり、この形は元原稿の
実態とも一致する。

段階の切り出しは**書いた場所ではなく、スライドを閉じるときにまとめて**行う。
``@step`` の位置で即座にスナップショットを取ると、その後に書いた ``@layout`` や
``` ```note ``` が前の段に入らず、「どこに書いたか」で結果が変わってしまう。
段階はスライドの一部であって別のスライドではないので、**タイトルもディレクティブも
全段で同じ**になるほうが原稿の読み方と合う。

外から見える ``parse()`` だけを叩く。内部の切り出し関数やマークの持ち方には
触れない——そこは実装の都合で、固定したいのは「原稿がどう読まれるか」だけ。
"""
from __future__ import annotations

from md2pptx.ir import CONTENT_LAYOUT, Line, Table
from md2pptx.parser import parse

_FM = "---\ntheme: t.pptx\n---\n\n"


def _texts(slide):
    """スライド本文の Line を文字列で返す（表・図は除く）．"""
    return [b.text for b in slide.blocks if isinstance(b, Line)]


def _col_texts(slide):
    """多カラムスライドの各カラムの Line を文字列で返す．"""
    return [[b.text for b in col if isinstance(b, Line)] for col in slide.columns]


def test_one_source_slide_becomes_several():
    """``@step`` を 2 つ書けば 3 枚になる（区切りの数 + 1）．"""
    src = _FM + ("### TCP 3-way handshake\n\n"
                 "- SYN を送る\n"
                 "<!-- @step -->\n"
                 "- SYN+ACK が返る\n"
                 "<!-- @step -->\n"
                 "- ACK を返して確立\n")
    assert len(parse(src).slides) == 3


def test_each_step_accumulates():
    """段は積み上がる——1 枚目は 1 行、2 枚目は 2 行、3 枚目は 3 行．

    「消えて次が出る」ではなく「足されていく」。アニメーションの開始効果
    （アピール・ワイプ・フェード）を置き換えるのが目的なので、既に見せたものは
    そのまま残る。
    """
    src = _FM + ("### TCP 3-way handshake\n\n"
                 "- SYN を送る\n"
                 "<!-- @step -->\n"
                 "- SYN+ACK が返る\n"
                 "<!-- @step -->\n"
                 "- ACK を返して確立\n")
    a, b, c = parse(src).slides
    assert _texts(a) == ["SYN を送る"]
    assert _texts(b) == ["SYN を送る", "SYN+ACK が返る"]
    assert _texts(c) == ["SYN を送る", "SYN+ACK が返る", "ACK を返して確立"]


def test_the_title_is_shared_by_every_step():
    """タイトルは全段で同じ（段はスライドの一部であって別のスライドではない）．"""
    src = _FM + "### 輻輳ウィンドウ\n\n- a\n<!-- @step -->\n- b\n"
    assert [s.title for s in parse(src).slides] == ["輻輳ウィンドウ"] * 2


def test_directives_are_shared_even_when_written_after_a_step():
    """ディレクティブは全段に効く——``@step`` より後に書いても．

    切り出しをスライドの終わりまで遅らせている理由がここ。書いた場所で結果が
    変わると、「ディレクティブはスライド先頭に書く」という書き方の約束と
    食い違う。
    """
    src = _FM + ("### 図\n\n"
                 "- a\n"
                 "<!-- @step -->\n"
                 "- b\n"
                 "<!-- @layout: 5 -->\n")
    assert [s.layout for s in parse(src).slides] == [5, 5]


def test_notes_go_to_the_last_step_only():
    """発表者ノートは最終段だけに付く．

    全段に複製すると発表者ビューで同じ原稿が何度も出て読みにくい。
    段の集まりで 1 つの話なので、ノートも 1 つでよい。
    """
    src = _FM + ("### DNS 反復問い合わせ\n\n"
                 "- root に聞く\n"
                 "<!-- @step -->\n"
                 "- jp に聞く\n"
                 "\n```note\nここで反復問い合わせの説明をする。\n```\n")
    first, last = parse(src).slides
    assert first.notes is None
    assert last.notes == "ここで反復問い合わせの説明をする。"


def test_tables_and_other_blocks_accumulate_too():
    """積み上がるのは箇条書きだけではない（表・図も同じ）．"""
    src = _FM + ("### NAT 変換表\n\n"
                 "変換前\n"
                 "<!-- @step -->\n"
                 "| 方向 | 変換後 |\n"
                 "|:--|:--|\n"
                 "| 送信 | 203.0.113.1 |\n")
    first, last = parse(src).slides
    assert not [b for b in first.blocks if isinstance(b, Table)]
    assert len([b for b in last.blocks if isinstance(b, Table)]) == 1


def test_a_step_before_any_content_gives_a_title_only_slide():
    """本文より先に ``@step`` を置けば、タイトルだけの段から始まる．"""
    src = _FM + "### まとめ\n\n<!-- @step -->\n- 要点\n"
    first, last = parse(src).slides
    assert first.title == "まとめ" and _texts(first) == []
    assert _texts(last) == ["要点"]


def test_a_step_inside_a_column_keeps_the_other_column():
    """多カラムでも積み上がる——右カラムの段で左カラムは消えない．"""
    src = _FM + ("### 比較\n\n"
                 "- 左1\n"
                 "<!-- @col -->\n"
                 "- 右1\n"
                 "<!-- @step -->\n"
                 "- 右2\n")
    first, last = parse(src).slides
    assert _col_texts(first) == [["左1"], ["右1"]]
    assert _col_texts(last) == [["左1"], ["右1", "右2"]]


def test_a_step_before_the_column_break_leaves_the_right_column_empty():
    """カラム区切りより前の段は、右カラムが空の状態になる．

    段はどれも**最終的なカラム構成**で描く。途中の段だけ単一カラムにすると、
    レイアウトが段ごとに変わって行頭の位置が動いてしまう。
    """
    src = _FM + ("### 比較\n\n"
                 "- 左1\n"
                 "<!-- @step -->\n"
                 "- 左2\n"
                 "<!-- @col -->\n"
                 "- 右1\n")
    first, last = parse(src).slides
    assert _col_texts(first) == [["左1"], []]
    assert _col_texts(last) == [["左1", "左2"], ["右1"]]
    assert [s.layout for s in parse(src).slides] == [3, 3]


def test_steps_do_not_leak_into_the_next_slide():
    """段の区切りは次の見出しで終わる．"""
    src = _FM + ("### A\n\n- a1\n<!-- @step -->\n- a2\n\n"
                 "### B\n\n- b1\n")
    slides = parse(src).slides
    assert [s.title for s in slides] == ["A", "A", "B"]
    assert _texts(slides[2]) == ["b1"]


def test_a_slide_without_steps_is_unchanged():
    """``@step`` を書かなければ、これまでどおり 1 枚のまま．"""
    src = _FM + "### ふつう\n\n- a\n- b\n"
    slide, = parse(src).slides
    assert slide.layout == CONTENT_LAYOUT
    assert _texts(slide) == ["a", "b"]


def test_the_steps_share_no_block_list_with_each_other():
    """段どうしがブロック列を共有しない（後から足しても前の段は変わらない）．

    切り出しを遅らせている以上、スライスの取り方を間違えると全段が同じものを
    指してしまう。それが起きていないことをここで見ておく。
    """
    src = _FM + "### x\n\n- a\n<!-- @step -->\n- b\n"
    first, last = parse(src).slides
    assert first.blocks is not last.blocks
    assert len(first.blocks) == 1 and len(last.blocks) == 2
