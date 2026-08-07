#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""flow のノードID と格子配置を固定するテスト（Issue #109）．

flow DSL は横一列／縦一列しか書けなかった。``_build`` が直前のノードとの間にしか
エッジを張らないうえ、ノードを名前で指す手段が無いためで、分岐も格子も原理的に
書けない。講義スライドの実測では**2次元の自由作図が95枚（全544枚の17%）**あり、
ここが最大のボトルネックだった。

**新しいブロック型は増やさない。** 文法追加は2つだけ:

- ``[#pc PC | 192.168.0.2]`` — ``#pc`` でノードに名前を付ける
- ``--`` だけの行 — そこから次の段

名前を ``#`` で書くのは、``[id: ラベル]`` だと
``[HTTP: ハイパーテキスト転送プロトコル]`` のような**ラベル中のコロンと区別できない**
ため。``#`` はアンカーの慣習と合い、ラベルの先頭に来ることはまず無い。

外から見える ``parse_flow`` と ``plan_flow`` だけを叩く。トークンの内部表現には触れない。
"""
from __future__ import annotations

import pytest

from md2pptx.flow import parse_flow, plan_flow
from md2pptx.layout import emu


def _ids(flow):
    return [n.node_id for n in flow.nodes]


def _labels(flow):
    return [n.label for n in flow.nodes]


def _edges(flow):
    return [(e.src, e.dst, e.label) for e in flow.edges]


# ---------------------------------------------------------------- 既存の原稿

def test_a_plain_chain_is_unchanged():
    """名前も段区切りも書かなければ、これまでと同じ一列のフロー．"""
    f = parse_flow("[a] -変換-> [b] -> [c]")
    assert _labels(f) == ["a", "b", "c"]
    assert _edges(f) == [(0, 1, "変換"), (1, 2, None)]
    assert _ids(f) == [None, None, None]
    assert f.rows == []          # 段を書かなければ段は無い（＝一列）


def test_the_existing_layout_is_unchanged():
    """一列のときの座標はこれまでどおり（段の機能が既存の見た目を変えない）．"""
    plan = plan_flow(parse_flow("[a] -> [b]"), 0, 0, emu(10), emu(5))
    first, second = (b.rect for b in plan.boxes)
    assert first.left < second.left
    assert first.center_y == second.center_y


# ---------------------------------------------------------------- ノードID

def test_a_node_can_be_named():
    """``#pc`` で名前が付き、ラベルからは取り除かれる．"""
    f = parse_flow("[#pc PC | 192.168.0.2] -> [#srv Server]")
    assert _ids(f) == ["pc", "srv"]
    assert _labels(f) == ["PC", "Server"]
    assert f.nodes[0].sublabel == "192.168.0.2"


def test_a_colon_in_a_label_is_not_a_name():
    """ラベル中のコロンは名前と解釈しない（``#`` を使う理由）．"""
    f = parse_flow("[HTTP: ハイパーテキスト転送プロトコル]")
    assert _ids(f) == [None]
    assert _labels(f) == ["HTTP: ハイパーテキスト転送プロトコル"]


def test_edges_can_reference_names():
    """名前で任意のノード間を結べる（**隣り合っていなくてよい**）．

    これが無いと分岐も合流も書けない——``_build`` が直前のノードとの間にしか
    エッジを張れなかったのが、一列しか書けなかった理由。
    """
    f = parse_flow("[#a A] -> [#b B] -> [#c C]\na -飛び越し-> c")
    assert _edges(f) == [(0, 1, None), (1, 2, None), (0, 2, "飛び越し")]


def test_a_name_can_be_reused_in_several_edges():
    """1つのノードから何本でも出せる（分岐）．"""
    f = parse_flow("[#hub Hub]\n--\n[#a A] [#b B]\nhub -> a\nhub -> b")
    assert _edges(f) == [(0, 1, None), (0, 2, None)]


def test_an_unknown_name_stops():
    """知らない名前はタイポとみなして止める（黙って線を落とさない）．"""
    with pytest.raises(ValueError, match="unknown flow node name"):
        parse_flow("[#a A]\na -> nosuch")


def test_a_duplicate_name_stops():
    """名前の重複は止める（どちらを指すか決められない）．"""
    with pytest.raises(ValueError, match="duplicate flow node name"):
        parse_flow("[#a A] -> [#a B]")


# ---------------------------------------------------------------- 段

def test_a_row_separator_starts_a_new_row():
    """``--`` だけの行で段が変わる．"""
    f = parse_flow("[a] [b]\n--\n[c]")
    assert f.rows == [[0, 1], [2]]
    assert _labels(f) == ["a", "b", "c"]


def test_rows_stack_downwards():
    """段は上から下へ積まれ、段の中は左から右（``direction: lr``）．"""
    f = parse_flow("[a] [b]\n--\n[c] [d]")
    plan = plan_flow(f, 0, 0, emu(10), emu(5))
    a, b, c, d = (x.rect for x in plan.boxes)
    assert a.center_y == b.center_y          # 同じ段は同じ高さ
    assert c.center_y == d.center_y
    assert a.center_y < c.center_y           # 次の段は下
    assert a.left < b.left                   # 段の中は左から右


def test_rows_of_different_length_line_up_from_the_left():
    """段ごとに要素数が違ってもよい（欠けたところは空く）．"""
    f = parse_flow("[a] [b] [c]\n--\n[d]")
    plan = plan_flow(f, 0, 0, emu(12), emu(6))
    a, b, c, d = (x.rect for x in plan.boxes)
    assert a.left == d.left                  # 各段の先頭は同じ列
    assert a.width == d.width


def test_an_edge_between_rows_is_drawn():
    """段をまたぐ矢印も引ける（トポロジ図の縦のつながり）．"""
    f = parse_flow("[#a A]\n--\n[#b B]\na -> b")
    plan = plan_flow(f, 0, 0, emu(8), emu(6))
    assert len(plan.arrows) + len(plan.lines) == 1


def test_a_trailing_separator_does_not_make_an_empty_row():
    """末尾の ``--`` で空の段を作らない（原稿の書き癖を素通しする）．"""
    f = parse_flow("[a]\n--\n[b]\n--\n")
    assert f.rows == [[0], [1]]
