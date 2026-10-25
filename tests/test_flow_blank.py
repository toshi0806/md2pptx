#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""フロー図の空セル ``[]`` を固定するテスト（Issue #202）。

**段ごとに要素数が違うとき、少ない段の位置を決められなかった。** 段は左詰めなので、
3対1 の図では下段が列0 に落ちる。置きたいのは真ん中の列——3本の矢印が左右対称に
集まる位置——だが、そこへ寄せる手段が無かった。

逃げ道はどれも噛み合わなかった。``[…]`` を前に置くと**省略していない段に省略記号が
出る**。全角空白 ``[　]`` は ``_make_node`` の ``strip()`` がラベルを落とすので
**枠の見える空の box** になる。しかも box が1つ増えるぶん ``_box_palette`` の
割り当てがずれて、**関係の無いノードの色が変わる**（IX が金色ではなく緑になった）。

``[]`` は1列ぶんの場所を取り、何も描かない。外から見える ``parse_flow`` /
``plan_flow`` だけを叩く（内部のトークン表現には触れない）。
"""
from __future__ import annotations

import pytest

from md2pptx.flow import parse_flow, plan_flow
from md2pptx.layout import emu


# Issue #202 の実例（cn2026-05 の IX の図）。下段の書き方だけ差し替えて使う。
_3_TO_1 = ("[#a AS A] [#b AS B] [#c AS C]\n"
           "--\n"
           "{row}\n"
           "a -> x\n"
           "b -> x\n"
           "c -> x")


def _rects(plan):
    """ラベル → 矩形。段をまたいでも引けるように dict で返す。"""
    return {b.node.label: b.rect for b in plan.boxes}


# ------------------------------------------------------------ 空セルの書き方

@pytest.mark.parametrize("cell", ["[]", "[ ]", "[　]"])
def test_an_empty_label_is_a_blank_cell(cell):
    """``[]`` / ``[ ]`` / ``[　]`` はすべて空セル（ラベルの無い box ではない）。

    全角空白が書けてしまったのは ``strip()`` が落とすからで、結果は
    「枠の見える空の box」だった。同じ書き方を空セルの意味へ寄せる。
    """
    f = parse_flow(f"[A] {cell} [B]")
    assert [n.kind for n in f.nodes] == ["box", "blank", "box"]


def test_a_blank_cell_draws_nothing():
    """空セルは塗り・輪郭・影を描かない（``boxes`` にも ``ellipses`` にも出ない）。"""
    plan = plan_flow(parse_flow("[A] [] [B]"), 0, 0, emu(10), emu(5))

    assert [b.node.label for b in plan.boxes] == ["A", "B"]
    assert plan.ellipses == []


def test_a_blank_cell_does_not_consume_a_box_colour():
    """空セルは ``_box_palette`` を消費しない。

    色は render が ``plan.boxes`` を数えて割り当てるので、ここに空セルが混じると
    **以降のノードの色がずれる**。並び（＝割り当ての順番）が空セルの有無で
    変わらないことを見る。``[…]`` が box ではないので色を消費しないのと同じ理屈。

    **これは間接的な検証**——``flow.py`` は python-pptx 非依存で色を持たないので、
    ここで見られるのは「割り当ての順番」まで。``render`` が
    ``enumerate(plan.boxes)`` で ``_box_palette`` を引くという前提が成り立つ間だけ
    色の保証になる。引き方を変えるなら、こちらも併せて見直すこと。
    """
    with_blank = plan_flow(parse_flow("[A] [] [B] [] [C]"),
                           0, 0, emu(12), emu(5))
    without = plan_flow(parse_flow("[A] [B] [C]"), 0, 0, emu(12), emu(5))

    assert [b.node.label for b in with_blank.boxes] == \
        [b.node.label for b in without.boxes]


# ------------------------------------------------------------ 格子（段がある図）

def test_a_blank_cell_shifts_the_shorter_row_by_one_column():
    """3対1 の図で ``[] [#x IX] []`` と書くと IX が列1 に来る（Issue #202 の実例）。"""
    plan = plan_flow(parse_flow(_3_TO_1.format(row="[] [#x IX] []")),
                     0, 0, emu(12), emu(6))
    r = _rects(plan)

    assert r["IX"].left == r["AS B"].left


def test_a_blank_cell_in_a_grid_is_one_column_wide():
    """格子の空セルは box と同じ1列ぶん（省略記号の固定幅ではない）。

    ``[…]`` は「…」1文字の注記なので幅が狭い。空セルは「box があるはずの場所」の
    代わりなので、同じ幅でなければ列が揃わない。
    """
    blank = plan_flow(parse_flow(_3_TO_1.format(row="[] [#x IX]")),
                      0, 0, emu(12), emu(6))
    box = plan_flow(parse_flow(_3_TO_1.format(row="[D] [#x IX]")),
                    0, 0, emu(12), emu(6))

    assert _rects(blank)["IX"].left == _rects(box)["IX"].left


def test_a_blank_cell_does_not_widen_the_grid():
    """空セルは列数にだけ効く（列の幅は段の要素数で決まるので変わらない）。"""
    plan = plan_flow(parse_flow(_3_TO_1.format(row="[] [#x IX] []")),
                     0, 0, emu(12), emu(6))
    r = _rects(plan)

    assert r["IX"].width == r["AS B"].width


# ------------------------------------------------------------ 一列（lr / tb）

def test_a_blank_cell_takes_one_column_in_lr():
    """横一列でも box と同じ幅を取る（後ろのノードが box 1つぶん押される）。"""
    blank = plan_flow(parse_flow("[A] [] [B]"), 0, 0, emu(10), emu(5))
    box = plan_flow(parse_flow("[A] [X] [B]"), 0, 0, emu(10), emu(5))

    assert _rects(blank)["B"].left == _rects(box)["B"].left
    assert _rects(blank)["A"].width == _rects(box)["A"].width


def test_a_blank_cell_is_wider_than_an_ellipsis_in_lr():
    """空セルと ``[…]`` は別物——省略記号は固定幅なので後ろが詰まる。"""
    blank = plan_flow(parse_flow("[A] [] [B]"), 0, 0, emu(10), emu(5))
    ell = plan_flow(parse_flow("[A] […] [B]"), 0, 0, emu(10), emu(5))

    assert _rects(ell)["B"].left < _rects(blank)["B"].left


def test_a_blank_cell_takes_one_row_in_tb():
    """縦一列では1行ぶんの高さを取る。

    受け入れ条件には挙がっていないが、``kind`` を増やすと分岐が配置の3経路
    （格子・横一列・縦一列）に出る。**片方だけ通すと、向きを変えた原稿で
    黙って崩れる**ので、tb も box と同じ扱いに揃えておく。
    """
    blank = plan_flow(parse_flow("direction: tb\n[A] [] [B]"),
                      0, 0, emu(6), emu(8))
    box = plan_flow(parse_flow("direction: tb\n[A] [X] [B]"),
                    0, 0, emu(6), emu(8))

    assert _rects(blank)["B"].top == _rects(box)["B"].top
    assert _rects(blank)["A"].height == _rects(box)["A"].height


# ------------------------------------------------------------ 書けないもの

# 何も描かないものに意味を与える書き方は、黙って無視すると**書いた人の意図と
# 出力が食い違う**。どれも止める。
@pytest.mark.parametrize("src, msg", [
    ("[A] [#x] [B]", "blank flow cell takes no name"),
    ("[A] []{accent2} [B]", "blank flow cell takes no color"),
    ("[A] [ | x] [B]", "blank flow cell takes no sublabel"),
])
def test_a_blank_cell_takes_no_attributes(src, msg):
    with pytest.raises(ValueError, match=msg):
        parse_flow(src)


@pytest.mark.parametrize("src", ["[A] -> []", "[] -> [A]", "[A] -付き-> []"])
def test_a_blank_cell_cannot_be_an_edge_endpoint(src):
    """矢印の端点にはならない（刺す先が無い）。"""
    with pytest.raises(ValueError,
                       match="blank flow cell cannot be an edge endpoint"):
        parse_flow(src)


def test_a_named_node_still_needs_a_label():
    """``[#x]`` は名前付きの空セル——名前を付けたのにラベルが無いので止める。

    これまでは ``#x`` が**ラベルとして描かれていた**（名前の正規表現が
    ``#名前`` のあとに空白を要求するため、名前として読まれない）。
    名前を付けたつもりの原稿が黙って ``#x`` という箱になるほうが分かりにくい。
    """
    with pytest.raises(ValueError, match="blank flow cell takes no name"):
        parse_flow("[#x]")


def test_a_named_node_with_a_label_is_unchanged():
    """``[#pc PC]`` はこれまでどおり（名前の読み方を変えても既存の原稿は動く）。"""
    f = parse_flow("[#pc PC | 192.168.0.2] -> [#srv Server]")

    assert [n.node_id for n in f.nodes] == ["pc", "srv"]
    assert [n.label for n in f.nodes] == ["PC", "Server"]
    assert [n.sublabel for n in f.nodes] == ["192.168.0.2", None]
    assert [n.kind for n in f.nodes] == ["box", "box"]
