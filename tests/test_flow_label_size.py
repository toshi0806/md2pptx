#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""矢印ラベルの枠を**実際に描く文字サイズ**で見積もる（Issue #178）．

`_label_rect` の枠は `_LABEL_PT = 16pt` 前提の高さだったが、`render_flow` は
ラベルを本文標準サイズ（テーマによっては 30pt）で描いていた。枠は
``anchor=MIDDLE`` なので、大きい字は枠の**上下へ等しくはみ出す**——下側が
0.04in 下から始まる box に乗る（cn2026-12 p.42「VPN の姿」で「暗号化」が
`拠点 A` の上辺に食い込んでいた）。

`seq` は #168 で「描くほうを 16pt に縮める」向きで直したが、flow の矢印ラベルは
説明の一部（「暗号化」「委任」「NAMEPREP」）なので読ませたい。ここでは
**見積もりを実寸に合わせる**——render は自分が描くサイズを知っている。
"""
from __future__ import annotations

import pytest

from md2pptx.flow import _label_height, parse_flow, plan_flow
from md2pptx.layout import emu


def _plan(src: str, **kw):
    return plan_flow(parse_flow(src.strip()), 0, 0, emu(10.0), emu(4.0), **kw)


def test_the_label_box_grows_with_the_font():
    small = _plan("[A] -暗号化-> [B]", label_pt=16.0).labels[0].rect
    big = _plan("[A] -暗号化-> [B]", label_pt=30.0).labels[0].rect
    assert big.height > small.height
    assert big.width > small.width


def test_the_label_clears_the_boxes_it_sits_above():
    """30pt で描いても box に掛からない．

    枠は ``anchor=MIDDLE`` なので、字は枠の**中心**から上下へ広がる。
    枠の高さが実寸ぶんあれば、字は枠に収まり box には届かない。
    """
    plan = _plan("[A] -暗号化-> [B]", label_pt=30.0)
    lab = plan.labels[0]
    top_of_boxes = min(b.rect.top for b in plan.boxes)
    assert lab.rect.top + lab.rect.height <= top_of_boxes, "ラベルの枠が box に掛かっている"
    # 枠が 30pt ぶんの高さで作られているか——手で計算し直すと
    # ``_label_height`` の式と二重管理になるので、実装をそのまま突き合わせる．
    assert lab.rect.height == _label_height(30.0)


def test_the_default_is_unchanged():
    """既定は従来どおり 16pt 相当——呼び出し側が渡さなければ何も変わらない．"""
    assert _plan("[A] -x-> [B]").labels[0].rect.height == \
           _plan("[A] -x-> [B]", label_pt=16.0).labels[0].rect.height


@pytest.mark.parametrize("src", [
    "[A] -暗号化-> [B]",
    "direction: tb\n[A]\n-暗号化-> [B]",
    "[#a A] -暗号化-> [#b B]\n--\n[#c C]\nc -> b",
])
def test_every_layout_sizes_its_labels(src):
    """横並び・縦並び・格子のどれでもラベルは実寸で見積もられる．"""
    small = _plan(src, label_pt=16.0).labels[0].rect
    big = _plan(src, label_pt=30.0).labels[0].rect
    assert big.height > small.height
