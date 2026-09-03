#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""縦並び flow の box 幅を中身から決める（Issue #184）．

``bw = min(_emu(3.2), width * 0.5)`` は中身を見ていなかったので、長いラベルが
折り返して box が縦に伸び、図が窮屈になっていた（cn2026-07 p.34 の
「日本語ﾄﾞﾒｲﾝ名ＥＸＡＭＰＬＥ。ｊｐ」）。

いちばん長いラベルが 1 行に収まる幅にする。上限は帯、下限は従来の 3.2in
——短い名前ばかりのときに box が痩せて見えないように。
"""
from __future__ import annotations

from md2pptx.flow import _label_width, parse_flow, plan_flow
from md2pptx.layout import emu

BAND_W, BAND_H = emu(12.0), emu(4.5)


def _boxes(src: str, label_pt: float = 30.0):
    plan = plan_flow(parse_flow(src.strip()), 0, 0, BAND_W, BAND_H,
                     label_pt=label_pt)
    return plan.boxes


def test_a_long_label_gets_a_wider_box():
    short = _boxes("direction: tb\n[A]\n-> [B]")[0].rect.width
    long_ = _boxes("direction: tb\n[日本語ドメイン名EXAMPLE。jp]\n-> [B]")[0].rect.width
    assert long_ > short


def test_the_box_fits_the_longest_label_on_one_line():
    label = "日本語ドメイン名EXAMPLE。jp"
    box = _boxes(f"direction: tb\n[{label}]\n-> [B]")[0].rect
    assert box.width >= _label_width(label, 30.0)


def test_short_labels_keep_the_old_width():
    """短い名前ばかりなら従来どおり——box が痩せて見えないための下限．"""
    assert _boxes("direction: tb\n[A]\n-> [B]")[0].rect.width == \
        min(emu(3.2), int(BAND_W * 0.5))


def test_the_box_never_exceeds_the_band():
    box = _boxes("direction: tb\n[" + "あ" * 200 + "]\n-> [B]")[0].rect
    assert box.width <= BAND_W


def test_all_boxes_share_one_width():
    """幅は全 box 共通——1 つだけ広いと列に見えない．"""
    boxes = _boxes("direction: tb\n[短]\n-> [とても長い名前のノードEXAMPLE。jp]")
    assert len({b.rect.width for b in boxes}) == 1


def test_the_column_stays_centred():
    boxes = _boxes("direction: tb\n[日本語ドメイン名EXAMPLE。jp]\n-> [B]")
    r = boxes[0].rect
    assert abs((r.left - 0) - (BAND_W - (r.left + r.width))) <= 1


def test_the_gap_makes_room_for_an_arrow_label():
    """ラベルのある縦並びは、box のすき間がラベルの高さより広い．

    box を中身に合わせて広げたら折り返しが消え、そのぶん box が縦に伸びて
    すき間が詰まった。すき間はラベルの高さを見て決める。
    """
    from md2pptx.flow import _label_height
    plan = plan_flow(parse_flow("direction: tb\n[A]\n-NAMEPREP-> [B]"),
                     0, 0, BAND_W, BAND_H, label_pt=30.0)
    a, b = (x.rect for x in plan.boxes)
    gap = b.top - (a.top + a.height)
    assert gap >= _label_height(30.0), f"すき間 {gap} がラベルより狭い"


def test_a_label_free_column_keeps_the_tight_gap():
    """ラベルが無ければ従来どおり詰めて置く（間延びさせない）．"""
    plan = plan_flow(parse_flow("direction: tb\n[A]\n-> [B]"),
                     0, 0, BAND_W, BAND_H, label_pt=30.0)
    a, b = (x.rect for x in plan.boxes)
    assert b.top - (a.top + a.height) == emu(0.35)
