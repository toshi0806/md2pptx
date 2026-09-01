#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""縦並びの flow で矢印ラベルを左寄せにする（Issue #176）．

`direction: tb` のラベルは矢印の**右横**に置く。ところが render は中央揃えで
折り返さずに描くので、``_label_width`` の見積もりより実際の字が広いと
**左右に等しくはみ出し、左側が矢印に掛かる**。`NAMEPREP` の `N` が矢尻の下に
隠れて `AMEPREP` に見えていた（cn2026-07 p.34）。

係数を上げても別の語で外れるので、**はみ出す向きを決める**。縦並びのラベルは
左寄せにして、余りは矢印から離れる側（右）へ出す。

横並び・格子のラベルは矢印の**真上**に置いており、左右どちらへ出ても矢印には
掛からない。中央揃えのままにする——見た目が変わらないほうがよい。
"""
from __future__ import annotations

import pytest

from md2pptx.flow import parse_flow, plan_flow
from md2pptx.layout import emu


def _plan(src: str, width: float = 10.0, height: float = 4.0):
    """既定の帯は 10×4in——講義スライドの本文枠（12×4.95in）より一回り小さい．

    ここで見ているのはラベルと矢印の**前後関係**であって隙間の大きさではない。
    `_label_rect` は矢印の x から `_emu(0.18)` 右に枠を置くので、帯の寸法に
    依存しないはず——**そう書くだけでは確かめたことにならない**ので、
    `test_the_verdict_does_not_depend_on_the_band_size` で実際に振ってある。
    """
    return plan_flow(parse_flow(src.strip()), 0, 0, emu(width), emu(height))


def test_a_column_label_stays_clear_of_the_arrow():
    """左寄せ＋枠の左端が矢印より右——この 2 つが揃って初めて掛からない．

    枠が矢印より右にあるだけでは足りない（もともとそうなっている）。
    中央揃えのままだと見積もりを超えたぶんが**左へも**出て矢印に掛かる。
    """
    plan = _plan("""
direction: tb
[A]
-NAMEPREP-> [B]
""")
    assert len(plan.labels) == 1
    assert len(plan.arrows) == 1
    arrow = plan.arrows[0]
    # 縦並びの矢印は垂直線なので、両端の x が矢印の横位置そのもの．
    # ここを固定しておくと、``PlacedArrow`` の意味が変わったときに
    # **下の比較が黙って別のものを測り始める**のを防げる．
    assert arrow.x1 == arrow.x2, "縦並びの矢印が垂直でない"
    arrow_x = arrow.x1
    lab = plan.labels[0]
    assert lab.rect.left > arrow_x, "枠の左端が矢印より左に来てしまった"
    assert lab.align == "left", "中央揃えのままでは見積もり超過ぶんが矢印側へ出る"


def test_a_row_label_keeps_its_centre_alignment():
    plan = _plan("[A] -TCP-> [B]")
    assert len(plan.labels) == 1
    assert plan.labels[0].align == "center"


def test_a_grid_label_keeps_its_centre_alignment():
    plan = _plan("""
[#a A] -x-> [#b B]
--
[#c C]
c -> b
""")
    # 本数まで固定する——``assert plan.labels`` だけだと、ラベルが
    # 生成されなくなったときに ``all(...)`` が空で真になって**素通りする**．
    assert len(plan.labels) == 1
    assert plan.labels[0].align == "center"


@pytest.mark.parametrize("width,height", [
    (3.0, 1.5),     # 帯が最小（box が下限に張り付く）
    (10.0, 4.0),    # 既定
    (12.0, 4.95),   # 講義スライドの本文枠の実寸
    (24.0, 9.0),    # 極端に広い
])
def test_the_verdict_does_not_depend_on_the_band_size(width, height):
    """帯を振ってもラベルは矢印の右に残る．

    ヘルパが渡す 10×4in がたまたま都合の良い寸法だっただけ、という
    取り違えを防ぐ（レビュー指摘）。
    """
    plan = _plan("""
direction: tb
[A]
-NAMEPREP-> [B]
""", width, height)
    arrow = plan.arrows[0]
    assert arrow.x1 == arrow.x2
    assert plan.labels[0].rect.left > arrow.x1
    assert plan.labels[0].align == "left"
