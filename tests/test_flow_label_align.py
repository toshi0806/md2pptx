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

from md2pptx.flow import parse_flow, plan_flow
from md2pptx.layout import emu


def _plan(src: str):
    return plan_flow(parse_flow(src.strip()), 0, 0, emu(10.0), emu(4.0))


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
    lab = plan.labels[0]
    assert lab.rect.left > plan.arrows[0].x1, "枠の左端が矢印より左にある"
    assert lab.align == "left", "中央揃えだと見積もり超過ぶんが矢印側へ出る"


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
    assert plan.labels
    assert all(lab.align == "center" for lab in plan.labels)
