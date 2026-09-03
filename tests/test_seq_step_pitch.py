#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""``` ```seq ``` の ``@step`` で図が上下に動かない（Issue #182）．

`upto(n)` はメッセージを切り詰めるので、`plan_seq` が**その本数で送り幅を
決めて**いた。矢印が少ない段ほど間隔が広がり、段が進むたびに図全体が動く
（cn2026-08 p.25〜27 の SMTP ラダー）。

`Seq.upto` の docstring は同じ理由で**登場人物は減らさない**と書いている——
横は直っていて縦が残っていた。最終段の本数で送り幅を決め、描くのは
その段のぶんだけにする。
"""
from __future__ import annotations

import pytest

from md2pptx.layout import emu
from md2pptx.seq import parse_seq, plan_seq

SRC = """
lifelines: A, B
A -> B: 1
B -> A: 2
@step
A -> B: 3
B -> A: 4
@step
A -> B: 5
B -> A: 6
"""


def _plan(seq):
    return plan_seq(seq, 0, 0, emu(10.0), emu(4.0))


def _pitch(plan):
    """隣り合う矢印の間隔．1 本しか無ければ None．"""
    ys = sorted({a.y1 for a in plan.arrows})
    return ys[1] - ys[0] if len(ys) > 1 else None


def test_every_step_uses_the_same_pitch():
    full = parse_seq(SRC.strip())
    assert full.steps, "@step が読めていない"
    pitches = []
    for n in full.steps + [len(full.messages)]:
        p = _pitch(_plan(full.upto(n)))
        if p is not None:
            pitches.append(p)
    assert len(set(pitches)) == 1, f"段ごとに間隔が違う: {pitches}"


def test_every_step_starts_at_the_same_y():
    """1 本目の矢印は、どの段でも同じ高さに出る．"""
    full = parse_seq(SRC.strip())
    firsts = {_plan(full.upto(n)).arrows[0].y1
              for n in full.steps + [len(full.messages)]}
    assert len(firsts) == 1, f"1 本目の位置が段で違う: {sorted(firsts)}"


def test_each_step_draws_only_its_own_arrows():
    """場所は最終段ぶん取るが、**描くのはその段まで**．"""
    full = parse_seq(SRC.strip())
    for n in full.steps + [len(full.messages)]:
        assert len(_plan(full.upto(n)).arrows) == n


def test_a_seq_without_steps_is_unchanged():
    """段を持たない図の見た目は変えない．"""
    one = parse_seq("lifelines: A, B\nA -> B: 1\nB -> A: 2")
    assert _pitch(_plan(one)) == _pitch(_plan(parse_seq(
        "lifelines: A, B\nA -> B: 1\nB -> A: 2")))
