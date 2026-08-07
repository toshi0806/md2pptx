#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""図の内部で ``@step`` を切れることを固定するテスト（Issue #125）．

``@step``（#104）はスライドのブロック列を**累積して切り出す**仕組みで、段の
区切りはブロック境界にしか置けなかった。そのため図の内部で段階を切れない。

**これは実用上いちばん多い形**でもある。前年度の実測では、アニメーション対象の
82% が「図の部品を1つずつ出す」ビルドアップだった（箇条書きの段階表示は少数派）。

図の中の ``@step`` は、フェンスの中に ``@step`` の行を書く。地の文は従来どおり
累積したまま、**図だけがその段階の姿に差し替わる**。
"""
from __future__ import annotations

from md2pptx.ir import Flow, Line, Seq
from md2pptx.parser import parse

_FM = "---\ntheme: t.pptx\n---\n\n"
_F = "```"


def _figs(slide, cls):
    return [b for b in slide.blocks if isinstance(b, cls)]


def _texts(slide):
    return [b.text for b in slide.blocks if isinstance(b, Line)]


# ---------------------------------------------------------------- seq

def test_a_seq_builds_up_one_message_at_a_time():
    """``@step`` の数だけスライドが増え、矢印が1本ずつ増える．"""
    src = _FM + (f"### TCP\n\n{_F}seq\nlifelines: A, B\n"
                 "A -> B: SYN\n@step\nB -> A: SYN+ACK\n@step\nA -> B: ACK\n"
                 f"{_F}\n")
    slides = parse(src).slides
    assert [len(_figs(s, Seq)[0].messages) for s in slides] == [1, 2, 3]


def test_the_lifelines_stay_the_same_in_every_step():
    """登場人物は最初から全員いる（途中で増えると図が横に動いて読みにくい）．"""
    src = _FM + (f"### x\n\n{_F}seq\nlifelines: A, B, C\n"
                 "A -> B: 1\n@step\nB -> C: 2\n"
                 f"{_F}\n")
    for s in parse(src).slides:
        assert _figs(s, Seq)[0].lifelines == ["A", "B", "C"]


def test_notes_appear_with_their_step():
    """図の中の注記は、その位置の矢印が出た段から現れる．"""
    src = _FM + (f"### x\n\n{_F}seq\nlifelines: A, B\n"
                 "A -> B: 1\n@step\nnote: ここで待つ\nB -> A: 2\n"
                 f"{_F}\n")
    first, last = parse(src).slides
    assert _figs(first, Seq)[0].notes == []
    assert [n.text for n in _figs(last, Seq)[0].notes] == ["ここで待つ"]


# ---------------------------------------------------------------- flow

def test_a_flow_builds_up_one_node_at_a_time():
    """flow はノード単位で増える．"""
    src = _FM + (f"### x\n\n{_F}flow\n[a] -> [b]\n@step\n-> [c]\n{_F}\n")
    slides = parse(src).slides
    assert [len(_figs(s, Flow)[0].nodes) for s in slides] == [2, 3]


def test_edges_come_with_the_node_they_reach():
    """矢印は、その先のノードが出た段で一緒に出る（宙ぶらりんの線を出さない）．"""
    src = _FM + (f"### x\n\n{_F}flow\n[a] -> [b]\n@step\n-> [c]\n{_F}\n")
    first, last = parse(src).slides
    assert len(_figs(first, Flow)[0].edges) == 1
    assert len(_figs(last, Flow)[0].edges) == 2


def test_rows_survive_the_truncation():
    """段（``--``）を持つ図でも、その段までの姿になる．"""
    src = _FM + (f"### x\n\n{_F}flow\n[a] [b]\n@step\n--\n[c]\n{_F}\n")
    first, last = parse(src).slides
    assert _figs(first, Flow)[0].rows == [[0, 1]]
    assert _figs(last, Flow)[0].rows == [[0, 1], [2]]


# ---------------------------------------------------------------- 併用

def test_the_prose_still_accumulates():
    """地の文は従来どおり累積する（図だけが差し替わる）．"""
    src = _FM + (f"### x\n\n- 導入\n\n{_F}seq\nlifelines: A, B\n"
                 "A -> B: 1\n@step\nB -> A: 2\n"
                 f"{_F}\n\n→ 結論\n")
    first, last = parse(src).slides
    assert _texts(first) == ["導入"]
    assert _texts(last) == ["導入", "→ 結論"]


def test_slide_level_steps_still_work():
    """スライドの ``@step`` は従来どおり（図を持たないスライドで挙動が変わらない）．"""
    src = _FM + "### x\n\n- a\n<!-- @step -->\n- b\n"
    first, last = parse(src).slides
    assert _texts(first) == ["a"] and _texts(last) == ["a", "b"]


def test_a_figure_without_steps_is_one_slide():
    """図の中に ``@step`` が無ければ、これまでどおり1枚のまま．"""
    src = _FM + (f"### x\n\n{_F}seq\nlifelines: A, B\nA -> B: 1\nB -> A: 2\n{_F}\n")
    slide, = parse(src).slides
    assert len(_figs(slide, Seq)[0].messages) == 2
