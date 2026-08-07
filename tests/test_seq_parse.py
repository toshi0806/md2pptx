#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""``` ```seq ``` ブロック（ラダー図）を固定するテスト（Issue #110）．

ネットワークの講義でいちばん出てくる図が**プロトコルのラダー図**で、時間軸を持つ
ため flow の格子（#109）では表現できない。実測では TCP 3-way handshake・輻輳
ウィンドウ・Fast Retransmission（矢印25本以上）・HTTP・SMTP/POP・DNS 反復問い合わせ
と、ほぼ毎回出てくる。

**文法は3要素だけ**に固定する——``lifelines:`` と ``A -> B: ラベル`` と、
flow と共通の ``caption:`` / ``note(top):`` / ``note(bottom):``。
activation box も alt/loop も参加者の生成消滅も入れない。

外から見える ``parse_seq`` と ``plan_seq`` だけを叩く。``flow`` と同じ形にしてある
ので、内部のトークン表現には触れない。
"""
from __future__ import annotations

import pytest

from md2pptx.layout import emu
from md2pptx.seq import parse_seq, plan_seq


def _msgs(seq):
    return [(m.src, m.dst, m.label) for m in seq.messages]


# ---------------------------------------------------------------- パース

def test_it_reads_lifelines_and_messages():
    """``lifelines:`` で登場人物、``A -> B: ラベル`` でやりとりを書く．"""
    s = parse_seq("lifelines: 送信側, 受信側\n"
                  "送信側 -> 受信側: SYN\n"
                  "受信側 -> 送信側: SYN+ACK\n"
                  "送信側 -> 受信側: ACK\n")
    assert s.lifelines == ["送信側", "受信側"]
    assert _msgs(s) == [(0, 1, "SYN"), (1, 0, "SYN+ACK"), (0, 1, "ACK")]


def test_a_label_is_optional():
    """ラベルの無い矢印も書ける．"""
    s = parse_seq("lifelines: A, B\nA -> B\n")
    assert _msgs(s) == [(0, 1, None)]


def test_a_label_may_contain_a_colon():
    """ラベル中のコロンはラベルの一部（最初の ``:`` だけで割る）．

    ``ack 1: seq 100`` のような書き方は普通に出てくる。
    """
    s = parse_seq("lifelines: A, B\nA -> B: ack 1: seq 100\n")
    assert _msgs(s) == [(0, 1, "ack 1: seq 100")]


def test_lifelines_can_be_declared_implicitly():
    """``lifelines:`` を書かなければ、出てきた順に登場人物を作る．

    2 者のやりとりを書くだけなら宣言は要らない——書くことを減らす。
    """
    s = parse_seq("送信側 -> 受信側: SYN\n受信側 -> 送信側: SYN+ACK\n")
    assert s.lifelines == ["送信側", "受信側"]
    assert _msgs(s) == [(0, 1, "SYN"), (1, 0, "SYN+ACK")]


def test_a_self_message_is_rejected():
    """自分から自分への矢印は受けない（描き方を決めていない）．"""
    with pytest.raises(ValueError, match="cannot send to itself"):
        parse_seq("lifelines: A, B\nA -> A: x\n")


def test_an_unknown_lifeline_stops():
    """``lifelines:`` を書いたなら、そこに無い名前はタイポとみなして止める．"""
    with pytest.raises(ValueError, match="unknown lifeline"):
        parse_seq("lifelines: A, B\nA -> C: x\n")


def test_junk_stops():
    """ノート・設定・メッセージのどれでもない行は止める（flow と同じ方針）．"""
    with pytest.raises(ValueError, match="invalid seq syntax"):
        parse_seq("lifelines: A, B\nA <- B\n")


def test_settings_are_shared_with_flow():
    """``caption:`` / ``note(top):`` / ``note(bottom):`` は flow と同じ書き方．"""
    s = parse_seq("lifelines: A, B\n"
                  "note(top): 上の地の文\n"
                  "A -> B: x\n"
                  "caption: 図の説明\n"
                  "note(bottom): 下の地の文\n")
    assert (s.caption, s.note_top, s.note_bottom) == (
        "図の説明", "上の地の文", "下の地の文")


def test_a_note_line_annotates_the_diagram():
    """``note: …`` は図の中の注記（時間軸のその位置に置く）．"""
    s = parse_seq("lifelines: A, B\nA -> B: x\nnote: ここで待つ\nB -> A: y\n")
    assert [n.text for n in s.notes] == ["ここで待つ"]
    assert s.notes[0].after == 1        # 1 本目の矢印の後ろ


# ---------------------------------------------------------------- 配置

def test_lifelines_are_spread_across_the_width():
    """ライフラインは幅いっぱいに等間隔で並ぶ．"""
    plan = plan_seq(parse_seq("lifelines: A, B, C\nA -> B: x\n"),
                    0, 0, emu(9), emu(5))
    xs = [h.rect.center_x for h in plan.heads]
    assert len(xs) == 3
    assert xs[1] - xs[0] == xs[2] - xs[1]


def test_each_lifeline_gets_a_vertical_line():
    """各ライフラインに縦線が引かれる（矢尻は付かない）．"""
    plan = plan_seq(parse_seq("lifelines: A, B\nA -> B: x\n"),
                    0, 0, emu(8), emu(5))
    assert len(plan.lines) == 2
    for ln in plan.lines:
        assert ln.x1 == ln.x2            # 縦線
        assert ln.y1 < ln.y2


def test_messages_go_down_the_time_axis():
    """メッセージは上から下へ、書いた順に並ぶ（これが時間軸）．"""
    plan = plan_seq(parse_seq("lifelines: A, B\nA -> B: 1\nB -> A: 2\nA -> B: 3\n"),
                    0, 0, emu(8), emu(6))
    ys = [a.y1 for a in plan.arrows]
    assert ys == sorted(ys) and len(set(ys)) == 3


def test_a_message_starts_and_ends_on_its_lifelines():
    """矢印の両端はライフラインの上にある（横向きの線分）．"""
    seq = parse_seq("lifelines: A, B\nA -> B: SYN\n")
    plan = plan_seq(seq, 0, 0, emu(8), emu(5))
    xs = sorted(h.rect.center_x for h in plan.heads)
    a = plan.arrows[0]
    assert (a.x1, a.x2) == (xs[0], xs[1])
    assert a.y1 == a.y2


def test_a_reply_points_back():
    """返しの矢印は逆向き（始点と終点が入れ替わる）．"""
    plan = plan_seq(parse_seq("lifelines: A, B\nB -> A: ACK\n"),
                    0, 0, emu(8), emu(5))
    a = plan.arrows[0]
    assert a.x1 > a.x2


def test_labels_sit_above_their_arrows():
    """ラベルは矢印の上（線に重ならない位置）．"""
    plan = plan_seq(parse_seq("lifelines: A, B\nA -> B: SYN\n"),
                    0, 0, emu(8), emu(5))
    assert plan.labels[0].rect.bottom <= plan.arrows[0].y1


def test_many_messages_still_fit_in_the_band():
    """本数が増えても帯からはみ出さない（Fast Retransmission は矢印25本超）．"""
    src = "lifelines: A, B\n" + "".join(
        f"A -> B: m{i}\n" if i % 2 else f"B -> A: m{i}\n" for i in range(26))
    plan = plan_seq(parse_seq(src), 0, 0, emu(9), emu(5))
    assert len(plan.arrows) == 26
    assert max(a.y1 for a in plan.arrows) <= emu(5)
