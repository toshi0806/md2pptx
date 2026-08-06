#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""``---`` によるスライド分割の非推奨を固定するテスト（Issue #92）．

非推奨は「そのうち消す」の予告で，いま壊すことではない．だから見るのは 2 つ——
**警告が出ること**と，**出力が 1 つも変わらないこと**．後者を落とすと，予告のはずが
移行を強制することになる．

段落直後の ``---`` を分割として扱う挙動もここで押さえる．CommonMark ではそれは
setext 見出しで，**この食い違いが非推奨にする理由のひとつ**——直すのは削除のときで，
それまでは変わらないことを保証する側にいる．
"""
from __future__ import annotations

from md2pptx.parser import parse

_FM = "---\ntheme: t.pptx\n---\n\n"


def _slides(src):
    return [(s.title, s.layout, [b.text for b in s.blocks])
            for s in parse(src).slides]


def test_the_rule_still_splits_slides(capsys):
    """``---`` はタイトルなしスライドを開始する（従来どおり）．"""
    assert _slides(_FM + "## A\n\n- x\n\n---\n\n- y\n") == [
        ("A", 1, ["x"]),
        (None, 1, ["y"]),
    ]
    capsys.readouterr()   # 警告は別のテストで見る


def test_the_rule_warns_with_a_line_number(capsys):
    """行番号と移行先を添えて警告する．"""
    parse(_FM + "## A\n\n- x\n\n---\n\n- y\n")
    err = capsys.readouterr().err
    assert "deprecated" in err
    assert "line 9" in err
    assert "## 見出し" in err
    assert "@layout: 6" in err


def test_each_rule_warns_once(capsys):
    """``---`` 1 件につき 1 回（どれを直すか分かるように）．"""
    parse(_FM + "## A\n\n- x\n\n---\n\n- y\n\n---\n\n- z\n")
    err = capsys.readouterr().err
    assert err.count("deprecated") == 2
    assert "line 9" in err and "line 13" in err


def test_front_matter_delimiters_do_not_warn(capsys):
    """front matter の ``---`` は対象外（同じ記号だが意味が違う）．"""
    parse(_FM + "## A\n\n- x\n")
    assert "deprecated" not in capsys.readouterr().err


def test_a_rule_after_a_paragraph_still_splits(capsys):
    """段落の直後でも分割のまま（CommonMark なら setext 見出しになる形）．"""
    assert _slides(_FM + "## A\n\n前の段落\n---\n\n- 次\n") == [
        ("A", 1, ["前の段落"]),
        (None, 1, ["次"]),
    ]
    assert "deprecated" in capsys.readouterr().err
