#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""表・図の帯が結論文に食い込まないことを固定するテスト（Issue #131）．

結論文が実際に始まる位置は、本文プレースホルダに流し込んだ**空行の数**で決まる。
帯の高さをそれとは別の値（``band_h``）から取ると、空行数の切り捨てぶんだけ
帯のほうが下まで伸び、表が結論文の上に重なる。**警告も出ない**ので、
pptx を開くまで気づけない。

実測（cn2026-05「まとめ」右カラム）では、表の下端 5.99in に対して
結論文の開始位置が約 5.65in——0.34in 重なっていた。

結論文の y は python-pptx から直接は読めない（実際の行組みは PowerPoint がやる）ので、
**render と同じ規則**——枠の上端から「本文標準サイズ×1.32」の行を積む——で求める。
"""
from __future__ import annotations

import pytest
from pptx import Presentation
from pptx.util import Emu, Pt

from md2pptx import render
from md2pptx.parser import parse


_FM = "---\ntheme: t.pptx\n---\n\n"

_TABLE = """| 型 | プロトコル | 使う場所 |
|:--|:--|:--|
| 距離ベクトル | RIP | AS 内（小規模） |
| リンク状態 | OSPF | AS 内（大規模） |
| パスベクトル | BGP | AS 間 |
"""


def _build(tmp_path, src):
    theme = tmp_path / "theme.pptx"
    Presentation().save(str(theme))
    out = tmp_path / "out.pptx"
    r = render.Renderer(str(theme))
    r.render(parse(src))
    r.save(str(out))
    return Presentation(str(out)), int(Pt(r._body_font_size()) * 1.32)


def _objects(slide):
    """表・図（プレースホルダでないシェイプ）を返す．"""
    return [sh for sh in slide.shapes if not sh.is_placeholder]


def _conclusion_top(slide, line_h, needle="結論文"):
    """結論文の描き始めの y（EMU）．枠の上端＋段落番号×行高．"""
    for sh in slide.shapes:
        if not sh.has_text_frame or sh == slide.shapes.title:
            continue
        for i, p in enumerate(sh.text_frame.paragraphs):
            if needle in p.text:
                return sh, sh.top + i * line_h
    pytest.fail(f"結論文 {needle!r} が本文に無い")


def _assert_above(slide, line_h):
    objs = _objects(slide)
    assert objs, "表・図が描かれていない"
    _, concl_top = _conclusion_top(slide, line_h)
    bottom = max(o.top + o.height for o in objs)
    assert bottom <= concl_top, (
        f"帯の下端 {Emu(bottom).inches:.2f}in が "
        f"結論文 {Emu(concl_top).inches:.2f}in に食い込んでいる")


# ---------------------------------------------------------------- 単一カラム

def test_table_stops_above_the_conclusion(tmp_path):
    """表の下端が、結論文の描き始めより上にある．"""
    prs, line_h = _build(
        tmp_path, _FM + "### 表\n\n導入文\n\n" + _TABLE + "\n→ 結論文\n")
    _assert_above(prs.slides[-1], line_h)


def test_figure_stops_above_the_conclusion(tmp_path):
    """フロー図でも同じ（帯の高さの求め方は表と共通）．"""
    prs, line_h = _build(
        tmp_path,
        _FM + "### 図\n\n導入文\n\n```flow\n[#a A] -> [#b B]\n```\n\n→ 結論文\n")
    _assert_above(prs.slides[-1], line_h)


def test_no_intro_line(tmp_path):
    """導入文が無く結論文だけの場合（``nb == 0``）も食い込まない．"""
    prs, line_h = _build(tmp_path, _FM + "### 表\n\n" + _TABLE + "\n→ 結論文\n")
    _assert_above(prs.slides[-1], line_h)


# ---------------------------------------------------------------- 多カラム

def test_table_in_a_column_stops_above_the_conclusion(tmp_path):
    """2カラムの右側に置いた表でも食い込まない（Issue #131 の再現形）．"""
    prs, line_h = _build(
        tmp_path,
        _FM + "### まとめ\n\n- 左の話\n\n<!-- @col -->\n\n" + _TABLE + "\n→ 結論文\n")
    _assert_above(prs.slides[-1], line_h)


# ---------------------------------------------------------------- 削りすぎない

def test_the_band_keeps_most_of_the_frame(tmp_path):
    """食い込みを直すために帯を削りすぎていない（1行ぶん程度に収まる）．"""
    prs, line_h = _build(
        tmp_path, _FM + "### 表\n\n導入文\n\n" + _TABLE + "\n→ 結論文\n")
    slide = prs.slides[-1]
    tbl, = _objects(slide)
    body, _ = _conclusion_top(slide, line_h)
    assert tbl.height >= body.height - 3 * line_h


# ---------------------------------------------------------------- overflow

def test_overflow_may_still_extend_below(tmp_path):
    """``@overflow: true`` は従来どおり下へはみ出せる（この修正の対象外）．"""
    src = (_FM + "### 表\n<!-- @overflow: true -->\n\n導入文\n\n"
           + _TABLE + "\n→ 結論文\n")
    prs, line_h = _build(tmp_path, src)
    tbl, = _objects(prs.slides[-1])
    assert tbl.height > 0
