#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""行の高さと段落前アキの数え方を固定するテスト（Issue #152）．

**PowerPoint で実測した値に合わせる。** cn2026-theme（BIZ UDPゴシック・lvl1 30pt）で
``spcBef`` だけを変えたテーマを 3 つ作り、``@autofit: 100``（縮小なし）で
段落の送りを PDF から測ったところ：

| ``spcPct`` | 段落の送り |
|--:|--:|
| 0 | 36.0pt |
| 20000 | 43.2pt |
| 40000 | 50.4pt |

- 行の高さ ＝ フォントサイズ × **1.20**（36.0 / 30）
- ``spcPct`` は「**行の高さ**」に対する割合（7.2 = 0.20 × 36.0、14.4 = 0.40 × 36.0）

以前は 1.32 と「フォントサイズに対する割合」で、lvl1 の段落 1 つあたり 5.6% 大きく
数えていた。ずれる向きは常に下なので、``{box}`` の枠が段落の下半分から次の項目へ
掛かっていた（cn2026-02 p.2）。
"""
from __future__ import annotations

import pytest
from pptx import Presentation
from pptx.oxml.ns import qn
from pptx.util import Pt

from md2pptx import render


def _theme(tmp_path, spc_pct=None, spc_pts=None, size=3000):
    tmp_path.mkdir(parents=True, exist_ok=True)   # サブディレクトリを渡すため
    prs = Presentation()
    body = prs.slide_masters[0].element.find(
        qn("p:txStyles") + "/" + qn("p:bodyStyle"))
    for lvl in range(1, 10):
        el = body.find(qn(f"a:lvl{lvl}pPr"))
        if el is None:
            continue
        for tag in ("a:spcBef", "a:defRPr"):
            old = el.find(qn(tag))
            if old is not None:
                el.remove(old)
        if spc_pct is not None or spc_pts is not None:
            spc = el.makeelement(qn("a:spcBef"), {})
            spc.append(spc.makeelement(
                qn("a:spcPct") if spc_pct is not None else qn("a:spcPts"),
                {"val": str(spc_pct if spc_pct is not None else spc_pts)}))
            el.insert(0, spc)
        el.append(el.makeelement(qn("a:defRPr"), {"sz": str(size)}))
    path = tmp_path / "theme.pptx"
    prs.save(str(path))
    return path


def _r(tmp_path, **kw):
    return render.Renderer(str(_theme(tmp_path, **kw)))


def _pt(emu):
    return emu / 12700.0


# ------------------------------------------------------------------ 行の高さ

@pytest.mark.parametrize("size", [30.0, 26.0, 22.0, 18.0])
def test_a_line_is_1_20_times_the_font(tmp_path, size):
    r = _r(tmp_path / str(size))
    assert _pt(r._line_height(size)) == pytest.approx(size * 1.20, abs=0.05)


# ------------------------------------------------------------ 段落前アキ（%）

@pytest.mark.parametrize("pct,expected", [(0, 0.0), (20000, 7.2), (40000, 14.4)])
def test_space_before_is_a_share_of_the_line(tmp_path, pct, expected):
    """実測した 3 点をそのまま置く（30pt の段落）．"""
    r = _r(tmp_path / str(pct), spc_pct=pct)
    assert _pt(r._space_before(0, 30.0)) == pytest.approx(expected, abs=0.1)


def test_the_paragraph_advance_matches_the_measurement(tmp_path):
    """段落の送り（行 ＋ アキ）が実測 43.2pt に一致する．"""
    r = _r(tmp_path, spc_pct=20000)
    assert _pt(r._para_height(0, 30.0)) == pytest.approx(43.2, abs=0.2)


def test_the_share_follows_the_font_size(tmp_path):
    """割合なので、字が小さければアキも小さい．"""
    r = _r(tmp_path, spc_pct=20000)
    assert _pt(r._space_before(0, 15.0)) == pytest.approx(3.6, abs=0.1)


# ---------------------------------------------------------- 段落前アキ（絶対）

def test_absolute_space_before_is_taken_as_written(tmp_path):
    """``spcPts`` は 1/100 pt の絶対値——行の高さは掛けない．"""
    r = _r(tmp_path, spc_pts=1200)
    assert _pt(r._space_before(0, 30.0)) == pytest.approx(12.0, abs=0.1)
    assert _pt(r._space_before(0, 15.0)) == pytest.approx(12.0, abs=0.1)


def test_no_space_before_at_all(tmp_path):
    r = _r(tmp_path)
    assert r._space_before(0, 30.0) == 0
    assert _pt(r._para_height(0, 30.0)) == pytest.approx(36.0, abs=0.2)
