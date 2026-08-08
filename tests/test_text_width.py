#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""文字幅の概算を固定するテスト（Issue #156）．

日本語のプロポーショナルフォント（cn2026-theme の BIZ UDPゴシック、Yu Gothic、
Meiryo など）では**かなは 1em より狭い**。全角をすべて 1em で数えていたので
折り返しを多く見積もり、``{box}`` の枠が 2 行ぶんの高さになって次の項目まで
覆っていた（cn2026-02 p.2 の「ネットワークコミュニケーション」）。

cn2026-theme・30pt で PDF から実測した文字送り：

| 種類 | 1文字 | em 比 |
|---|--:|--:|
| カタカナ | 26.66pt | 0.889 |
| 漢字 | 29.87pt | 0.996 |
| ASCII | 15.65pt | 0.522 |
"""
from __future__ import annotations

import pytest

from md2pptx.render import Renderer


def _w(text, pt=30.0):
    return Renderer._text_width_pt(text, pt)


# ---------------------------------------------------------------- 実測に合う

@pytest.mark.parametrize("text,per_char", [
    ("ネットワークコミュニケーシ", 26.66),   # カタカナ
    ("ひらがなのれんしゅうです", 26.66),     # ひらがなも同じ扱い
    ("情報科学科計算機通信網講義", 29.87),   # 漢字
    ("Networkcommunication", 15.65),         # ASCII
])
def test_the_estimate_matches_the_measurement(text, per_char):
    """実測との差を 1 文字あたり 1.5pt 以内に収める（30pt の場合）．"""
    assert _w(text, 30.0) / len(text) == pytest.approx(per_char, abs=1.5)


def test_kana_is_narrower_than_kanji():
    assert _w("アイウエオ") < _w("亜以宇江於")


def test_the_width_scales_with_the_font():
    assert _w("あいうえお", 15.0) == pytest.approx(_w("あいうえお", 30.0) / 2)


# ------------------------------------------------------------ 折り返しの判定

# ------------------------------------------------------------ 折り返しの行数

def test_a_line_that_fits_is_not_wrapped():
    """cn2026-02 のシラバスの項目．

    実測の幅は 397.9pt、カラムで使える幅は 382.6pt——4.0% 超えているのに
    PowerPoint は 1 行で出す（日本語を詰めて改行を避ける）。ここを 2 行と
    決めつけると、``{box}`` の枠が次の項目まで覆う。
    """
    assert Renderer._wrapped_lines("ネットワークコミュニケーション", 30.0, 382.6) == 1


def test_a_long_line_still_wraps():
    """許容幅を入れても、本当に折り返す行は折り返すと数える．"""
    assert Renderer._wrapped_lines(
        "コンピュータネットワークとインターネット", 30.0, 382.6) == 2


def test_three_lines(tmp_path=None):
    assert Renderer._wrapped_lines("あ" * 45, 30.0, 382.6) == 3


def test_an_empty_width_does_not_divide_by_zero():
    assert Renderer._wrapped_lines("あいう", 30.0, 0.0) == 1


# ---------------------------------------------------------------- 端の扱い

def test_an_empty_text_is_zero():
    assert _w("") == 0.0
    assert _w(None) == 0.0


def test_the_long_vowel_mark_counts_as_kana():
    """``ー``（U+30FC）はカタカナの並びの一部——1em では広すぎる．"""
    assert _w("ー", 30.0) < 30.0


def test_full_width_symbols_stay_at_one_em():
    """``＠`` や ``（）`` のような全角記号は詰めない（かなだけを狭くする）．"""
    assert _w("＠", 30.0) == pytest.approx(30.0)
