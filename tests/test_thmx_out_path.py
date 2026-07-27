#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""変換に失敗したとき何を消して何を残すかを固定するテスト（Issue #35 の途中で追加）．

``thmx_to_pptx`` は出力先を省略できる．省略されたら一時ファイルを自分で作るので，
**失敗したらそれは自分で片付ける**——空の pptx が ``/tmp`` に溜まっていくのは
呼び出し側からは見えない．逆に**出力先を渡されたときは触らない**．そこにある
ファイルは呼び出し側のもので、変換が失敗したからといって消してよい理由がない．

この 2 つを分けているのが ``created_out`` で、#35 で注釈を入れる際に
「``out_path`` はこの行以降 None ではない」という不変条件が変数の中にしか
無いことが分かり、分岐の条件を ``out_path`` 自身へ変えた．振る舞いは同じだが、
**同じであること自体が誰にも確かめられていなかった**のでここで固定する．

壊れた zip を入力にするのは、``tempfile`` を作る行を通り抜けてから失敗させる
ためで、そうしないと片付けの経路に入らない（入力が無い場合はその手前で返る）．
"""
from __future__ import annotations

import glob
import os
import tempfile

import pytest

from md2pptx.thmx2pptx import ThmxError, thmx_to_pptx


@pytest.fixture
def broken_thmx(tmp_path):
    """zip として開けないファイル．展開の直前まで進んでから失敗する．"""
    path = tmp_path / "broken.thmx"
    path.write_text("これは zip ではない")
    return str(path)


def _stray_temps():
    return set(glob.glob(os.path.join(tempfile.gettempdir(),
                                      "md2pptx-base-*.pptx")))


def test_it_cleans_up_the_temp_file_it_made_itself(broken_thmx):
    """出力先を省略した場合，失敗しても一時ファイルを残さない．"""
    before = _stray_temps()

    with pytest.raises(ThmxError, match="not a valid thmx"):
        thmx_to_pptx(broken_thmx)

    assert _stray_temps() - before == set()


def test_it_does_not_touch_a_file_the_caller_named(broken_thmx, tmp_path):
    """出力先を渡された場合，失敗してもそのファイルには手を出さない．

    消すと，前回の変換結果を出力先にしていた人が**失敗しただけで成果物を失う**．
    """
    out = tmp_path / "mine.pptx"
    out.write_text("呼び出し側が用意した中身")

    with pytest.raises(ThmxError, match="not a valid thmx"):
        thmx_to_pptx(broken_thmx, str(out))

    assert out.read_text() == "呼び出し側が用意した中身"


def test_a_missing_input_fails_before_making_anything(tmp_path):
    """入力が無ければ一時ファイルを作る前に返る（作ってから消すより素直）．"""
    before = _stray_temps()

    with pytest.raises(ThmxError, match="thmx not found"):
        thmx_to_pptx(str(tmp_path / "nope.thmx"))

    assert _stray_temps() - before == set()
