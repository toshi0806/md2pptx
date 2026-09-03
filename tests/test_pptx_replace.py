#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""出力 pptx の差し替え方を固定するテスト（Issue #56）．

守りたいのは **保存中も前回の pptx が壊れて見えない**こと．python-pptx は
``zipfile.ZipFile(path, "w")`` で書くので，出力先へ直接保存すると**開いた瞬間に
切り詰められる**——その間に開いた人は 0 バイトか途中までの zip を掴む．PowerPoint で
開いたまま作り直す，``--watch`` と手で叩くのを併用する，といった使い方で普通に踏む．

もう 1 つ固定するのが **失敗したら前回の pptx を残す**こと．ここは PDF
（pptx2pdf の tests/test_pdf_replace.py）と**逆の契約**なので，取り違えて「揃える」
修正が入らない
ように理由ごと残す：PDF 変換の失敗は終了コードを変えない（警告だけ）ので古い PDF が
「新しい出力」に見えてしまうが，pptx の保存失敗は cli が終了コード 1 で終えるため
取り違えようがなく，それなら主成果物を消さない方がよい．

python-pptx の実物は使わず ``Renderer`` の ``prs`` を差し替えるので，テーマも要らず
一瞬で回る．

作業ディレクトリの**片付け方**（消せなくても投げない／黙らない）は ``workdir.discard``
の責務で、pptx2pdf の tests/test_workdir.py が固定する．ここが見るのは「``save`` が
作業場所を経由し，最後に残骸を残さない」という配線まで．
"""
from __future__ import annotations

from pathlib import Path
from types import SimpleNamespace

import pytest

from md2pptx import render

OLD = b"PK\x03\x04 old pptx"
NEW = b"PK\x03\x04 new pptx"


@pytest.fixture
def renderer(tmp_path):
    """``save`` だけを取り出した Renderer（描画も base pptx も通さない）．

    ``prs`` は「保存を頼まれたら書く」だけの替え玉．保存の**最中**に出力先がどう
    見えていたかを ``seen`` に残す．

    出力先は ``dst`` として持たせる．テスト側で別名を組み立てると，替え玉が見張る
    ファイルとずれても**気づけないまま通ってしまう**ので，1 か所を正にする．
    """
    r = render.Renderer.__new__(render.Renderer)      # __init__ は base pptx を開く
    state = SimpleNamespace(dst=tmp_path / "out.pptx", seen=None, staged=None,
                            action=None)

    def fake_save(where):
        # 本物（``zipfile.ZipFile(where, "w")``）と同じ **2 段階**で動く：開いた瞬間に
        # 切り詰め，それから中身を書く．この順序こそが「保存中は壊れて見える」の正体
        # なので，替え玉が一気に書いてしまうと，出力先へ直接保存する実装でもテストが
        # 通ってしまう（実際それで素通りしていた）．
        Path(where).write_bytes(b"")
        state.seen = state.dst.read_bytes() if state.dst.exists() else None
        state.staged = Path(where)
        if state.action is not None:
            state.action(where)
        else:
            Path(where).write_bytes(NEW)

    r.prs = SimpleNamespace(save=fake_save)
    state.renderer = r
    state.save = lambda: r.save(str(state.dst))
    return state


def test_the_existing_pptx_stays_readable_while_saving(renderer):
    """保存中も前回の pptx がそのまま読める．

    直接書くと ``ZipFile(path, "w")`` が開いた瞬間に切り詰めるので，ここが崩れると
    「作り直している間だけ壊れた pptx が見える」状態に戻る．
    """
    dst = renderer.dst
    dst.write_bytes(OLD)

    renderer.save()

    assert renderer.seen == OLD, "保存中に出力が消えても切り詰められてもいけない"
    assert dst.read_bytes() == NEW, "保存後は新しい内容になること"


def test_the_save_goes_somewhere_else(renderer):
    """python-pptx に渡すのは最終パスではない（名前は保つ）．"""
    dst = renderer.dst
    dst.write_bytes(OLD)

    renderer.save()

    assert renderer.staged != dst
    assert renderer.staged.name == dst.name
    assert not renderer.staged.exists(), "作業場所は残さない"


def test_a_failed_save_keeps_the_previous_pptx(renderer):
    """**失敗したら前回の pptx を残す**——PDF とは逆の契約．

    pptx の保存失敗は cli が ``BuildError`` にして終了コード 1 で終えるので，古い
    ファイルが「新しい出力」と取り違えられることはない．それなら PowerPoint で開いて
    いるかもしれない主成果物を消さない方がよい．
    """
    dst = renderer.dst
    dst.write_bytes(OLD)

    def explode(where):
        raise ValueError("boom")

    renderer.action = explode

    with pytest.raises(ValueError):
        renderer.save()

    assert dst.read_bytes() == OLD, "失敗したのに前回の pptx を失ってはいけない"


def test_a_half_written_pptx_never_reaches_the_output(renderer):
    """途中まで書いて落ちた pptx は，作業場所ごと捨てる．"""
    dst = renderer.dst
    dst.write_bytes(OLD)

    def half_then_fail(where):
        Path(where).write_bytes(b"PK\x03\x04 trunc")
        raise OSError("disk full")

    renderer.action = half_then_fail

    with pytest.raises(OSError):
        renderer.save()

    assert dst.read_bytes() == OLD


@pytest.mark.parametrize("succeeds", [True, False])
def test_no_working_directory_is_left_behind(renderer, tmp_path, succeeds):
    """成功・失敗のどちらでも作業ディレクトリを残さない．"""
    dst = renderer.dst

    def action(where):
        if not succeeds:
            raise ValueError("boom")
        Path(where).write_bytes(NEW)

    renderer.action = action

    try:
        renderer.save()
    except ValueError:
        pass

    leftovers = [p.name for p in tmp_path.iterdir() if p.name.startswith(".md2pptx-")]
    assert leftovers == []


def test_an_unusable_output_directory_says_what_failed(renderer, monkeypatch):
    """作業場所を作れないときは，何をしようとして失敗したかを添える．

    素の errno だけだと，利用者には見覚えのない一時ディレクトリ名しか残らない．
    元の例外は ``__cause__`` に残す——errno を見たいときの手掛かりを捨てない．
    """
    def refuse(*args, **kwargs):
        raise PermissionError(13, "Permission denied")

    monkeypatch.setattr(render.workdir, "create", refuse)

    with pytest.raises(OSError, match="cannot create a working directory") as found:
        renderer.save()

    assert isinstance(found.value.__cause__, PermissionError)
    assert found.value.__cause__.errno == 13


def test_a_first_save_needs_no_existing_file(renderer):
    """初回（出力先がまだ無い）でも普通に書ける．"""
    dst = renderer.dst

    renderer.save()

    assert dst.read_bytes() == NEW
    assert renderer.seen is None
