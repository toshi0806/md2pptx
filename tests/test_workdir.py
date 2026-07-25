#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""作業ディレクトリの片付け方を固定するテスト（Issue #58）．

守りたいのは 2 つで、**どちらか片方だけでは意味がない**組み合わせ。

- **片付けの失敗で処理の成否を変えない。** 片付けに入る時点で保存や変換は終わっている。
  ここで例外を投げると成功した実行が失敗になり、本体が投げた例外があれば置き換えてしまう。
- **それでも黙って残さない。** 消せなかったと誰も知らないと ``--watch`` では保存のたびに
  溜まり、しかも出力先ディレクトリのものは利用者の目に触れる。

この 2 つを同時に満たす形は「消せるだけ消して、残っていたら言う」しかない。片方に寄せた
実装（例外を投げる／黙る）に戻すと、下のどれかが落ちる。
"""
from __future__ import annotations

import os

import pytest

from md2pptx import workdir


def test_it_removes_the_directory_and_its_contents(tmp_path):
    """中身ごと消す（保存が途中で落ちた場合は書きかけが入っている）。"""
    work = workdir.create(str(tmp_path))
    (tmp_path / os.path.basename(work) / "half-written.pptx").write_bytes(b"...")

    workdir.discard(work)

    assert not os.path.exists(work)


@pytest.mark.skipif(hasattr(os, "geteuid") and os.geteuid() == 0,
                    reason="root は書き込み不可のディレクトリにも書けるので再現できない")
def test_a_directory_that_cannot_be_removed_is_reported_not_raised(tmp_path, capsys):
    """**本当に消せない**状況で、黙らず・投げないことを同時に確かめる．

    ``rmtree`` を差し替えるのではなく親ディレクトリを書き込み不可にして、実物の
    ``shutil.rmtree(..., ignore_errors=True)`` に当てる——差し替えると
    ``ignore_errors`` を無視するスタブになりがちで、実物ではありえない経路を
    確かめてしまう（実際それで一度落ちた）．
    """
    work = workdir.create(str(tmp_path))
    open(os.path.join(work, "half-written.pptx"), "wb").close()
    os.chmod(tmp_path, 0o500)                # 中身を消せなくする
    try:
        workdir.discard(work)                # 投げないこと（例外が出ればテストが落ちる）
    finally:
        os.chmod(tmp_path, 0o700)

    assert os.path.isdir(work), "前提: この状況では実際に消せない"
    assert f"could not remove {work}" in capsys.readouterr().err


def test_a_quiet_success_says_nothing(tmp_path, capsys):
    """普通に消せたときは何も出さない（毎回の実行で雑音を出さない）．"""
    workdir.discard(workdir.create(str(tmp_path)))

    assert capsys.readouterr().err == ""


def test_it_can_build_beside_a_given_directory(tmp_path):
    """出力先の隣に作れる——``os.replace`` を同一ファイルシステム内に収めるため．"""
    work = workdir.create(str(tmp_path))

    assert os.path.dirname(work) == str(tmp_path)
    assert os.path.basename(work).startswith(".md2pptx-")
    assert os.stat(work).st_dev == os.stat(tmp_path).st_dev


def test_a_creation_failure_is_left_to_the_caller(tmp_path, monkeypatch):
    """作成の失敗はそのまま送出する．

    「作業場所を用意できなかった」ことをどう伝えるかは呼び出し側で違う
    （``pdf.convert`` は PdfError にして終了コードを変えず，``render.save`` は
    何をしようとして失敗したかを添える）ので、ここでは整形しない．
    """
    def refuse(*args, **kwargs):
        raise PermissionError(13, "Permission denied")

    monkeypatch.setattr(workdir.tempfile, "mkdtemp", refuse)

    with pytest.raises(PermissionError):
        workdir.create(str(tmp_path))
