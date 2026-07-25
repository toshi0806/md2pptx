#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""出力 PDF の差し替え方を固定するテスト．

守りたいのは **変換中も出力 PDF が一瞬たりとも消えない**こと．編集しながら見る運用では
PDF ビューアがファイルを監視しているが，多くの実装は削除を検知するとそのファイルを監視
から外す（LaTeX Workshop は 250ms で確定し，フォルダの監視ごと破棄する）．変換には 1 秒
から数秒かかるので，「消してから書く」実装だと**最初のリビルドでプレビューが死ぬ**．

このテストが無いと壊れても気づけない：`convert` の先頭で `os.remove(dst)` に戻しても
mypy は通り，`example.md` の生成も PDF のページ数も変わらない．気づけるのは
「保存しても PDF タブが更新されなくなった」ときで，原因が md2pptx 側にあるとは
まず思い当たらない．

同時に，差し替えに移る前から守ってきた契約——**失敗したら古い PDF を残さない**——も
ここで固定する（PDF 変換の失敗は終了コードを変えないので，警告を見落とした人が前回の
内容を新しい出力と取り違えてしまう）．

`pdf.convert` から呼ばれるバックエンドを差し替えるので，外部プロセスは一切起こさない．
"""
from __future__ import annotations

from pathlib import Path
from types import SimpleNamespace

import pytest

from md2pptx import pdf

OLD = b"%PDF-1.4 old"
NEW = b"%PDF-1.4 new"


@pytest.fixture
def deck(tmp_path):
    """変換の入力に見せかける pptx（中身は読まれない）．"""
    src = tmp_path / "slide.pptx"
    src.write_bytes(b"not really a pptx")
    return src


@pytest.fixture
def backend(monkeypatch, tmp_path):
    """変換器のふりをして，変換の**最中**に見えていたものを記録する．

    ``setup(action)`` の ``action(dst)`` が 1 回の変換の中身．呼ばれた時点で
    最終出力（``tmp_path/out.pdf``）がどう見えていたかを ``seen`` に残す．
    """
    state = SimpleNamespace(seen=None, staged=None)

    def setup(action):
        def fake(src, dst, timeout=None, attended=True):
            final = tmp_path / "out.pdf"
            state.seen = final.read_bytes() if final.exists() else None
            state.staged = Path(dst)
            return action(dst)

        monkeypatch.setattr(pdf, "_convert_powerpoint", fake)

    state.setup = setup
    return state


def test_the_existing_pdf_stays_readable_while_converting(deck, tmp_path, backend):
    """変換中も前回の PDF がそのまま読める（ビューアが削除を見ない）．"""
    dst = tmp_path / "out.pdf"
    dst.write_bytes(OLD)
    backend.setup(lambda staged: Path(staged).write_bytes(NEW))

    pdf.convert(str(deck), str(dst), "powerpoint", timeout=1)

    assert backend.seen == OLD, "変換中に出力が消えても差し替わってもいけない"
    assert dst.read_bytes() == NEW, "変換後は新しい内容になること"


def test_the_converter_writes_somewhere_else(deck, tmp_path, backend):
    """変換器に渡すのは最終パスではない．

    毎回まっさらな場所へ書かせることで，「無音失敗した変換器が残した前回の PDF を
    成功と誤判定する」問題が構造的に消える（以前は先に消すことで防いでいた）．
    """
    dst = tmp_path / "out.pdf"
    dst.write_bytes(OLD)
    backend.setup(lambda staged: Path(staged).write_bytes(NEW))

    pdf.convert(str(deck), str(dst), "powerpoint", timeout=1)

    assert backend.staged != dst
    assert backend.staged.name == dst.name, "名前は保つ（出力名を見るツールがいる）"
    assert not backend.staged.exists(), "作業場所は残さない"


def test_a_failed_conversion_leaves_no_stale_pdf(deck, tmp_path, backend):
    """失敗したら古い PDF は残さない（差し替え以前からの契約）．"""
    dst = tmp_path / "out.pdf"
    dst.write_bytes(OLD)

    def explode(staged):
        raise pdf.PdfError("powerpoint failed: boom")

    backend.setup(explode)

    with pytest.raises(pdf.PdfError):
        pdf.convert(str(deck), str(dst), "powerpoint", timeout=1)

    assert not dst.exists()


def test_a_half_written_pdf_never_reaches_the_output(deck, tmp_path, backend):
    """打ち切られた変換の書きかけは，作業場所ごと捨てる．"""
    dst = tmp_path / "out.pdf"

    def half_then_fail(staged):
        Path(staged).write_bytes(b"%PDF-1.4 trunc")
        raise pdf.PdfError("powerpoint timed out after 180s")

    backend.setup(half_then_fail)

    with pytest.raises(pdf.PdfError):
        pdf.convert(str(deck), str(dst), "powerpoint", timeout=1)

    assert not dst.exists()


@pytest.mark.parametrize("succeeds", [True, False])
def test_no_working_directory_is_left_behind(deck, tmp_path, backend, succeeds):
    """成功・失敗のどちらでも作業ディレクトリを残さない．"""
    dst = tmp_path / "out.pdf"

    def action(staged):
        if not succeeds:
            raise pdf.PdfError("boom")
        Path(staged).write_bytes(NEW)

    backend.setup(action)

    try:
        pdf.convert(str(deck), str(dst), "powerpoint", timeout=1)
    except pdf.PdfError:
        pass

    leftovers = [p.name for p in tmp_path.iterdir() if p.name.startswith(".md2pptx-")]
    assert leftovers == []


def test_an_outdir_converter_does_not_write_over_the_output(deck, tmp_path, monkeypatch):
    """``--outdir`` 方式の変換器に最終出力のあるディレクトリを触らせない．

    LibreOffice は ``--outdir`` へ ``<入力 basename>.pdf`` を書く．``slide.pptx`` →
    ``slide.pdf`` という既定の組み合わせでは**それが最終出力そのもの**なので，作業場所を
    挟まないと変換器が出力を直接・逐次的に書くことになる（ビューアが書きかけを読む）．
    """
    dst = tmp_path / "slide.pdf"        # 入力の basename と一致させるのが要点
    dst.write_bytes(OLD)
    seen: dict[str, object] = {}

    def fake_run(cmd, what, input=None, timeout=None):
        outdir = Path(cmd[cmd.index("--outdir") + 1])
        seen["outdir"] = outdir
        seen["final_during"] = dst.read_bytes()
        (outdir / "slide.pdf").write_bytes(NEW)

    monkeypatch.setattr(pdf, "_which_libreoffice", lambda: "/usr/bin/soffice")
    monkeypatch.setattr(pdf, "_run", fake_run)

    pdf.convert(str(deck), str(dst), "libreoffice", timeout=1)

    assert seen["outdir"] != tmp_path, "最終出力のあるディレクトリを渡してはいけない"
    assert seen["final_during"] == OLD
    assert dst.read_bytes() == NEW


def test_a_failed_replacement_leaves_no_stale_pdf(deck, tmp_path, backend, monkeypatch):
    """置き換えに失敗したときも古い PDF は残さない．

    変換自体は成功していても，そこに残っているのは前回の内容．終了コードは 0 なので，
    警告を見落とした人が古い PDF を新しい出力と取り違えるのは変換が失敗したときと同じ．
    """
    dst = tmp_path / "out.pdf"
    dst.write_bytes(OLD)
    backend.setup(lambda staged: Path(staged).write_bytes(NEW))

    real_replace = pdf.os.replace

    def refuse(source, target):
        # 目的の PDF への置き換えだけを失敗させる．os.replace を丸ごと差し替えると
        # このテストの間だけとはいえ無関係な処理まで巻き込む．
        if str(target) == str(dst):
            raise OSError("device busy")
        return real_replace(source, target)

    monkeypatch.setattr(pdf.os, "replace", refuse)

    with pytest.raises(pdf.PdfError, match="cannot replace existing PDF"):
        pdf.convert(str(deck), str(dst), "powerpoint", timeout=1)

    assert not dst.exists(), "置き換えに失敗しても古い PDF を残してはいけない"


def test_an_unusable_output_directory_is_a_pdf_failure(deck, tmp_path, monkeypatch):
    """作業場所を作れないときも ``PdfError``——素の ``OSError`` を通さない．

    通してしまうと cli の ``except PdfError`` をすり抜けて ``SystemExit`` になり，
    **pptx は保存できているのに終了コードが 1** になる（`main` が ``PermissionError``
    を握って整形するため）．「PDF が作れなくても pptx は成功」（Issue #39）が崩れ，
    編集しながらの運用が出力先の権限ひとつで止まってしまう．
    """
    def refuse(*args, **kwargs):
        raise PermissionError(13, "Permission denied")

    monkeypatch.setattr(pdf.workdir, "create", refuse)

    with pytest.raises(pdf.PdfError, match="cannot create a working directory"):
        pdf.convert(str(deck), str(tmp_path / "out.pdf"), "powerpoint", timeout=1)


@pytest.mark.parametrize("converter", ["auto", "powerpoint"])
def test_unattended_reaches_the_backend(deck, tmp_path, monkeypatch, converter):
    """``unattended`` は「止まっても PowerPoint を前面化しない」まで届く．

    見るのは ``convert`` からバックエンドまでの**配線**であって，上限の決め方
    （test_pdf_timeout.py）でも探索順（test_pdf_converter_choice.py）でもない．
    ``auto`` と名指しは ``_dispatch`` の**別々の呼び出し箇所**なので両方を通す
    ——片方だけ ``attended`` を落としても気づけるように．
    """
    seen: dict[str, object] = {}

    def fake(src, dst, timeout=None, attended=True):
        seen["attended"] = attended
        Path(dst).write_bytes(NEW)

    monkeypatch.setattr(pdf, "_convert_powerpoint", fake)

    pdf.convert(str(deck), str(tmp_path / "a.pdf"), converter, timeout=1)
    assert seen["attended"] is True, "既定は従来どおり（tty なら前面化する）"

    pdf.convert(str(deck), str(tmp_path / "b.pdf"), converter, timeout=1,
                unattended=True)
    assert seen["attended"] is False, "unattended なら前面化しないと伝わること"
