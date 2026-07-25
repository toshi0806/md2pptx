#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""変換器の選び方を固定するテスト（Issue #46 / #49）．

守りたいのは **auto が探索だけを行い、失敗の肩代わりをしない**こと．
`pdf.convert` から呼ばれるバックエンドを差し替えるので，PowerPoint も
LibreOffice も要らず，外部プロセスを一切起こさない．

このテストが無いと壊れても気づけない：`except _Unavailable` を
`except PdfError` の後ろへ動かしても mypy は通り，`example.md` の生成も
PDF のページ数も変わらない．気づけるのは「PowerPoint が失敗する環境で
出てきた PDF の組版が違う」ときだけで，それは #46 が無くそうとした
見つけにくさそのもの．
"""
from __future__ import annotations

from pathlib import Path
from types import SimpleNamespace

import pytest

from md2pptx import pdf


@pytest.fixture
def deck(tmp_path):
    """変換の入力に見せかける pptx（中身は読まれない）．"""
    src = tmp_path / "slide.pptx"
    src.write_bytes(b"not really a pptx")
    return src


@pytest.fixture
def called(monkeypatch):
    """どのバックエンドが呼ばれたかを記録し，好きな結果を返させる．

    使い方: ``called.setup(powerpoint=..., libreoffice=...)``．値が例外なら
    送出し，そうでなければ「成功して PDF を書いた」ことにする．
    """
    log: list[str] = []

    def setup(powerpoint=None, libreoffice=None):
        def make(name, outcome):
            def backend(src, dst, timeout=None):
                log.append(name)
                if isinstance(outcome, Exception):
                    raise outcome
                Path(dst).write_bytes(b"%PDF-1.4 " + name.encode())
            return backend

        monkeypatch.setattr(pdf, "_convert_powerpoint", make("powerpoint", powerpoint))
        monkeypatch.setattr(pdf, "_convert_libreoffice", make("libreoffice", libreoffice))

    return SimpleNamespace(log=log, setup=setup)


def test_auto_skips_a_converter_that_is_not_installed(deck, tmp_path, called):
    """**無い**ものは飛ばす——何も失われないので黙って次へ進んでよい．"""
    called.setup(powerpoint=pdf._Unavailable("no PowerPoint here"), libreoffice=None)

    dst = tmp_path / "out.pdf"
    pdf.convert(str(deck), str(dst), "auto", timeout=1)

    assert called.log == ["powerpoint", "libreoffice"]
    assert dst.read_bytes().endswith(b"libreoffice")


def test_auto_stops_when_an_installed_converter_fails(deck, tmp_path, called):
    """**在る**ものの失敗は握らない（Issue #46）．

    ここで LibreOffice へ落ちると，忠実度という成果物の性質が黙って
    入れ替わり，利用者が直せる原因（承認の拒否など）も隠れてしまう．
    """
    called.setup(powerpoint=pdf.PdfError("powerpoint failed: boom"), libreoffice=None)

    dst = tmp_path / "out.pdf"
    with pytest.raises(pdf.PdfError) as excinfo:
        pdf.convert(str(deck), str(dst), "auto", timeout=1)

    assert called.log == ["powerpoint"], "LibreOffice へ落ちてはいけない"
    assert "boom" in str(excinfo.value)
    assert not dst.exists(), "失敗したのに PDF が残ってはいけない"


def test_failure_points_at_libreoffice_only_when_it_is_there(deck, tmp_path, called,
                                                            monkeypatch):
    """案内は行き先があるときだけ．無い物を勧めない．"""
    called.setup(powerpoint=pdf.PdfError("powerpoint failed: boom"))

    monkeypatch.setattr(pdf, "_which_libreoffice", lambda: "/usr/bin/soffice")
    with pytest.raises(pdf.PdfError) as found:
        pdf.convert(str(deck), str(tmp_path / "a.pdf"), "auto", timeout=1)
    assert "--pdf-converter libreoffice" in str(found.value)

    monkeypatch.setattr(pdf, "_which_libreoffice", lambda: None)
    with pytest.raises(pdf.PdfError) as missing:
        pdf.convert(str(deck), str(tmp_path / "b.pdf"), "auto", timeout=1)
    assert "--pdf-converter libreoffice" not in str(missing.value)


def test_auto_reports_both_reasons_when_nothing_is_installed(deck, tmp_path, called):
    """どちらも無いときだけ「変換器が無い」と言う．"""
    called.setup(powerpoint=pdf._Unavailable("no PowerPoint here"),
                 libreoffice=pdf._Unavailable("no LibreOffice here"))

    with pytest.raises(pdf.PdfError) as excinfo:
        pdf.convert(str(deck), str(tmp_path / "out.pdf"), "auto", timeout=1)

    message = str(excinfo.value)
    assert "no PDF converter available" in message
    assert "no PowerPoint here" in message and "no LibreOffice here" in message


@pytest.mark.parametrize("name", ["powerpoint", "libreoffice"])
def test_naming_a_converter_reports_that_it_is_missing(deck, tmp_path, called, name):
    """名指しなら「無い」もそのまま利用者へ届く（勝手に別の物を使わない）．"""
    called.setup(powerpoint=pdf._Unavailable("no PowerPoint here"),
                 libreoffice=pdf._Unavailable("no LibreOffice here"))

    with pytest.raises(pdf.PdfError) as excinfo:
        pdf.convert(str(deck), str(tmp_path / "out.pdf"), name, timeout=1)

    assert called.log == [name]
    assert "no PDF converter available" not in str(excinfo.value)


class TestCustomCommand:
    """任意コマンド指定．ツールが PDF をどこへ書くかは指定形式で決まる．"""

    def _run(self, monkeypatch, command, src, dst, writes):
        """変換器コマンドを実行せず，``writes(cmd)`` で成果物を作らせる．"""
        seen: list[list[str]] = []

        def fake_run(cmd, what, input=None, timeout=None):
            seen.append(cmd)
            writes(cmd)

        monkeypatch.setattr(pdf, "_run", fake_run)
        pdf.convert(str(src), str(dst), command, timeout=1)
        return seen[0]

    def test_output_placeholder_is_written_in_place(self, deck, tmp_path, monkeypatch):
        dst = tmp_path / "named.pdf"
        cmd = self._run(monkeypatch, "mytool -o {output} {input}", deck, dst,
                        lambda cmd: Path(cmd[2]).write_bytes(b"%PDF"))
        assert cmd == ["mytool", "-o", str(dst), str(deck)]
        assert dst.exists()

    def test_outdir_placeholder_collects_the_input_basename(self, deck, tmp_path,
                                                            monkeypatch):
        """soffice 方式：出力先ディレクトリに <入力 basename>.pdf を書く．"""
        dst = tmp_path / "renamed.pdf"
        self._run(monkeypatch, "soffice --outdir {outdir} {input}", deck, dst,
                  lambda cmd: (tmp_path / "slide.pdf").write_bytes(b"%PDF"))
        assert dst.exists(), "入力名の PDF を出力先の名前へ揃えること"
        assert not (tmp_path / "slide.pdf").exists()

    def test_a_tool_without_placeholders_gets_the_input_appended(self, deck, tmp_path,
                                                                 monkeypatch):
        """出力先を取れないツール：入力を末尾に足し，その隣の PDF を回収する．"""
        dst = tmp_path / "out.pdf"
        cmd = self._run(monkeypatch, "convert-to-pdf", deck, dst,
                        lambda cmd: deck.with_suffix(".pdf").write_bytes(b"%PDF"))
        assert cmd == ["convert-to-pdf", str(deck)]
        assert dst.exists()

    def test_a_tool_that_writes_nothing_is_a_failure(self, deck, tmp_path, monkeypatch):
        dst = tmp_path / "out.pdf"
        with pytest.raises(pdf.PdfError):
            self._run(monkeypatch, "mytool {output}", deck, dst, lambda cmd: None)
