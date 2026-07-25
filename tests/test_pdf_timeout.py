#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""待ちの上限とその決め方を固定するテスト（Issue #48 / #49）．

守りたいのは **人が見ているかどうかで待ち方を変える**という判断．承認ダイアログの
ように「人が今すぐ直せる」停止では待つ意味があるので tty では打ち切らず，誰も応答
できない環境（cron / CI / エディタ拡張）でだけ打ち切る．

外部プロセスは起こさない．``subprocess`` を差し替えて挙動だけを見る．
"""
from __future__ import annotations

import subprocess
from pathlib import Path
from types import SimpleNamespace

import pytest

from md2pptx import pdf


@pytest.fixture(autouse=True)
def no_env(monkeypatch):
    """環境変数の影響を受けないようにする（実行環境で結果が変わらないため）．"""
    monkeypatch.delenv(pdf.ENV_TIMEOUT, raising=False)


@pytest.fixture
def stderr_is(monkeypatch):
    """stderr が tty かどうかを差し替える．出力も拾えるようにする．"""
    def use(tty: bool):
        written: list[str] = []
        fake = SimpleNamespace(isatty=lambda: tty, write=written.append)
        monkeypatch.setattr(pdf.sys, "stderr", fake)
        return written
    return use


class TestPolicy:
    """上限の決め方：明示 > 環境変数 > tty かどうか．"""

    def test_a_terminal_waits_for_the_person(self, stderr_is):
        """tty では打ち切らない——30 秒の案内で直してもらえる見込みがある．"""
        stderr_is(True)
        assert pdf._resolve_timeout(None) is None

    def test_nobody_watching_means_give_up(self, stderr_is):
        """非 tty では誰も応答しないので待っても状況は変わらない．"""
        stderr_is(False)
        assert pdf._resolve_timeout(None) == pdf._TIMEOUT_UNATTENDED

    def test_the_environment_overrides_the_default(self, monkeypatch, stderr_is):
        stderr_is(True)          # tty でも環境変数が勝つ
        monkeypatch.setenv(pdf.ENV_TIMEOUT, "42")
        assert pdf._resolve_timeout(None) == 42.0

    def test_the_option_overrides_the_environment(self, monkeypatch, stderr_is):
        stderr_is(False)
        monkeypatch.setenv(pdf.ENV_TIMEOUT, "42")
        assert pdf._resolve_timeout(7.0) == 7.0

    def test_zero_asks_for_no_limit(self, stderr_is):
        stderr_is(False)
        assert pdf._resolve_timeout(0.0) is None

    @pytest.mark.parametrize("bad", [-5.0, float("nan"), float("inf")])
    def test_values_that_are_not_a_duration_are_rejected(self, bad, stderr_is):
        """特に負値：``-5``（``5`` の打ち間違い）を無制限と読むと，この機能が
        防ごうとしている「無人で永久に待つ」状態そのものを作ってしまう．"""
        stderr_is(False)
        with pytest.raises(pdf.PdfError, match="0 = no limit"):
            pdf._resolve_timeout(bad)

    def test_a_junk_environment_value_is_rejected(self, monkeypatch, stderr_is):
        stderr_is(False)
        monkeypatch.setenv(pdf.ENV_TIMEOUT, "soon")
        with pytest.raises(pdf.PdfError, match=pdf.ENV_TIMEOUT):
            pdf._resolve_timeout(None)

    def test_the_limit_reaches_the_converter(self, tmp_path, monkeypatch, stderr_is):
        """決めた上限が実際にバックエンドへ渡ること．"""
        stderr_is(False)
        seen: list[float | None] = []

        def backend(src, dst, timeout=None):
            seen.append(timeout)
            Path(dst).write_bytes(b"%PDF")

        monkeypatch.setattr(pdf, "_convert_libreoffice", backend)
        src = tmp_path / "slide.pptx"
        src.write_bytes(b"x")
        pdf.convert(str(src), str(tmp_path / "out.pdf"), "libreoffice")
        assert seen == [pdf._TIMEOUT_UNATTENDED]


class FakeProc:
    """``communicate`` が最初は時間切れになり，2 回目で返る子プロセス．"""

    def __init__(self, calls_before_return=1, returncode=0):
        self.remaining = calls_before_return
        self.returncode = returncode
        self.timeouts: list[float | None] = []
        self.killed = False

    def communicate(self, input=None, timeout=None):
        self.timeouts.append(timeout)
        if self.remaining > 0:
            self.remaining -= 1
            raise subprocess.TimeoutExpired(cmd="osascript", timeout=timeout or 0)
        return ("", "")

    def kill(self):
        self.killed = True
        self.remaining = 0


class TestSlowPowerPoint:
    """案内を出したあとの振る舞い（macOS 経路）．"""

    def _patch(self, monkeypatch, proc):
        opened: list[list[str]] = []
        monkeypatch.setattr(pdf.subprocess, "Popen", lambda *a, **k: proc)
        monkeypatch.setattr(pdf.subprocess, "run",
                            lambda cmd, **k: opened.append(cmd) or SimpleNamespace(
                                returncode=0, stdout="", stderr=""))
        return opened

    def test_a_terminal_gets_the_notice_and_powerpoint_up_front(self, tmp_path,
                                                                monkeypatch, stderr_is):
        written = stderr_is(True)
        proc = FakeProc()
        opened = self._patch(monkeypatch, proc)

        dst = tmp_path / "out.pdf"
        dst.write_bytes(b"%PDF")     # 成功したことにする（存在と非空だけ見る）
        pdf._macos_run_applescript(str(tmp_path / "in.pptx"), str(dst), None)

        assert "taking longer" in "".join(written)
        assert any("Microsoft PowerPoint" in " ".join(cmd) for cmd in opened), \
            "tty では前面に出して応答してもらう"
        assert proc.timeouts[-1] is None, "tty なら 2 回目は無制限に待つ"
        assert not proc.killed

    def test_without_a_terminal_nothing_is_brought_to_the_front(self, tmp_path,
                                                               monkeypatch, stderr_is):
        """非 tty での前面化は，押せないダイアログのために作業画面を奪うだけ．"""
        written = stderr_is(False)
        proc = FakeProc()
        opened = self._patch(monkeypatch, proc)

        dst = tmp_path / "out.pdf"
        dst.write_bytes(b"%PDF")
        pdf._macos_run_applescript(str(tmp_path / "in.pptx"), str(dst), None)

        assert "taking longer" in "".join(written)
        assert opened == [], "前面化してはいけない"

    def test_giving_up_kills_only_our_child(self, tmp_path, monkeypatch, stderr_is):
        stderr_is(False)
        proc = FakeProc(calls_before_return=2)   # 2 回とも時間切れ
        self._patch(monkeypatch, proc)

        with pytest.raises(pdf.PdfError, match="timed out after 60s"):
            pdf._macos_run_applescript(str(tmp_path / "in.pptx"),
                                       str(tmp_path / "out.pdf"), 60.0)
        assert proc.killed, "打ち切るなら起こした子は回収する"

    def test_a_limit_below_the_notice_gives_up_without_it(self, tmp_path, monkeypatch,
                                                          stderr_is):
        """上限が案内より短ければ，案内を待たずにそこで打ち切る．"""
        written = stderr_is(False)
        proc = FakeProc(calls_before_return=2)
        self._patch(monkeypatch, proc)

        with pytest.raises(pdf.PdfError, match="timed out after 5s"):
            pdf._macos_run_applescript(str(tmp_path / "in.pptx"),
                                       str(tmp_path / "out.pdf"), 5.0)
        # 1 回目が上限そのもの（案内の 30 秒を待たない）．続く 10.0 は kill 後の回収．
        assert proc.timeouts[0] == 5.0
        assert "taking longer" not in "".join(written)
