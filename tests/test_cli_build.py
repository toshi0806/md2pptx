#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""``build_once`` への分解で一発実行が変わっていないことを固定するテスト（Issue #39）．

``--watch`` のために ``_run`` を ``build_once`` へ切り出した．watch から呼ぶには失敗を
``SystemExit`` ではなく ``BuildError`` で返す必要があるが，**一発実行のメッセージと
終了コードは 1 バイトも変えない**のが条件だった（唯一の例外は下の
``unsupported theme format``）．型検査も example.md の生成も通ってしまうので，
ここで固定しないとメッセージの取り違えに気づけない．

もう 1 つ固定するのは **watch が見張る対象**．自分の出力を見張ると「作る → 変わった →
また作る」の無限ループになり、依存を取りこぼすと保存しても何も起きない．

``render.build`` は差し替えるので pptx は実際には作らない（テーマも要らない）．
"""
from __future__ import annotations

import pytest

from md2pptx import cli
from md2pptx.ir import Image


@pytest.fixture
def project(tmp_path, monkeypatch):
    """最小の原稿一式（Markdown・テーマ・画像）と，描画を差し替えた環境．"""
    theme = tmp_path / "theme.pptx"
    theme.write_bytes(b"not really a pptx")
    picture = tmp_path / "fig.png"
    picture.write_bytes(b"not really a png")
    md = tmp_path / "slide.md"
    md.write_text(
        "---\ntitle: t\n---\n\n## one\n\n- a\n\n"
        f"## two\n\n![cap]({picture.name})\n"
    )

    written: list[str] = []

    def fake_build(deck, base_pptx_path, out_path, base_dir=None):
        written.append(out_path)

    monkeypatch.setattr(cli.render, "build", fake_build)
    return type("Project", (), {
        "md": md, "theme": theme, "picture": picture,
        "out": tmp_path / "out.pptx", "written": written,
    })


def _main(argv):
    """``main`` を呼び，(終了コード, SystemExit のメッセージ) を返す．"""
    try:
        return cli.main(argv), None
    except SystemExit as e:
        code = e.code
        if isinstance(code, str):
            return 1, code
        return code, None


class TestOneShotIsUnchanged:
    """一発実行の失敗メッセージと終了コード（分解前と同一であること）．"""

    def test_a_successful_run_returns_zero(self, project, capsys):
        code, message = _main([str(project.md), "--theme", str(project.theme),
                               "-o", str(project.out)])

        assert (code, message) == (0, None)
        # 3 = タイトルスライド（front matter の title 由来）＋ 見出し 2 枚．
        assert capsys.readouterr().out == f"saved: {project.out} slides: 3\n"

    def test_a_missing_input_is_reported(self, tmp_path):
        code, message = _main([str(tmp_path / "nope.md"), "--theme", "t.pptx",
                               "-o", "o.pptx"])

        assert code == 1
        assert message == f"md2pptx: input not found: {tmp_path / 'nope.md'}"

    def test_an_empty_pdf_output_is_reported(self, project):
        code, message = _main([str(project.md), "--theme", str(project.theme),
                               "-o", str(project.out), "--pdf-output", " "])

        assert (code, message) == (1, "md2pptx: --pdf-output requires a path")

    def test_a_parse_failure_is_reported(self, tmp_path, project):
        broken = tmp_path / "broken.md"
        broken.write_text("---\ntitle: [unclosed\n---\n\n## x\n")

        code, message = _main([str(broken), "--theme", str(project.theme),
                               "-o", str(project.out)])

        assert code == 1
        assert message.startswith(f"md2pptx: failed to parse {broken}: ")

    def test_a_missing_theme_is_reported(self, project):
        code, message = _main([str(project.md), "-o", str(project.out)])

        assert (code, message) == (
            1, "md2pptx: no theme specified (use --theme or front matter 'theme')")

    def test_a_missing_output_is_reported(self, project):
        code, message = _main([str(project.md), "--theme", str(project.theme)])

        assert (code, message) == (
            1, "md2pptx: no output specified (use -o or front matter 'output')")

    def test_a_render_failure_is_reported(self, project, monkeypatch):
        def explode(deck, base_pptx_path, out_path, base_dir=None):
            raise ValueError("boom")

        monkeypatch.setattr(cli.render, "build", explode)

        code, message = _main([str(project.md), "--theme", str(project.theme),
                               "-o", str(project.out)])

        assert (code, message) == (
            1, f"md2pptx: failed to render {project.md}: boom")

    def test_front_matter_of_the_wrong_type_is_reported(self, tmp_path, project):
        typo = tmp_path / "typo.md"
        typo.write_text("---\ntheme: 123\n---\n\n## x\n")

        code, message = _main([str(typo), "-o", str(project.out)])

        assert (code, message) == (
            1, "md2pptx: front matter 'theme' must be a string, got int (123)")

    def test_an_unsupported_theme_gains_the_common_prefix(self, tmp_path, project):
        """**唯一変わったメッセージ**：他と同じ ``md2pptx: `` 接頭辞が付く．

        分解前はここだけ接頭辞が無かった（``SystemExit`` に直接書いていたため）．
        """
        theme = tmp_path / "theme.key"
        theme.write_bytes(b"")

        code, message = _main([str(project.md), "--theme", str(theme),
                               "-o", str(project.out)])

        assert (code, message) == (
            1, "md2pptx: unsupported theme format: .key (expected .thmx or .pptx)")


class TestWatchedSources:
    """``BuildResult.sources``：watch が次に見張るファイル．"""

    def _build(self, project, *extra):
        return cli.build_once(cli._parse_args(
            [str(project.md), "--theme", str(project.theme),
             "-o", str(project.out), *extra]))

    def test_the_markdown_theme_and_images_are_all_sources(self, project):
        """どれが変わっても作り直したい——テーマ駆動がこのツールの本体なので
        テーマも，原稿に貼った画像も外せない．"""
        result = self._build(project)

        assert str(project.md) in result.sources
        assert str(project.theme) in result.sources
        assert str(project.picture) in result.sources

    def test_the_output_is_never_a_source(self, project):
        """自分の出力を見張ると「作る → 変わった → また作る」の無限ループになる．"""
        result = self._build(project)

        assert str(project.out) not in result.sources
        assert result.output == str(project.out)

    def test_images_in_columns_are_found_too(self, tmp_path, project):
        """2 カラム目の画像も見張る（走査が ``blocks`` だけだと落ちる）．"""
        picture = tmp_path / "col.png"
        picture.write_bytes(b"x")
        project.md.write_text(
            "---\ntitle: t\n---\n\n## two columns\n\n"
            f"- left\n\n<!-- @col -->\n\n![a]({picture.name})\n"
        )
        # 本当に多カラムになり，画像が blocks からは見えないことを先に固定する
        # ——記法を書き間違えると，走査が columns を見ていなくても通ってしまう．
        slide = cli.md_parser.parse_file(str(project.md)).slides[0]
        assert len(slide.columns) == 2
        assert not any(isinstance(b, Image) for b in slide.blocks)
        assert any(isinstance(b, Image) for b in slide.columns[1])

        assert str(picture) in self._build(project).sources

    def test_a_failure_still_reports_what_it_learned(self, project, monkeypatch):
        """失敗しても分かった依存は返す——画像を置いたら作り直したいから．"""
        def explode(deck, base_pptx_path, out_path, base_dir=None):
            raise ValueError("boom")

        monkeypatch.setattr(cli.render, "build", explode)

        with pytest.raises(cli.BuildError) as found:
            self._build(project)

        assert str(project.md) in found.value.sources
        assert str(project.picture) in found.value.sources

    def test_an_early_failure_still_reports_the_input(self, tmp_path, project):
        """テーマも解決できていない段階の失敗でも，入力 Markdown だけは返る．"""
        with pytest.raises(cli.BuildError) as found:
            cli.build_once(cli._parse_args([str(project.md), "-o", str(project.out)]))

        assert found.value.sources == frozenset({str(project.md)})


class TestWatchLoop:
    """``--watch`` のループが満たすべき性質（止まらないこと）．"""

    def test_a_failing_build_does_not_stop_the_watch(self, project, monkeypatch,
                                                     capsys):
        """文法エラーで落ちたら，直して保存しても誰も作り直さなくなる．

        編集中は失敗しているのが普通の状態なので，表示だけして待ち続ける．
        """
        calls: list[int] = []

        def explode(deck, base_pptx_path, out_path, base_dir=None):
            calls.append(1)
            raise ValueError("boom")

        monkeypatch.setattr(cli.render, "build", explode)
        # 2 回ビルドさせたところで止める（watch.run 本体は test_watch.py が見る）．
        def fake_run(build, label, **kwargs):
            build()
            build()
            return 0

        monkeypatch.setattr(cli.watch, "run", fake_run)

        code, message = _main([str(project.md), "--theme", str(project.theme),
                               "-o", str(project.out), "--watch"])

        assert (code, message) == (0, None), "失敗しても watch は 0 で終わる"
        assert len(calls) == 2, "1 回目の失敗で止まってはいけない"
        assert capsys.readouterr().err.count("failed to render") == 2

    def test_the_watch_tells_the_converter_that_nobody_is_looking(self, project,
                                                                  monkeypatch):
        """人が見ているのはエディタと PDF であって，このタスク端末ではない．"""
        seen: dict[str, object] = {}

        def fake_convert(src, dst, converter, timeout=None, *, unattended=False):
            seen["unattended"] = unattended

        monkeypatch.setattr(cli.pdf_backend, "convert", fake_convert)
        monkeypatch.setattr(cli.watch, "run", lambda build, label, **kw: (build(), 0)[1])

        _main([str(project.md), "--theme", str(project.theme),
               "-o", str(project.out), "--watch", "--pdf"])

        assert seen["unattended"] is True

    def test_a_one_shot_run_leaves_the_converter_alone(self, project, monkeypatch):
        """一発実行の待ち方は変えない（tty なら従来どおり待ち続ける）．"""
        seen: dict[str, object] = {}

        def fake_convert(src, dst, converter, timeout=None, *, unattended=False):
            seen["unattended"] = unattended

        monkeypatch.setattr(cli.pdf_backend, "convert", fake_convert)

        _main([str(project.md), "--theme", str(project.theme),
               "-o", str(project.out), "--pdf"])

        assert seen["unattended"] is False
