#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""``--watch`` の見張り方を固定するテスト（Issue #39）．

守りたいのは 4 つ．

- **初回は保存を待たない**（起動したのに何も起きないと壊れて見える）
- **1 回の保存で 1 回だけ作り直す**（エディタは書き込みを複数回に分けることがある）
- **書き込みの途中では作り直さない**（中途半端なファイルを読んで失敗する）
- **ビルド中に来た保存を取りこぼさない**（PDF 変換の数秒はまるごと死角になりうる）

``sleep`` を差し替えて「待つたびに次の手順を 1 つ実行する」台本を渡すので，実時間 0 で
決定的に回る．外部プロセスも実際の待ちも発生しない．
"""
from __future__ import annotations

import os

import pytest

from md2pptx import watch


class Script:
    """``sleep`` の代わり．呼ばれるたびに次の手順を 1 つ実行する．

    手順を使い切ったら ``KeyboardInterrupt`` を投げて watch を終わらせる
    （``Ctrl-C`` と同じ経路なので，後始末まで含めて本番と同じ流れになる）．
    """

    def __init__(self, *steps):
        self.steps = list(steps)
        self.slept = 0

    def __call__(self, seconds):
        self.slept += 1
        if not self.steps:
            raise KeyboardInterrupt
        step = self.steps.pop(0)
        if step is not None:
            step()


class Build:
    """``build`` の代わり．呼ばれた回数と，そのとき見えていた内容を覚える．"""

    def __init__(self, *source_sets, watching=()):
        # 呼び出しごとに返す集合（尽きたら最後のものを返し続ける）．
        self.source_sets = [list(s) for s in source_sets] or [list(watching)]
        self.calls = 0
        self.seen: list[bytes | None] = []
        self.during = None      # 呼ばれている間に実行する副作用

    def __call__(self):
        sources = self.source_sets[min(self.calls, len(self.source_sets) - 1)]
        self.calls += 1
        first = sources[0] if sources else None
        self.seen.append(first.read_bytes() if first and first.exists() else None)
        if self.during is not None:
            self.during()
        return [str(p) for p in sources]


@pytest.fixture
def deck(tmp_path):
    """入力の Markdown に見立てたファイル（中身は読まれない）．"""
    path = tmp_path / "slide.md"
    path.write_text("one")
    return path


def run(build, script, label="slide.md"):
    """既定の引数をまとめた ``watch.run``（ログは捨てる）．"""
    return watch.run(build, label, interval=0.01, sleep=script, log=lambda m: None)


def test_the_first_build_does_not_wait_for_a_change(deck):
    """起動したらすぐ 1 回作る——最初の保存まで何も出ないと壊れて見える．"""
    build = Build([deck])

    assert run(build, Script()) == 0
    assert build.calls == 1


def test_a_quiet_watch_does_not_rebuild(deck):
    """何も変わらなければ作り直さない（保存していないのに走っては困る）．"""
    build = Build([deck])

    run(build, Script(None, None, None))

    assert build.calls == 1


def test_one_save_rebuilds_once(deck):
    """保存 1 回につきビルド 1 回．

    エディタは 1 度の保存で複数回書き込むことがあり，素直に反応すると同じ内容を
    何度も作り直すことになる（PDF 変換が数秒なので体感に響く）．
    """
    build = Build([deck])

    run(build, Script(lambda: deck.write_text("two"), None))

    assert build.calls == 2


def test_a_file_still_being_written_is_left_alone(deck):
    """書き込みが続いている間は待つ．

    大きい画像のコピーは途中経過が見える．そこで作り直すと，壊れたファイルを読んで
    失敗したあと，もう一度作り直すことになる．
    """
    build = Build([deck])

    run(build, Script(
        lambda: deck.write_text("partial"),     # 書き始め
        lambda: deck.write_text("partial-more"),  # まだ伸びている
        None,                                     # ここで落ち着く
    ))

    assert build.calls == 2
    assert build.seen[-1] == b"partial-more", "落ち着いた後の内容で作ること"


def test_a_file_that_never_settles_is_built_anyway(deck):
    """落ち着かない相手でも，いつかは作る——**黙り込むのが最悪**だから．

    別プロセスが書き換え続けるファイルや、非常に遅い回線越しのコピーに当たると
    「静かになった瞬間」が来ない．待ち続けると一度も作り直さないまま無言になり，
    利用者からは watch が壊れたようにしか見えない．
    """
    build = Build([deck])
    notes: list[str] = []
    # 毎周期書き換え続ける（落ち着く瞬間が来ない）．
    forever = [(lambda n=i: deck.write_text(f"v{n}")) for i in range(20)]

    watch.run(build, "slide.md", interval=1.0, settle_limit=5.0,
              sleep=Script(*forever), log=notes.append)

    assert build.calls >= 2, "待ち続けて一度も作らないのは不可"
    assert any("keeps changing" in n for n in notes), "諦めた理由を出すこと"


def test_a_save_during_the_build_is_not_missed(deck):
    """ビルド中に来た保存を取りこぼさない．

    PDF 変換は数秒かかる．その間の保存を「もう反映済み」と誤認すると，次の保存まで
    古い PDF を見続けることになる——指紋をビルド**前**に取っているのはこのため．
    """
    build = Build([deck])
    # 台本は 2 本ある．build.during は**ビルドの回数**で進み，Script は**待った回数**で
    # 進む．噛み合わせは次のとおり:
    #   ビルド1（保存なし）→ 待1: 保存 → 待2: 落ち着く → ビルド2（最中に保存）
    #   → 待3: その保存を検出 → 待4: 落ち着く → ビルド3 → 待5: 台本切れで停止
    # 初回ビルド中に保存しないのは，seed 無しの run ではそこが死角だから
    # （seed 版は test_a_save_during_the_first_build_is_not_missed が見る）．
    saves = iter([lambda: None,                              # ビルド1
                  lambda: deck.write_text("during-2"),       # ビルド2
                  lambda: None])                             # ビルド3
    build.during = lambda: next(saves)()

    run(build, Script(lambda: deck.write_text("saved"), None, None, None))

    assert build.calls == 3, "ビルド中の保存も次のビルドに反映されること"


def test_a_save_during_the_first_build_is_not_missed(deck):
    """**初回**ビルド中の保存も拾う（``seed`` が要る理由）．

    依存が判明するのはビルドが返った後なので，何も渡さないと初回だけは比較の起点を
    持てず，その間の保存が「反映済み」に見えてしまう．実 PowerPoint での初回は実測
    6 秒かかるので，起動してすぐ書き始めると普通に踏む．
    """
    build = Build([deck])
    saves = iter([lambda: deck.write_text("during-first"),
                  lambda: None, lambda: None])
    build.during = lambda: next(saves)()

    watch.run(build, "slide.md", seed=[str(deck)], interval=0.01,
              sleep=Script(None, None), log=lambda m: None)

    assert build.calls == 2, "初回ビルド中の保存も次のビルドに反映されること"


def test_the_watched_set_follows_the_build(tmp_path):
    """見張る対象は毎回 build の戻り値で入れ替える．

    原稿から画像を消したらもう見張らない——消えた依存を追い続けると，無関係な
    ファイルの変更で作り直してしまう．
    """
    gone = tmp_path / "dropped.png"
    kept = tmp_path / "slide.md"
    gone.write_text("a")
    kept.write_text("b")
    build = Build([gone], [kept])       # 2 回目以降は kept だけを見張る

    run(build, Script(
        lambda: gone.write_text("changed"),   # 1 回目：見張っているので反応する
        None,
        lambda: gone.write_text("again"),     # 2 回目：もう見張っていない
        None,
    ))

    assert build.calls == 2


def test_a_new_dependency_is_picked_up(tmp_path):
    """ビルドで初めて分かった依存も次から見張る（画像を足した直後など）．"""
    md = tmp_path / "slide.md"
    picture = tmp_path / "fig.png"
    md.write_text("a")
    picture.write_text("b")
    build = Build([md], [md, picture])

    run(build, Script(
        lambda: md.write_text("uses a picture now"),
        None,
        lambda: picture.write_text("edited"),   # 新しい依存の変更で作り直す
        None,
    ))

    assert build.calls == 3


def test_a_missing_file_is_still_watched(tmp_path):
    """まだ無いファイルも見張る．

    「画像が見つからない」で失敗したときこそ，その画像が置かれたら作り直したい．
    不在も 1 つの状態として指紋に持つのでそのまま扱える．
    """
    missing = tmp_path / "not-yet.png"
    build = Build([missing])

    run(build, Script(lambda: missing.write_text("arrived"), None))

    assert build.calls == 2


def test_stopping_is_not_an_error(deck):
    """``Ctrl-C`` / SIGTERM は意図した停止なので終了コード 0．"""
    assert run(Build([deck]), Script()) == 0


class TestSignature:
    """指紋の取り方（``os.stat`` 1 回で「変わったか」を判定する）．"""

    def test_absence_is_a_state(self, tmp_path):
        assert watch._signature(str(tmp_path / "nope")) is None

    def test_same_content_same_signature(self, deck):
        assert watch._signature(str(deck)) == watch._signature(str(deck))

    def test_a_rewrite_the_clock_did_not_notice_is_still_caught(self, deck):
        """**mtime と長さが揃っていても**書き換えに気づく．

        解像度が 1 秒の FS では，同じ秒に同じ長さで書き直すと mtime も size も
        変わらない．ここではその状況を ``os.utime`` で作る——ただ書き直すだけでは，
        高解像度の FS だと mtime_ns が動いてしまい，指紋が mtime しか見ていなくても
        通ってしまう（＝何も確かめていないテストになる）．
        """
        before = os.stat(deck)
        deck.write_text("ONE")                      # 長さは同じ（3 文字）
        os.utime(deck, ns=(before.st_atime_ns, before.st_mtime_ns))
        after = os.stat(deck)
        # 前提：この状況では mtime・size・inode のどれも変わっていない．
        assert (after.st_mtime_ns, after.st_size, after.st_ino) == (
            before.st_mtime_ns, before.st_size, before.st_ino)

        assert watch._signature(str(deck)) != (
            before.st_mtime_ns, before.st_ctime_ns, before.st_size, before.st_ino)
