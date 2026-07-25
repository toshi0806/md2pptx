#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""入力の変更を見張って作り直し続ける（``--watch``．DESIGN.md §7 / Issue #39）．

「Markdown を編集しながら PDF を見る」運用の最後の一片．保存するたびに手で
コマンドを叩かなくて済むようにするだけで，**速くはならない**——1 回のビルドは
pptx 側が実測 0.2 秒弱（うち 0.13 秒はインタプリタ起動）で，支配項は PDF 変換の
1〜数秒だから，常駐しても縮むのは 1 割に満たない．

**監視は自前のポーリング**で行い，watchdog のような依存は足さない．見張るのは
「そのビルドが実際に読んだファイル」——Markdown・テーマ・画像の数個で，ディレクトリ
全体ではない．0.25 秒ごとに数個の ``os.stat`` を呼ぶだけなので実質ゼロコストであり，
必要なのは「変わったか」だけでイベントの種別も要らない．``sleep`` を差し替えられる
ようにしてあるので，テストは実時間 0 で決定的に回る．ディレクトリ全体を見張りたく
なったら，そのときに再考する．

**止まらないこと**が設計の中心．文法エラーで落ちてしまうと，直して保存しても誰も
作り直してくれない——編集中は失敗しているのが普通の状態なので，``build`` は失敗を
表示だけして次の変更を待つ．

このモジュールは cli にも python-pptx にも依存しない（``build`` は呼び出し側が
渡す）．監視と，いつ作り直すかの判断だけを担う．
"""
from __future__ import annotations

import os
import signal
import sys
import time
from typing import Any, Callable, Iterable

# ポーリング間隔（秒）．監視対象は既知のファイル数個なので，1 周あたり数十マイクロ秒．
# 保存から検知までは最悪 2 周（変更の検出＋落ち着いたことの確認）＝ 0.5 秒．
POLL_INTERVAL = 0.25

# 「落ち着くまで待つ」の上限（秒）．書き込みが延々と続く相手——別のプロセスが書き換え
# 続けるファイル，非常に遅い回線越しのコピー——に当たると，落ち着く瞬間が来ないまま
# **一度も作り直さずに黙り込む**．これは失敗の仕方として最悪で，利用者からは「watch が
# 壊れた」としか見えない．そこで諦めて作りに行く．中途半端なファイルを読んで失敗しても
# watch は止まらず，落ち着いた後の変更でまた作り直せるので，黙るよりはるかにましである．
SETTLE_LIMIT = 10.0

# 変更検知に使う指紋の型．mtime だけでは足りない——ファイルシステムによっては解像度が
# 1 秒しかなく，同じ秒に同じ大きさで書き直されると見落とす．1 回の os.stat から取れる
# 値を組み合わせて，取りこぼしを実用上無くす．None は「そこに無い」．
Signature = tuple[int, int, int, int] | None


def _signature(path: str) -> Signature:
    """変更検知に使う指紋．取れなければ None（不在も 1 つの状態として扱う）．

    不在を None として持つので，「消された」「置かれた」も変更として検出できる
    ——足りない画像を置いたら作り直したい，という編集中の動きがそのまま拾える．
    """
    try:
        st = os.stat(path)
    except OSError:
        return None
    return (st.st_mtime_ns, st.st_ctime_ns, st.st_size, st.st_ino)


def _snapshot(paths: Iterable[str]) -> dict[str, Signature]:
    """与えられたパス集合の現在の指紋を取る．"""
    return {path: _signature(path) for path in paths}


def _first_change(before: dict[str, Signature],
                  after: dict[str, Signature]) -> str | None:
    """変わった最初のパス．どれも変わっていなければ None．

    比べるのは ``before`` の鍵だけ．``after`` は同じ集合を ``_snapshot`` し直したもの
    という前提で，監視対象の増減はここではなく ``run`` がビルドの戻り値で行う．
    値の ``None``（不在）も指紋の一種なので，消えた・置かれたも差として出る．
    """
    for path, sig in before.items():
        if after.get(path) != sig:
            return path
    return None


def _wait_for_change(watched: dict[str, Signature], interval: float,
                     sleep: Callable[[float], None],
                     log: Callable[[str], None],
                     settle_limit: float = SETTLE_LIMIT) -> str:
    """どれかが変わるまで待ち，**落ち着いてから**最初に変わったパスを返す．

    変化を見つけても即座には返さず，1 周期分静かになるまで待つ．エディタの保存は
    書き込みを複数回に分けることがあり，大きい画像のコピーは途中経過が見える．
    そのまま作り直すと，中途半端なファイルを読んで失敗したあと，もう一度作り直す
    ことになる．

    ただし待つのは ``settle_limit`` 秒まで．書き込みが止まらない相手だと落ち着く瞬間が
    来ず，**黙り込んだまま一度も作り直さない**——それは watch が壊れたようにしか
    見えないので，上限に達したら理由を出して作りに行く．
    """
    while True:
        sleep(interval)
        current = _snapshot(watched)
        changed = _first_change(watched, current)
        if changed is None:
            continue
        waited = 0.0
        while True:
            sleep(interval)
            waited += interval
            settled = _snapshot(watched)
            if settled == current:
                return changed
            current = settled
            if waited >= settle_limit:
                log(_stamp(f"{os.path.basename(changed)} keeps changing after "
                           f"{settle_limit:.0f}s — rebuilding anyway"))
                return changed


def _stamp(message: str) -> str:
    """進捗行．時刻を入れるのは，同じ行が何度も流れるログで前回との区別が要るため．"""
    return f"md2pptx: {time.strftime('%H:%M:%S')} {message}"


def _to_stderr(message: str) -> None:
    """進捗は stderr へ（stdout の ``saved:`` 行は機械可読なまま汚さない）．"""
    sys.stderr.write(message + "\n")
    sys.stderr.flush()


def _raise_interrupt(signum: int, frame: Any) -> None:
    raise KeyboardInterrupt


# 「SIGTERM を差し替えていない」ことを表す印．signal.signal は「Python 側で設定
# していないハンドラ」に対して None を返すので，None を「未設定」に使えない．
_NOT_INSTALLED: Any = object()


def _install_sigterm() -> Any:
    """SIGTERM を ``Ctrl-C`` と同じ経路へ寄せる（戻り値は復帰用）．

    既定の SIGTERM は即死で ``finally`` が走らない．VS Code の「タスクの終了」や
    ``kill`` で止めたときに，変換中の作業ディレクトリが残ってしまう．**watch の
    ときだけ**入れる——一発実行の挙動は変えない．
    """
    if not hasattr(signal, "SIGTERM"):
        return _NOT_INSTALLED           # Windows の一部環境
    try:
        return signal.signal(signal.SIGTERM, _raise_interrupt)
    except (ValueError, OSError):
        # メインスレッド以外からは設定できない（ライブラリとして呼ばれた場合）．
        return _NOT_INSTALLED


def _restore_sigterm(previous: Any) -> None:
    """``_install_sigterm`` で置き換えたハンドラを戻す．"""
    if previous is _NOT_INSTALLED:
        return
    try:
        signal.signal(signal.SIGTERM, previous)
    except (ValueError, OSError):
        pass


def run(build: Callable[[], Iterable[str]], label: str, *,
        seed: Iterable[str] = (),
        interval: float = POLL_INTERVAL,
        settle_limit: float = SETTLE_LIMIT,
        sleep: Callable[[float], None] = time.sleep,
        log: Callable[[str], None] = _to_stderr) -> int:
    """変更のたびに ``build()`` を呼び続ける．戻り値は終了コード（常に 0）．

    Args:
        build: 1 回ビルドし，**次に監視すべきファイル集合**を返す呼び出し可能オブジェクト．
            失敗しても例外を投げず（呼び出し側が表示して）集合を返すこと——投げると
            watch が止まり，直して保存しても誰も作り直さなくなる．
        label: 進捗表示に使う入力の名前．
        seed: **ビルドする前から分かっている**監視対象（入力の Markdown）．初回ビルドの
            最中に来た保存を取りこぼさないために要る——依存が判明するのは ``build`` が
            返った後なので，これが無いと比較の起点になる「ビルド前の指紋」を持てず，
            初回ビルド中の保存が「反映済み」に見えてしまう．実 PowerPoint での初回は
            数秒かかる（実測 6 秒）ので，起動してすぐ書き始めると普通に踏む．
        interval: ポーリング間隔（秒）．
        settle_limit: 「落ち着くまで待つ」の上限（秒）．超えたら書き込みが続いていても
            作りに行く（``SETTLE_LIMIT``）．
        sleep: 待ち方（テストが差し替える）．
        log: 進捗の出力先（テストが差し替える）．**渡される文字列は完成した 1 行**で，
            ``md2pptx: `` の接頭辞も含む．log 側で飾り付けはしない（両方で付けると
            二重になる）．

    Returns:
        終了コード．``Ctrl-C`` / SIGTERM での停止は**意図した停止**なので 0．

    Note:
        ``seed`` に無い依存（テーマ・画像）を**初回ビルド中に**書き換えた場合だけは
        取りこぼす．そのファイルが依存だと分かるのがビルド後だからで，起動直後の
        数秒に限られる．
    """
    log(f"md2pptx: watching {label} — Ctrl-C to stop")
    watched: dict[str, Signature] = _snapshot(seed)
    changed: str | None = None
    previous = _install_sigterm()
    try:
        while True:
            trigger = f" ({os.path.basename(changed)} changed)" if changed else ""
            log(_stamp(f"rebuilding {label}{trigger}"))
            # 指紋はビルド**前**の値を使う．ビルド中（PDF 変換で数秒）に保存された
            # ぶんを取りこぼさないため——後から取ると，その保存を「もう反映済み」と
            # 誤認して次の変更まで止まってしまう．
            before = _snapshot(watched)
            sources = build()
            watched = {path: before[path] if path in before else _signature(path)
                       for path in sorted(sources)}
            log(_stamp("watching for changes"))
            changed = _wait_for_change(watched, interval, sleep, log, settle_limit)
    except KeyboardInterrupt:
        log("md2pptx: stopped watching")
        return 0
    finally:
        _restore_sigterm(previous)
