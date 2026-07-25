#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""pptx → PDF 変換（生成後プレビュー用の土台．DESIGN.md §7 / Issue #39）．

「Markdown を編集しながら PDF を見る」運用の基礎として，生成した pptx を
そのまま PDF にする．**忠実度は変換器による**：LibreOffice の出力はテーマ
フォントの解決差などで実 PowerPoint と一致しない（当たり確認どまり）が，
PowerPoint 経路は実 PowerPoint 自身の出力なので見た目の確認に使える（README 参照）．

変換器は 3 系統:

- ``auto``（既定）: native PowerPoint → LibreOffice の順に，**使えるもの**を選ぶ．
  選んだ変換器が失敗したら次へは落とさずそのまま失敗する（Issue #46）．
- ``powerpoint`` / ``libreoffice``: その系統を名指し．
- 任意のコマンド行: ``mytool -o {output} {input}`` のように直接指定．
  プレースホルダ ``{input}`` / ``{output}`` / ``{outdir}`` を置換する．1 つも
  無ければ末尾に ``{input}`` を補う（出力パスを取らないツール向け）．その場合
  ツールは入力の隣に ``<basename>.pdf`` を書く想定で，期待パスと違えば移動する．

**macOS の native PowerPoint 対応**：AppleScript 辞書に ``export`` コマンドは無いが，
``save … in (POSIX file p) as save as PDF`` は POSIX file への coerce により安定して動作
する（PowerPoint 16.111.1 で 14 ページの変換を実測）。「無反応でハング」して見えるときは
スクリプトの誤りではなくダイアログの応答待ちを疑うこと．``auto`` は macOS で PowerPoint.app
があれば実 PowerPoint を優先し，無い／失敗した場合は LibreOffice へフォールバックする．
Windows の PowerPoint は COM（``SaveAs`` format 32）で対応．

**PowerPoint を目立たせずに使う（macOS）**．2 つ組み合わせる（Issue #44）：

- ``activate`` を入れず ``open -g -j -a`` で非表示・非アクティブ起動する
  （``_macos_prelaunch_powerpoint_hidden``）．未起動からの変換なら全工程でウィンドウが
  出ない．既に表示して使っているインスタンスには効かない（``-j`` は起動の瞬間だけ）．
  ``save … as PDF`` は隠したアプリを自ら再表示するので，起動後に隠し直す方法は使えない．
- pptx をコンテナへコピーして**その中だけを触らせる**（``_macos_container_tmp``）．
  未承認の場所を直接開かせるとファイルアクセスの許可ダイアログ待ちで止まるが，隠して
  動かしている以上それは利用者から見えないので，そもそも出させない．入出力がどこに
  あっても動くという副次効果もある．

隠したことで気づけない停止（オートメーション承認など）に備え，``_MACOS_HINT_AFTER`` 秒で
案内を stderr に出す（前面化は tty のときだけ）．

**待ちの上限（Issue #48）**：止まり方には「人が今すぐ直せるもの」（承認ダイアログ・サインイン
画面）と「誰も直さないもの」（GUI セッションの無い cron / CI・クラッシュ）がある．時計では
区別できないので **stderr が tty か**で分ける——tty なら打ち切らず（案内が効くので待つ意味が
ある．``Ctrl-C`` で止められる），非 tty なら ``_TIMEOUT_UNATTENDED`` 秒で打ち切る．
``--pdf-timeout`` / ``MD2PPTX_PDF_TIMEOUT`` で上書きでき，``0`` は無制限．
``convert(..., unattended=True)`` は「端末は tty だが人は見ていない」を明示する入口で，
``--watch`` が使う（上限の決め方と，止まったときの前面化の両方に効く）．

**出力はアトミックに差し替える**．変換は出力先と同じディレクトリに作った使い捨ての作業
ディレクトリの中で行い，成功したときだけ ``os.replace`` で目的のパスへ移す．編集しながら
見る運用が前提なので，**変換中に出力 PDF が一瞬でも消えてはいけない**——PDF ビューアは
フォルダを監視していて，削除を確定するとそのファイルを監視から外してしまう（``convert``
の実装コメント参照）．

このモジュールは cli 以外に依存しない（python-pptx 非依存）．外部プロセスの
起動と，どのバイナリを使うかの解決だけを担う．
"""
from __future__ import annotations

import math
import os
import pathlib
import shlex
import shutil
import subprocess
import sys
import time

from . import workdir


class PdfError(Exception):
    """PDF 変換の失敗（原因メッセージ付き）．cli が警告表示に使う．"""


class _Unavailable(PdfError):
    """その変換器がこの環境に**無い**（＝失敗ではない）．

    ``auto`` はこれだけを握って次の変換器へ進む．無い物を飛ばしても何も失われない
    のに対し，**在る物の失敗**を飛ばすと忠実度の違う PDF を黙って掴ませることに
    なる（Issue #46）．名指し指定のときは PdfError としてそのまま利用者に届く．
    """


# 環境変数名（CLI 引数 --pdf-converter / --pdf-timeout が優先）．
ENV_CONVERTER = "MD2PPTX_PDF_CONVERTER"
ENV_TIMEOUT = "MD2PPTX_PDF_TIMEOUT"

# 人が見ていないときに変換を待つ上限（秒）．30 秒の案内を見てから応答するまでの猶予として
# 置いている——支配項は変換そのものの所要時間（実測で数秒〜1 分）ではなく人の応答時間．
_TIMEOUT_UNATTENDED = 180.0

# 補助コマンド（LaunchServices への問い合わせ）の上限．実測 50ms 前後なので，これだけ
# 待って返らなければ異常．人の応答を待つ場面ではないので tty かどうかで分けない．
_HELPER_TIMEOUT = 10.0

# Windows の COM 経路で「PowerPoint はあった」ことを示す目印．PowerShell に
# COM オブジェクト生成の直後で出力させ，これが出る前に落ちたか後で落ちたかで
# 「無い（_Unavailable）」と「失敗（PdfError）」を切り分ける．
_WIN_COM_READY = "MD2PPTX_POWERPOINT_READY"

# macOS の PowerPoint 変換がこれだけ待っても終わらなければ，ダイアログ待ちを疑って
# 案内を出す（変換は続ける）．コンテナ経由の変換は例で 1〜8 秒なので誤検知しない幅．
_MACOS_HINT_AFTER = 30.0


# macOS の実 PowerPoint で pptx → PDF にする AppleScript．osascript に stdin で渡し，
# 入出力パスは argv で渡す（パスを文字列リテラルに埋め込まないので，スペースや引用符を
# 含むパスでも構文が壊れない）．``POSIX file`` への coerce は Sonoma 以降の alias 問題の
# 回避に必須（素の POSIX パス文字列では保存先を解決できない）．
# ``activate`` は入れない——変換のたびに PowerPoint が前面に出て作業画面を奪うため
# （変換自体は activate 無しで成立する）．
_APPLESCRIPT_PPTX_TO_PDF = '''on run argv
    set inPath to item 1 of argv
    set outPath to item 2 of argv
    tell application "Microsoft PowerPoint"
        open (POSIX file inPath)
        set theDoc to active presentation
        save theDoc in (POSIX file outPath) as save as PDF
        close theDoc saving no
    end tell
end run'''


def _resolve_timeout(explicit: float | None,
                     unattended: bool = False) -> float | None:
    """変換の待ち上限（秒）を決める．``None`` は無制限．

    明示指定（``--pdf-timeout`` → ``MD2PPTX_PDF_TIMEOUT``）が最優先で，``0`` は
    無制限の意味．無指定なら **stderr が tty か**で分ける（Issue #48）：

    - **tty**（人が端末を見ている）→ 無制限．止まる原因の多くは承認ダイアログのような
      「人が今すぐ直せるもの」で，30 秒の案内はそれを直してもらうための仕掛けだから，
      上から打ち切ると自分で用意した解決手段を潰すことになる．``Ctrl-C`` で止められる．
    - **非 tty**（cron / CI / エディタ拡張）→ ``_TIMEOUT_UNATTENDED``．誰も応答しないので
      待っても状況は変わらない．

    ``unattended`` はこの tty からの推測を呼び出し側が**明示的に**打ち消す入口．
    ``--watch`` がこれを使う：端末は tty でも，人が見ているのはエディタと PDF であって
    タスクの端末ではない．無制限に待つと以後のプレビューが全部止まってしまうので，
    ``_TIMEOUT_UNATTENDED`` で打ち切って次の保存で作り直す方に賭ける．
    """
    if explicit is not None:
        return _checked(explicit, "--pdf-timeout")
    raw = (os.environ.get(ENV_TIMEOUT) or "").strip()
    if raw:
        try:
            value = float(raw)
        except ValueError:
            raise PdfError(f"invalid {ENV_TIMEOUT}: {raw!r} (seconds, 0 = no limit)")
        return _checked(value, ENV_TIMEOUT)
    if unattended:
        return _TIMEOUT_UNATTENDED
    return None if sys.stderr.isatty() else _TIMEOUT_UNATTENDED


def _checked(value: float, source: str) -> float | None:
    """秒数を検証する．``0`` は無制限，負値・``nan``・``inf`` は誤りとして弾く．

    ``float("nan")`` はどんな比較も偽になるのでそのまま subprocess へ渡ってしまい，
    「上限があるようで無い」不可解な状態になる．負値も同様に弾く——``-5``（``5`` の
    打ち間違い）を無制限と解釈すると，**この機能が防ごうとしている「無人で永久に待つ」
    状態を作ってしまう**．無制限にしたい人には ``0`` という明示の入口がある．
    """
    if math.isnan(value) or math.isinf(value) or value < 0:
        raise PdfError(f"invalid {source}: {value} (seconds, 0 = no limit)")
    return None if value == 0 else value


def _kill(proc: subprocess.Popen[str]) -> None:
    """打ち切りのために**自分で起こした子プロセスだけ**を殺して回収する．

    PowerPoint 本体は殺さない——利用者が開いて使っているインスタンスかもしれないし，
    ダイアログが原因なら画面に残っている方がよい（応答すれば次の実行は通る）．
    """
    proc.kill()
    try:
        proc.communicate(timeout=_HELPER_TIMEOUT)
    except (subprocess.TimeoutExpired, OSError):
        pass


def _timed_out(what: str, limit: float) -> PdfError:
    """打ち切りを PdfError にする．原因の見当と延ばし方を添える．"""
    return PdfError(
        f"{what} timed out after {limit:.0f}s (it may be waiting for a dialog, "
        f"e.g. the automation approval); answer it and run again, or raise the "
        f"limit with --pdf-timeout / {ENV_TIMEOUT}")


def _which_libreoffice() -> str | None:
    """LibreOffice の実行ファイルを探す．PATH 優先，無ければ OS 既知の場所．"""
    for name in ("soffice", "libreoffice"):
        found = shutil.which(name)
        if found:
            return found
    candidates: list[str] = []
    if sys.platform == "darwin":
        candidates.append("/Applications/LibreOffice.app/Contents/MacOS/soffice")
    elif sys.platform.startswith("win"):
        for env in ("ProgramFiles", "ProgramFiles(x86)"):
            base = os.environ.get(env)
            if base:
                candidates.append(
                    os.path.join(base, "LibreOffice", "program", "soffice.exe"))
    for path in candidates:
        if os.path.isfile(path):
            return path
    return None


def _macos_powerpoint_installed() -> bool:
    """macOS に Microsoft PowerPoint が入っているか．

    既定の場所にあれば即座に真（ほとんどはこれで済む）．無ければ LaunchServices に
    **名前で**問い合わせる．変換本体（``open -a`` と ``tell application``）も名前で
    解決するので，置き場所を変えている環境で**ここだけがパスで否定する**と，動くはずの
    PowerPoint を使わずに LibreOffice へ落ちる（＝#46 で消した無言の切り替えが戻る）．
    この問い合わせはアプリを起動しない（実測 46ms）．
    """
    if os.path.isdir("/Applications/Microsoft PowerPoint.app"):
        return True
    try:
        proc = subprocess.run(
            ["osascript", "-e", 'id of app "Microsoft PowerPoint"'],
            capture_output=True, text=True, timeout=_HELPER_TIMEOUT)
    except (OSError, subprocess.TimeoutExpired):
        return False
    return proc.returncode == 0


def _macos_prelaunch_powerpoint_hidden() -> None:
    """PowerPoint を非表示（``-j``）・非アクティブ（``-g``）で先に起動しておく．

    こうしてから AppleScript で文書を開くと，変換の全工程を通してウィンドウが画面に
    出ず，フォアグラウンドも移らない．**``-j`` が効くのは起動の瞬間だけ**なので，
    これで隠せるのは PowerPoint が未起動のときに限る．既に起動していれば何も起きない
    （利用者が表示して使っているインスタンスを勝手に隠すことはない）．

    失敗しても変換は AppleScript 側の暗黙起動で成立するので，ここでは握り潰す
    （PowerPoint の有無は呼び出し側が先に判定している）．
    """
    try:
        subprocess.run(["open", "-g", "-j", "-a", "Microsoft PowerPoint"],
                       capture_output=True, timeout=_HELPER_TIMEOUT)
    except (OSError, subprocess.TimeoutExpired):
        pass


def _macos_container_tmp() -> str | None:
    """PowerPoint のサンドボックスコンテナ内の作業場所（無ければ None）．

    ここに pptx を置いてから開かせると，**ファイルアクセスの承認ダイアログが出ない**．
    コンテナはアプリ自身のサンドボックス領域なので承認の対象外だからで，これにより
    どの場所の入出力でも（``/tmp`` でもネットワークボリュームでも）変換できる．

    ``tmp`` が無いだけなら作る（掃除された後など）．ただし**コンテナ本体（``Data``）が
    無いときは作らない**——コンテナを用意するのは containermanagerd の仕事で，手で
    骨組みだけ置くと正規の初期化を妨げうる．その場合は None を返し，承認ダイアログの
    出うる直接変換へ委ねる．
    """
    base = os.path.expanduser(
        "~/Library/Containers/com.microsoft.Powerpoint/Data/tmp")
    if os.path.isdir(base):
        return base
    if not os.path.isdir(os.path.dirname(base)):
        return None
    try:
        os.makedirs(base, exist_ok=True)
    except OSError:
        return None
    return base


def _macos_run_applescript(src: str, dst: str, timeout: float | None,
                           attended: bool = True) -> None:
    """AppleScript で src → dst を変換する．長引いたら理由を stderr に出す．

    PowerPoint を隠して動かしているので，何かのダイアログ（オートメーションの承認
    など）で止まると，利用者には Dock を見ない限り「無音で固まった」ようにしか見え
    ない．そこで ``_MACOS_HINT_AFTER`` 秒たっても終わらなければ，何が起きている
    可能性があるかを stderr に出す．

    **PowerPoint を前面に出すのは stderr が tty で，かつ ``attended`` のときだけ**
    （Issue #48）．非対話の呼び出し元（cron / エディタ拡張）では，そこで出ているのは
    *呼び出し元アプリ*に対する承認ダイアログなので PowerPoint を前面化しても押せず，
    作業中の画面からフォーカスを奪うだけになる．``--watch`` は端末が tty でも
    ``attended=False`` で呼ぶ——編集中に前面化されるのは邪魔にしかならない．

    その後の待ちは ``timeout``（``None`` で無制限）に従う．
    """
    started = time.monotonic()
    proc = subprocess.Popen(
        ["osascript", "-", src, dst],
        stdin=subprocess.PIPE, stdout=subprocess.PIPE, stderr=subprocess.PIPE,
        text=True)
    # 上限が案内より短ければ，案内を待たずにそこで打ち切る．
    first = _MACOS_HINT_AFTER
    if timeout is not None and timeout <= _MACOS_HINT_AFTER:
        first = timeout
    try:
        _, err = proc.communicate(input=_APPLESCRIPT_PPTX_TO_PDF, timeout=first)
    except subprocess.TimeoutExpired:
        if first != _MACOS_HINT_AFTER:
            _kill(proc)
            raise _timed_out("powerpoint", first)
        sys.stderr.write(
            "md2pptx: PowerPoint is taking longer than "
            f"{_MACOS_HINT_AFTER:.0f}s — it may be waiting for a dialog "
            "(e.g. the automation approval).\n")
        if attended and sys.stderr.isatty():
            sys.stderr.write(
                "md2pptx: bringing it to the front; answer the dialog to continue.\n")
            try:
                subprocess.run(["open", "-a", "Microsoft PowerPoint"],
                               capture_output=True, timeout=_HELPER_TIMEOUT)
            except (OSError, subprocess.TimeoutExpired):
                pass
        if timeout is None:
            _, err = proc.communicate()     # 無制限：応答があるまで待つ
        else:
            # 残りは経過時間から引く．案内と前面化にも時間がかかるので，定数
            # （_MACOS_HINT_AFTER）を引くと利用者の指定した上限を超えてしまう．
            rest = timeout - (time.monotonic() - started)
            try:
                _, err = proc.communicate(timeout=max(rest, 0.0))
            except subprocess.TimeoutExpired:
                _kill(proc)
                raise _timed_out("powerpoint", timeout)
    if proc.returncode != 0:
        detail = (err or "").strip().splitlines()
        tail = detail[-1] if detail else f"exit {proc.returncode}"
        raise PdfError(f"powerpoint failed: {tail}")


def _run(cmd: list[str], what: str, input: str | None = None,
         timeout: float | None = None) -> None:
    """外部コマンドを実行し，失敗を PdfError に変換する．

    成功時の出力は捨てる．失敗時のみ stderr（無ければ stdout）の末尾 1 行を
    原因として拾う（cli が警告に整形する）．input を渡すと stdin に流す
    （osascript にスクリプト本体を与えるのに使う）．``timeout`` を超えたら
    ``subprocess.run`` が子を kill して待ち直す（Python の仕様）ので，こちらは
    PdfError に変えるだけ．
    """
    try:
        proc = subprocess.run(cmd, capture_output=True, text=True, input=input,
                              timeout=timeout)
    except FileNotFoundError:
        raise PdfError(f"{what}: command not found: {cmd[0]}")
    except subprocess.TimeoutExpired as e:
        # 例外が持つ値を使う（timeout=None なら送出されないので None にならない）．
        raise _timed_out(what, e.timeout)
    if proc.returncode != 0:
        detail = (proc.stderr or proc.stdout or "").strip().splitlines()
        tail = detail[-1] if detail else f"exit {proc.returncode}"
        raise PdfError(f"{what} failed: {tail}")


def _convert_libreoffice(src: str, dst: str, timeout: float | None = None) -> None:
    """LibreOffice で src(pptx) → dst(pdf)．--outdir 方式なので後で改名する．"""
    soffice = _which_libreoffice()
    if soffice is None:
        raise _Unavailable(
            "LibreOffice not found (looked for soffice/libreoffice on PATH "
            "and the default install location)")
    outdir = os.path.dirname(os.path.abspath(dst)) or "."
    # 同一プロファイルの多重起動は失敗しうるので，毎回使い捨てのプロファイルを渡す．
    # 片付けは workdir.discard に任せる——TemporaryDirectory は**片付けに失敗すると
    # 例外を投げる**ので，変換が成功していても実行全体が終了コード 1 で終わってしまう
    # （Issue #58）．
    try:
        profile = workdir.create(prefix="md2pptx-lo-")
    except OSError as e:
        # 素の OSError を通すと cli の `except PdfError` をすり抜けて SystemExit に
        # なり，pptx は保存できているのに終了コードが 1 になる（Issue #58．#53 で
        # convert の作業場所について直したのと同じ理由）．
        raise PdfError(f"cannot create a LibreOffice profile directory ({e})") from e
    try:
        # as_uri() は Windows のドライブレターも file:///C:/... と正しく組む
        # （手組みの "file://"+path だと file://C:/... になり不正）．
        uri = pathlib.Path(os.path.abspath(profile)).as_uri()
        _run([
            soffice, "--headless",
            f"-env:UserInstallation={uri}",
            "--convert-to", "pdf", "--outdir", outdir, src,
        ], "libreoffice", timeout=timeout)
    finally:
        workdir.discard(profile)
    # soffice は <入力 basename>.pdf を outdir に書く．期待名と違えば移動する．
    # 使い捨てプロファイルは変換が終わった時点で不要なので，PDF の移動はそれを捨てた
    # 後に行う（プロファイルの寿命と成果物の移動を分離する）．
    produced = os.path.join(
        outdir, os.path.splitext(os.path.basename(src))[0] + ".pdf")
    _finish(produced, dst, "libreoffice")


def _convert_powerpoint(src: str, dst: str, timeout: float | None = None,
                        attended: bool = True) -> None:
    """native PowerPoint（macOS: AppleScript / Windows: COM）で変換する．

    Args:
        attended: 人がこの端末を見ているか．False なら止まったときに PowerPoint を
            前面化しない（``_macos_run_applescript``）．

    Raises:
        _Unavailable: この環境に PowerPoint が無いとき（``auto`` はこれだけを握る）．
        PdfError: PowerPoint はあったが変換に失敗したとき．
    """
    src_abs = os.path.abspath(src)
    dst_abs = os.path.abspath(dst)
    if sys.platform == "darwin":
        if not _macos_powerpoint_installed():
            raise _Unavailable(
                "PowerPoint is not installed (not in /Applications and unknown "
                "to LaunchServices)")
        # macOS は osascript 経由で実 PowerPoint を叩く．
        # スクリプト本体は stdin で，入出力パスは argv で渡す．
        # 文書を開く前に非表示で起動しておく——さもないと変換のたびにウィンドウが出る．
        _macos_prelaunch_powerpoint_hidden()
        stage = _macos_container_tmp()
        if stage is None:
            # コンテナが見つからない（サンドボックス外のビルド等）．その場で変換する．
            # この経路では未承認の場所を渡すとファイルアクセスの承認ダイアログが出る．
            _macos_run_applescript(src_abs, dst_abs, timeout, attended)
        else:
            # PowerPoint には**自分のコンテナの中だけ**を触らせる．承認ダイアログを
            # 出さずに済み，入出力がどこにあっても（/tmp でも外部ボリュームでも）動く．
            # 片付けを workdir.discard に任せる理由は _convert_libreoffice と同じ
            # （TemporaryDirectory は片付けの失敗で例外を投げる．Issue #58）．
            try:
                work = workdir.create(stage, prefix="md2pptx-")
            except OSError as e:
                # 上（_convert_libreoffice）と同じ理由で PdfError にする．
                raise PdfError(
                    f"cannot create a working directory in {stage} ({e})") from e
            try:
                staged_src = os.path.join(work, os.path.basename(src_abs))
                staged_dst = os.path.join(work, "out.pdf")
                shutil.copy2(src_abs, staged_src)
                _macos_run_applescript(staged_src, staged_dst, timeout, attended)
                _finish(staged_dst, dst_abs, "powerpoint")
            finally:
                workdir.discard(work)
    elif sys.platform.startswith("win"):
        # PowerShell + COM．32 = ppSaveAsPDF．パスは単一引用符文字列に埋めるので，
        # パス内の ' は '' にエスケープする（O'Brien 等でコマンドが壊れるのを防ぐ）．
        src_ps = src_abs.replace("'", "''")
        dst_ps = dst_abs.replace("'", "''")
        ps = (
            # COM の失敗は既定では非ゼロ終了にならず returncode 検査をすり抜ける．
            # Stop にして例外＝非ゼロで終わらせ、原因を拾えるようにする．
            "$ErrorActionPreference = 'Stop'; "
            "$ppt = New-Object -ComObject PowerPoint.Application; "
            # ここまで来れば PowerPoint は在る．以降の失敗は「無い」ではなく「失敗」．
            f"Write-Output '{_WIN_COM_READY}'; "
            "$pres = $ppt.Presentations.Open("
            f"'{src_ps}', $true, $false, $false); "
            f"$pres.SaveAs('{dst_ps}', 32); "
            "$pres.Close(); $ppt.Quit()"
        )
        try:
            proc = subprocess.run(["powershell", "-NoProfile", "-Command", ps],
                                  capture_output=True, text=True)
        except FileNotFoundError:
            raise _Unavailable("powerpoint: command not found: powershell")
        if proc.returncode != 0:
            detail = (proc.stderr or proc.stdout or "").strip().splitlines()
            tail = detail[-1] if detail else f"exit {proc.returncode}"
            if _WIN_COM_READY not in (proc.stdout or ""):
                # COM オブジェクトすら作れなかった＝PowerPoint が入っていない．
                raise _Unavailable(f"PowerPoint is not available: {tail}")
            raise PdfError(f"powerpoint failed: {tail}")
    else:
        raise _Unavailable("native PowerPoint is only available on macOS or Windows")
    # osascript/COM はいずれも無音失敗（exit 0 でも PDF が無い/空）がありうるので，
    # 終了コードだけでなく成果物の存在と非空を成功条件にする．
    if not os.path.isfile(dst_abs) or os.path.getsize(dst_abs) == 0:
        raise PdfError("powerpoint did not produce a (non-empty) PDF")


def _convert_custom(command: str, src: str, dst: str,
                    timeout: float | None = None) -> None:
    """任意のコマンド行で変換する．プレースホルダを置換して実行する．"""
    outdir = os.path.dirname(os.path.abspath(dst)) or "."
    parts = shlex.split(command)
    if not parts:
        raise PdfError(f"empty {ENV_CONVERTER}/--pdf-converter command")
    # 判定はすべて分割後のトークン（parts）で行い，元文字列との二重基準を避ける．
    # ツールが PDF をどこへ書くかは指定形式で決まる：
    #   {output} あり … その場所へ直接書く → dst をそのまま検査
    #   {outdir} あり … そのディレクトリに <入力 basename>.pdf を書く（soffice 方式）
    #   どちらも無し  … 入力の隣に <入力 basename>.pdf を書く（出力パス非対応ツール．{input} を補う）
    has_output = any("{output}" in p for p in parts)
    has_outdir = any("{outdir}" in p for p in parts)
    has_input = any("{input}" in p for p in parts)
    if not (has_output or has_outdir or has_input):
        parts.append("{input}")
    subst = {"input": src, "output": dst, "outdir": outdir}
    cmd = [p.format(**subst) for p in parts]
    _run(cmd, "converter", timeout=timeout)
    if has_output:
        # ツールが {output} をそのまま書いたはず．そこに無ければ失敗．
        if not os.path.isfile(dst):
            raise PdfError(f"converter did not write {dst}")
        return
    base = os.path.splitext(os.path.basename(src))[0] + ".pdf"
    if has_outdir:
        produced = os.path.join(outdir, base)
    else:
        produced = os.path.join(os.path.dirname(os.path.abspath(src)), base)
    _finish(produced, dst, "converter")


def _finish(produced: str, dst: str, what: str) -> None:
    """ツールが書いた PDF(produced) を期待パス(dst) へ収める．"""
    produced = os.path.abspath(produced)
    dst = os.path.abspath(dst)
    if not os.path.isfile(produced):
        raise PdfError(f"{what} did not produce a PDF (expected {produced})")
    if produced != dst:
        try:
            os.replace(produced, dst)   # 同一デバイスならアトミック
        except OSError:
            # produced と dst が別ファイルシステム（EXDEV）だと os.replace は失敗する．
            # 例: 入力の隣（/tmp）に書かせ，dst が別マウント上のとき．コピー＋削除で凌ぐ．
            shutil.copy2(produced, dst)
            # dst は書けたので変換は成功．元ファイルの削除に失敗しても（残骸が残るだけ
            # なので）成否は変えない——未捕捉の OSError で落とさない．
            try:
                os.remove(produced)
            except OSError:
                pass


def default_pdf_path(output_pptx: str) -> str:
    """--pdf を PATH 無しで使ったときの既定 PDF パス（出力 pptx と同じ場所・basename）．"""
    return os.path.splitext(output_pptx)[0] + ".pdf"


def convert(src: str, dst: str, converter: str | None,
            timeout: float | None = None, *, unattended: bool = False) -> None:
    """src(pptx) を dst(pdf) へ変換する．

    Args:
        src: 入力 pptx．
        dst: 出力 pdf．
        converter: 変換器の指定．None または "auto" で自動探索
            （PowerPoint → LibreOffice）．"powerpoint" / "libreoffice" で名指し．
            それ以外は任意のコマンド行として解釈する．
        timeout: 待ちの上限（秒）．``0`` で無制限，``None`` で「指定なし」．
            指定なしのときは ``MD2PPTX_PDF_TIMEOUT``，それも無ければ tty かどうかで
            決まる（``_resolve_timeout``）．
        unattended: 端末が tty でも「人は見ていない」として扱う（``--watch`` 用）．
            待ちの上限と，止まったときの PowerPoint 前面化の両方に効く．

    Raises:
        PdfError: 変換に失敗したとき（cli が警告に整形する）．``auto`` でも，
            **使える変換器が失敗したら**そのまま失敗する（次の変換器へは落とさない）．
    """
    if not os.path.isfile(src):
        raise PdfError(f"pptx not found: {src}")
    # 出力先ディレクトリの不在は，各バックエンドで「PDF ができない」曖昧な失敗に
    # なる．ここで一度だけ明示エラーにする（自動生成はしない——利用者の明示パス
    # を尊重し，タイポで勝手にディレクトリを作らない）．
    dst_dir = os.path.dirname(os.path.abspath(dst))
    if not os.path.isdir(dst_dir):
        raise PdfError(f"output directory does not exist: {dst_dir}")

    name = (converter or "auto").strip()
    limit = _resolve_timeout(timeout, unattended)
    # 出力はその場では作らない．同じディレクトリに使い捨ての作業場所を作り，そこで
    # 変換してから os.replace で置き換える．作業場所は dst_dir の**中に新しく作る**
    # ディレクトリなので dst と必ず同一ファイルシステム上にあり，置き換えは常に
    # アトミック——EXDEV は起こりえないので shutil.move（コピー＋削除）への
    # フォールバックは要らない．むしろ非アトミックな経路を足すと，下記 1) の
    # 「消えている時間を作らない」が崩れる．
    #
    # 1) **変換中に dst が「消えている」時間を作らない**．PDF ビューアはフォルダを監視
    #    していて，削除を確定するとそのファイルを監視集合から外す（LaTeX Workshop は
    #    250ms で確定し，集合が空になるとウォッチャごと破棄する）．変換は 1 秒から数秒
    #    かかるので必ず確定してしまい，以後どれだけ作り直しても再読込されない
    #    ——編集しながらのプレビューが最初のリビルドで死ぬ．
    # 2) 「無音失敗した変換器が残した前回の PDF を成功と誤判定する」問題（macOS の
    #    save as PDF は無音失敗しうる）は，**毎回まっさらな別名へ書かせる**ことで
    #    構造的に消える．以前は dst を先に消して防いでいたが，それは 1) と両立しない．
    # 3) ファイルではなくディレクトリなのは，出力パスを取らない変換器のため．
    #    LibreOffice は --outdir に <入力 basename>.pdf を書くので，slide.pptx →
    #    slide.pdf の既定運用では **dst を直接・逐次的に書いていた**．作業場所を
    #    挟むと outdir がそちらへ移り，この衝突も消える．
    try:
        work = workdir.create(dst_dir)
    except OSError as e:
        # 作業場所すら作れない（多くは出力先の書き込み権限）．**PDF 変換の失敗として
        # 扱う**——素の OSError を通すと cli の `except PdfError` をすり抜けて
        # SystemExit になり，pptx は保存できているのに終了コードが 1 になる．
        # 「PDF が作れなくても pptx は成功」（#39）が崩れ，編集しながらの運用が
        # 出力先の権限ひとつで止まる．
        raise PdfError(f"cannot create a working directory in {dst_dir} ({e})")
    # 後片付けは TemporaryDirectory ではなく自前で行う．with で包むと，変換の本体が
    # 投げた**想定外の**例外まで巻き込んで扱いを変えてしまう．ここで OSError を
    # PdfError に読み替えてよいのは「作業場所を用意できなかった」ときだけで，
    # 想定外の例外はトレースバックのまま伝播させる（バグを隠さない）．
    try:
        staged = os.path.join(work, os.path.basename(dst))
        try:
            _dispatch(name, src, staged, limit, not unattended)
            try:
                os.replace(staged, dst)
            except OSError as e:
                raise PdfError(f"cannot replace existing PDF: {dst} ({e})")
        except PdfError:
            # 失敗したら古い PDF は残さない．PDF 変換の失敗は終了コードを変えない
            # ので，警告を見落とした人が前回の内容を新しい出力と取り違えてしまう．
            # **置き換えに失敗したときも同じ**——変換自体は成功していても，そこに
            # 残っているのは前回の内容だから．書きかけは work ごと消える．
            # なお「消せるものだけ消える」のは意図どおり．rename も unlink も権限は
            # 親ディレクトリで決まるので，権限で置き換えられなかった dst は削除でき
            # ない（実測でどちらも EACCES）．無事な出力を巻き添えにはしない．
            try:
                os.remove(dst)
            except OSError:
                pass
            raise
    finally:
        workdir.discard(work)


def _dispatch(name: str, src: str, dst: str, limit: float | None,
              attended: bool = True) -> None:
    """変換器を選んで実行する（``convert`` の下請け）．"""
    if name == "auto":
        # auto は「**使えるものを探す**」だけ．実 PowerPoint（テーマ忠実度が高い）を
        # 優先し，無ければ LibreOffice を使う．
        # 在る物が失敗したときに次へ落とすことはしない（Issue #46）——忠実度という
        # 成果物の性質が黙って入れ替わるうえ，隠れる原因（オートメーション承認の拒否，
        # ライセンス未認証など）は利用者が直せるものだから．
        missing: list[str] = []
        try:
            _convert_powerpoint(src, dst, limit, attended)
            return
        except _Unavailable as e:
            # _Unavailable は PdfError のサブクラスなので，この except は必ず
            # PdfError より**先**に置くこと．入れ替えると失敗まで握って次の変換器へ
            # 落ちる＝#46 で消した挙動が黙って戻る．
            missing.append(str(e))
        except PdfError as e:
            # 案内は LibreOffice が実際に使えるときだけ添える（無い物を勧めない）．
            alt = ("; use --pdf-converter libreoffice to convert without PowerPoint"
                   if _which_libreoffice() else "")
            raise PdfError(f"{e}{alt}")
        try:
            # attended は渡さない——LibreOffice は --headless で走り，前面に出せる
            # ウィンドウも人が答えるダイアログも無い（この引数が効くのは macOS の
            # PowerPoint 経路だけ）．増やすなら _convert_libreoffice の側で受ける．
            _convert_libreoffice(src, dst, limit)
            return
        except _Unavailable as e:
            missing.append(str(e))
        raise PdfError(
            "no PDF converter available "
            "(tried PowerPoint / LibreOffice; use --pdf-converter or install "
            "LibreOffice)\n  - " + "\n  - ".join(missing))

    if name == "libreoffice":
        _convert_libreoffice(src, dst, limit)
    elif name == "powerpoint":
        _convert_powerpoint(src, dst, limit, attended)
    else:
        _convert_custom(name, src, dst, limit)
