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
案内を stderr に出し PowerPoint を前面に出す（変換自体は中断しない）．

このモジュールは cli 以外に依存しない（python-pptx 非依存）．外部プロセスの
起動と，どのバイナリを使うかの解決だけを担う．
"""
from __future__ import annotations

import os
import pathlib
import shlex
import shutil
import subprocess
import sys
import tempfile


class PdfError(Exception):
    """PDF 変換の失敗（原因メッセージ付き）．cli が警告表示に使う．"""


class _Unavailable(PdfError):
    """その変換器がこの環境に**無い**（＝失敗ではない）．

    ``auto`` はこれだけを握って次の変換器へ進む．無い物を飛ばしても何も失われない
    のに対し，**在る物の失敗**を飛ばすと忠実度の違う PDF を黙って掴ませることに
    なる（Issue #46）．名指し指定のときは PdfError として そのまま利用者に届く．
    """


# 環境変数名（CLI 引数 --pdf-converter が優先）．
ENV_CONVERTER = "MD2PPTX_PDF_CONVERTER"

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
            capture_output=True, text=True)
    except OSError:
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
                       capture_output=True)
    except OSError:
        pass


def _macos_container_tmp() -> str | None:
    """PowerPoint のサンドボックスコンテナ内の作業場所（無ければ None）．

    ここに pptx を置いてから開かせると，**ファイルアクセスの承認ダイアログが出ない**．
    コンテナはアプリ自身のサンドボックス領域なので承認の対象外だからで，これにより
    どの場所の入出力でも（``/tmp`` でもネットワークボリュームでも）変換できる．
    """
    base = os.path.expanduser(
        "~/Library/Containers/com.microsoft.Powerpoint/Data/tmp")
    return base if os.path.isdir(base) else None


def _macos_run_applescript(src: str, dst: str) -> None:
    """AppleScript で src → dst を変換する．長引いたら理由を stderr に出す．

    PowerPoint を隠して動かしているので，何かのダイアログ（オートメーションの承認
    など）で止まると，利用者には Dock を見ない限り「無音で固まった」ようにしか見え
    ない．そこで ``_MACOS_HINT_AFTER`` 秒たっても終わらなければ，何が起きている
    可能性があるかを stderr に出し，PowerPoint を前面に出してダイアログを見せる．
    **変換は中断しない**——単に遅いだけのときに壊さないため．
    """
    proc = subprocess.Popen(
        ["osascript", "-", src, dst],
        stdin=subprocess.PIPE, stdout=subprocess.PIPE, stderr=subprocess.PIPE,
        text=True)
    try:
        _, err = proc.communicate(input=_APPLESCRIPT_PPTX_TO_PDF,
                                  timeout=_MACOS_HINT_AFTER)
    except subprocess.TimeoutExpired:
        sys.stderr.write(
            "md2pptx: PowerPoint is taking longer than "
            f"{_MACOS_HINT_AFTER:.0f}s — it may be waiting for a dialog "
            "(e.g. the automation approval). Bringing it to the front; "
            "answer the dialog to continue.\n")
        try:
            subprocess.run(["open", "-a", "Microsoft PowerPoint"],
                           capture_output=True)
        except OSError:
            pass
        _, err = proc.communicate()
    if proc.returncode != 0:
        detail = (err or "").strip().splitlines()
        tail = detail[-1] if detail else f"exit {proc.returncode}"
        raise PdfError(f"powerpoint failed: {tail}")


def _run(cmd: list[str], what: str, input: str | None = None) -> None:
    """外部コマンドを実行し，失敗を PdfError に変換する．

    成功時の出力は捨てる．失敗時のみ stderr（無ければ stdout）の末尾 1 行を
    原因として拾う（cli が警告に整形する）．input を渡すと stdin に流す
    （osascript にスクリプト本体を与えるのに使う）．
    """
    try:
        proc = subprocess.run(cmd, capture_output=True, text=True, input=input)
    except FileNotFoundError:
        raise PdfError(f"{what}: command not found: {cmd[0]}")
    if proc.returncode != 0:
        detail = (proc.stderr or proc.stdout or "").strip().splitlines()
        tail = detail[-1] if detail else f"exit {proc.returncode}"
        raise PdfError(f"{what} failed: {tail}")


def _convert_libreoffice(src: str, dst: str) -> None:
    """LibreOffice で src(pptx) → dst(pdf)．--outdir 方式なので後で改名する．"""
    soffice = _which_libreoffice()
    if soffice is None:
        raise _Unavailable(
            "LibreOffice not found (looked for soffice/libreoffice on PATH "
            "and the default install location)")
    outdir = os.path.dirname(os.path.abspath(dst)) or "."
    # 同一プロファイルの多重起動は失敗しうるので，毎回使い捨てのプロファイルを渡す．
    with tempfile.TemporaryDirectory(prefix="md2pptx-lo-") as profile:
        # as_uri() は Windows のドライブレターも file:///C:/... と正しく組む
        # （手組みの "file://"+path だと file://C:/... になり不正）．
        uri = pathlib.Path(os.path.abspath(profile)).as_uri()
        _run([
            soffice, "--headless",
            f"-env:UserInstallation={uri}",
            "--convert-to", "pdf", "--outdir", outdir, src,
        ], "libreoffice")
    # soffice は <入力 basename>.pdf を outdir に書く．期待名と違えば移動する．
    # 使い捨てプロファイル（with）は変換が終わった時点で不要なので，PDF の移動は
    # with を抜けてから行う（プロファイルの寿命と成果物の移動を分離）．
    produced = os.path.join(
        outdir, os.path.splitext(os.path.basename(src))[0] + ".pdf")
    _finish(produced, dst, "libreoffice")


def _convert_powerpoint(src: str, dst: str) -> None:
    """native PowerPoint（macOS: AppleScript / Windows: COM）で変換する．

    Raises:
        _Unavailable: この環境に PowerPoint が無いとき（``auto`` はこれだけを握る）．
        PdfError: PowerPoint はあったが変換に失敗したとき．
    """
    src_abs = os.path.abspath(src)
    dst_abs = os.path.abspath(dst)
    if sys.platform == "darwin":
        if not _macos_powerpoint_installed():
            raise _Unavailable(
                "PowerPoint is not installed (no /Applications/Microsoft PowerPoint.app)")
        # macOS は osascript 経由で実 PowerPoint を叩く．
        # スクリプト本体は stdin で，入出力パスは argv で渡す．
        # 文書を開く前に非表示で起動しておく——さもないと変換のたびにウィンドウが出る．
        _macos_prelaunch_powerpoint_hidden()
        stage = _macos_container_tmp()
        if stage is None:
            # コンテナが見つからない（サンドボックス外のビルド等）．その場で変換する．
            # この経路では未承認の場所を渡すとファイルアクセスの承認ダイアログが出る．
            _macos_run_applescript(src_abs, dst_abs)
        else:
            # PowerPoint には**自分のコンテナの中だけ**を触らせる．承認ダイアログを
            # 出さずに済み，入出力がどこにあっても（/tmp でも外部ボリュームでも）動く．
            with tempfile.TemporaryDirectory(dir=stage, prefix="md2pptx-") as work:
                staged_src = os.path.join(work, os.path.basename(src_abs))
                staged_dst = os.path.join(work, "out.pdf")
                shutil.copy2(src_abs, staged_src)
                _macos_run_applescript(staged_src, staged_dst)
                _finish(staged_dst, dst_abs, "powerpoint")
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


def _convert_custom(command: str, src: str, dst: str) -> None:
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
    _run(cmd, "converter")
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


def convert(src: str, dst: str, converter: str | None) -> None:
    """src(pptx) を dst(pdf) へ変換する．

    Args:
        src: 入力 pptx．
        dst: 出力 pdf．
        converter: 変換器の指定．None または "auto" で自動探索
            （PowerPoint → LibreOffice）．"powerpoint" / "libreoffice" で名指し．
            それ以外は任意のコマンド行として解釈する．

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

    # 既存の出力は変換前に消す．残したままだと，変換が実際には失敗しても前回の PDF が
    # 「存在かつ非空」の成功条件を満たしてしまい，古い内容を見続けることになる
    # （macOS の save as PDF は無音失敗しうる）．
    try:
        os.remove(dst)
    except FileNotFoundError:
        pass
    except OSError as e:
        raise PdfError(f"cannot replace existing PDF: {dst} ({e})")

    name = (converter or "auto").strip()

    if name == "auto":
        # auto は「**使えるものを探す**」だけ．実 PowerPoint（テーマ忠実度が高い）を
        # 優先し，無ければ LibreOffice を使う．
        # 在る物が失敗したときに次へ落とすことはしない（Issue #46）——忠実度という
        # 成果物の性質が黙って入れ替わるうえ，隠れる原因（オートメーション承認の拒否，
        # ライセンス未認証など）は利用者が直せるものだから．
        missing: list[str] = []
        try:
            _convert_powerpoint(src, dst)
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
            _convert_libreoffice(src, dst)
            return
        except _Unavailable as e:
            missing.append(str(e))
        raise PdfError(
            "no PDF converter available "
            "(tried PowerPoint / LibreOffice; use --pdf-converter or install "
            "LibreOffice)\n  - " + "\n  - ".join(missing))

    if name == "libreoffice":
        _convert_libreoffice(src, dst)
    elif name == "powerpoint":
        _convert_powerpoint(src, dst)
    else:
        _convert_custom(name, src, dst)
