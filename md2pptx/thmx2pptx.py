#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""thmx → base pptx 変換（md2pptx ステージ0）．

PowerPoint テーマファイル（``.thmx``）を，python-pptx が土台として開ける
プレゼンテーション（``.pptx``）へ変換する．thmx は内部的に pptx とほぼ同型で，
差分は次の 3 点のみ（DESIGN.md §3.2）：

  1. パート配置 : ``theme/`` 配下 → pptx では ``ppt/`` 配下
  2. コンテンツタイプ : presentation が ``…template.main+xml`` → ``…presentation.main+xml``
  3. Override 欠落 : slideMaster / slideLayout の Override を追加（themeManager の Override は削除）

rels はすべて相対パスのため，ディレクトリ名を付け替えても中身の修正は不要．

関数として使う::

    from thmx2pptx import thmx_to_pptx
    base = thmx_to_pptx("theme.thmx")          # 一時ファイルへ変換しパスを返す
    base = thmx_to_pptx("theme.thmx", "b.pptx")  # 出力先を明示

単体 CLI として使う（デバッグ用）::

    python3 thmx2pptx.py theme.thmx -o base.pptx
"""
import os
import re
import shutil
import tempfile
import zipfile

# コンテンツタイプ文字列
CT_PRESENTATION = (
    "application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"
)
CT_SLIDE_MASTER = (
    "application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml"
)
CT_SLIDE_LAYOUT = (
    "application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml"
)


class ThmxError(Exception):
    """thmx 変換時のエラー．"""


def _count_layouts(extract_dir):
    """展開済みディレクトリ内の slideLayout 数を数える（テーマ差し替えに追従）．"""
    layout_dir = os.path.join(extract_dir, "theme", "slideLayouts")
    if not os.path.isdir(layout_dir):
        return 0
    return len([
        f for f in os.listdir(layout_dir)
        if re.match(r"slideLayout\d+\.xml$", f)
    ])


def _check_full_theme(extract_dir):
    """フルテーマ型（マスター＋レイアウト同梱）であることを検証する．

    色・フォントのみの簡易テーマはレイアウトを持たず，スライドの土台にできない．
    """
    master = os.path.join(extract_dir, "theme", "slideMasters", "slideMaster1.xml")
    if not os.path.isfile(master):
        raise ThmxError(
            "this thmx has no slideMaster (not a full theme). "
            "A theme with master+layouts is required to build slides."
        )
    n = _count_layouts(extract_dir)
    if n == 0:
        raise ThmxError("this thmx has no slideLayouts (not a full theme).")
    return n


def _rewrite_content_types(extract_dir, nlayout):
    """[Content_Types].xml を pptx 向けに書き換える（DESIGN.md §3.2 の 3 点）．"""
    path = os.path.join(extract_dir, "[Content_Types].xml")
    ct = open(path, encoding="utf-8").read()

    # (a) themeManager の Override を削除（pptx では不要）
    ct = re.sub(
        r'<Override PartName="/theme/theme/themeManager\.xml"[^>]*/>', "", ct
    )
    # (b) presentation の ContentType を presentation.main+xml に
    ct = ct.replace(
        "application/vnd.openxmlformats-officedocument.presentationml.template.main+xml",
        CT_PRESENTATION,
    )
    # (c) 全 PartName のパス theme/ → ppt/ （Override の参照先を付け替え）
    ct = ct.replace('PartName="/theme/', 'PartName="/ppt/')

    # (d) slideMaster / slideLayout の Override を追加（既存分は重複させない）
    adds = []
    master_pn = "/ppt/slideMasters/slideMaster1.xml"
    if f'PartName="{master_pn}"' not in ct:
        adds.append(
            f'<Override PartName="{master_pn}" ContentType="{CT_SLIDE_MASTER}"/>'
        )
    for i in range(1, nlayout + 1):
        layout_pn = f"/ppt/slideLayouts/slideLayout{i}.xml"
        if f'PartName="{layout_pn}"' not in ct:
            adds.append(
                f'<Override PartName="{layout_pn}" ContentType="{CT_SLIDE_LAYOUT}"/>'
            )
    if adds:
        ct = ct.replace("</Types>", "".join(adds) + "</Types>")

    open(path, "w", encoding="utf-8").write(ct)


def _rewrite_root_rels(extract_dir):
    """_rels/.rels の officeDocument リレーションシップを presentation 本体へ向ける．

    標準的な thmx では officeDocument は ``theme/theme/themeManager.xml`` を指す
    （presentation.xml ではない）ため，単純な文字列置換では付け替えられない．
    Type が officeDocument のリレーションシップの Target を
    ``ppt/presentation.xml`` に書き換える（presentation.xml を直接指す変種にも対応）．
    """
    path = os.path.join(extract_dir, "_rels", ".rels")
    s = open(path, encoding="utf-8").read()

    def _retarget(m):
        return re.sub(r'Target="[^"]*"', 'Target="ppt/presentation.xml"', m.group(0))

    s = re.sub(
        r'<Relationship\b[^>]*Type="http://schemas\.openxmlformats\.org/'
        r'officeDocument/2006/relationships/officeDocument"[^>]*/>',
        _retarget,
        s,
    )
    # presentation.xml を直接指す旧来の変種にも対応．
    s = s.replace("theme/presentation.xml", "ppt/presentation.xml")
    open(path, "w", encoding="utf-8").write(s)


def _repack(extract_dir, out_path):
    """展開ディレクトリを pptx（ZIP）として書き出す．[Content_Types].xml を先頭に置く．"""
    if os.path.exists(out_path):
        os.remove(out_path)
    with zipfile.ZipFile(out_path, "w", zipfile.ZIP_DEFLATED) as z:
        z.write(os.path.join(extract_dir, "[Content_Types].xml"), "[Content_Types].xml")
        for root, _, files in os.walk(extract_dir):
            for f in files:
                full = os.path.join(root, f)
                arc = os.path.relpath(full, extract_dir)
                if arc == "[Content_Types].xml":
                    continue
                z.write(full, arc)


def thmx_to_pptx(thmx_path, out_path=None):
    """thmx を base pptx に変換し，出力先パスを返す．

    out_path 省略時は一時ファイル（``*.pptx``）に書き出す（呼び出し側で破棄する）．
    """
    if not os.path.isfile(thmx_path):
        raise ThmxError(f"thmx not found: {thmx_path}")

    # out_path を内部で用意した場合のみ，失敗時に後始末する（呼び出し側指定は触らない）．
    created_out = out_path is None
    if created_out:
        fd, out_path = tempfile.mkstemp(suffix=".pptx", prefix="md2pptx-base-")
        os.close(fd)

    work = tempfile.mkdtemp(prefix="md2pptx-thmx-")
    try:
        try:
            with zipfile.ZipFile(thmx_path) as z:
                z.extractall(work)
        except zipfile.BadZipFile:
            raise ThmxError(f"not a valid thmx/zip file: {thmx_path}")

        nlayout = _check_full_theme(work)

        # theme/ ディレクトリを ppt/ へリネーム
        shutil.move(os.path.join(work, "theme"), os.path.join(work, "ppt"))

        _rewrite_root_rels(work)
        _rewrite_content_types(work, nlayout)
        _repack(work, out_path)
    except BaseException:
        # 変換失敗時，内部生成した空の一時 pptx が残らないよう掃除する．
        if created_out and os.path.exists(out_path):
            os.remove(out_path)
        raise
    finally:
        shutil.rmtree(work, ignore_errors=True)

    return out_path


def main(argv=None):
    import argparse

    ap = argparse.ArgumentParser(
        description="Convert a PowerPoint theme (.thmx) into a base .pptx."
    )
    ap.add_argument("thmx", help="input .thmx file")
    ap.add_argument("-o", "--output", help="output .pptx (default: <thmx stem>.pptx)")
    args = ap.parse_args(argv)

    out = args.output
    if out is None:
        out = os.path.splitext(os.path.basename(args.thmx))[0] + ".pptx"

    try:
        path = thmx_to_pptx(args.thmx, out)
    except ThmxError as e:
        raise SystemExit(f"thmx2pptx: {e}")
    print("base pptx:", path)


if __name__ == "__main__":
    main()
