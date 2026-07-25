#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""使い捨ての作業ディレクトリの片付け（Issue #58）．

md2pptx は成果物をアトミックに差し替えるために，出力先の隣へ使い捨ての作業ディレクトリ
を作ってそこで組み立てる（pptx は ``render.Renderer.save``，PDF は ``pdf.convert``）．
ほかにも thmx の展開先や LibreOffice の使い捨てプロファイルなど，**「自分で作って自分で
捨てる作業ディレクトリ」は 5 箇所**ある．その片付け方をここに 1 つだけ置く．

規則は 2 つで，どちらも外すと運用が壊れる．

- **片付けの失敗で処理の成否を変えない．** 片付けに入る時点で本来の仕事（保存・変換）は
  終わっている．そこで例外を投げると，成功した実行が失敗になり，しかも本体が投げた例外が
  あればそれを握りつぶして置き換えてしまう．
- **それでも黙って残さない．** 消せなかったことを誰にも伝えないと，``--watch`` では保存の
  たびに 1 つずつ溜まっていく．しかも出力先ディレクトリに溜まるものは利用者の目に触れる．
  原因は環境側（Windows で走査ソフトがファイルを掴んでいる等）なので**ここで再試行はせず**，
  起きた事実だけを伝える．

メッセージは利用者に見える文字列なので，写しを増やさない意味でもここが唯一の出どころ．
python-pptx を含め外部依存は持たない（render / pdf / thmx2pptx のどれからも使う）．
"""
from __future__ import annotations

import os
import shutil
import sys
import tempfile


def create(where: str | None = None, prefix: str = ".md2pptx-") -> str:
    """作業ディレクトリを作ってパスを返す．

    Args:
        where: 作る場所．``None`` ならシステムの一時領域．成果物を組み立てるときは
            **出力先と同じディレクトリ**を渡す——同一ファイルシステム上に置くことで，
            仕上げの ``os.replace`` が必ずアトミックになる（EXDEV が起こりえない）．
        prefix: 名前の接頭辞．既定がドット始まりなのは，**出力先に作る**もの（既定の
            使い方）を利用者の目に付かせないため．システムの一時領域へ作るときは
            隠す理由が無いので，``prefix="md2pptx-lo-"`` のように明示して渡す．

    Raises:
        OSError: 作れなかったとき．**ここでは整形しない**——「作業場所を用意できなかった」
            ことをどう伝えるかは呼び出し側の事情で違うため（``pdf.convert`` は
            ``PdfError`` にして終了コードを変えず，``render.save`` は何をしようとして
            失敗したかを添えて送出する）．
    """
    return tempfile.mkdtemp(dir=where, prefix=prefix)


def discard(work: str) -> None:
    """作業ディレクトリを捨てる．**例外は投げない**（モジュール docstring の規則）．

    消せなければ stderr に 1 行出すだけで、処理の成否は変えない．
    """
    # ignore_errors=True なのは，エラーが起きても走査を続けて**消せるものは消す**ため．
    # 外すと最初のエラーで止まり，残りが丸ごと残る．
    shutil.rmtree(work, ignore_errors=True)
    if os.path.isdir(work):
        sys.stderr.write(f"md2pptx: warning: could not remove {work}\n")
