#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""図表の座標計算に共通のプリミティブ（md2pptx．DESIGN.md §4.1）．

flow / seq などの図 DSL は「DSL を解釈して IR にする」と「IR を矩形領域へ
置いて座標を出す」を分け、**座標計算まで純 Python（EMU 整数）で完結**させる。
python-pptx を持ち込まないので、図の配置は pptx を書かずに試せる。

このモジュールはその共通部分——単位換算と、**どの DSL にも依存しない**幾何だけを持つ。
図ごとの並べ方（何をどこに置くか）も、その DSL 固有のノードを抱える入れ物
（flow の ``PlacedNode`` など）も、各 DSL のモジュールが持つ。
ここに 1 つの DSL の型を持ち込むと、次の DSL がそれを再利用できずに詰む。

**位置を名前で持つ。** 位置で持つと取り違えても誰も気づかない——幅と高さを
入れ替えても、右端と下端を取り違えても、型としては同じ整数で通ってしまい、
図が歪んで出てくるまで分からない（Issue #71）。
"""
from __future__ import annotations

from dataclasses import dataclass

EMU = 914400  # 1 インチ = 914400 EMU


def emu(inch: float) -> int:
    """インチを EMU 整数へ換算する．"""
    return int(inch * EMU)


@dataclass(frozen=True)
class Rect:
    """矩形．端や中心は毎回足し算で出さず，ここから読む．"""
    left: int
    top: int
    width: int
    height: int

    @property
    def right(self) -> int:
        return self.left + self.width

    @property
    def bottom(self) -> int:
        return self.top + self.height

    @property
    def center_x(self) -> int:
        return self.left + self.width // 2

    @property
    def center_y(self) -> int:
        return self.top + self.height // 2


@dataclass(frozen=True)
class PlacedText:
    """文字だけの要素（省略記号・矢印ラベル・キャプション）．

    ``align`` は枠の中での**横の寄せ**．既定は中央で、矢印の真上に置くラベルや
    キャプションはこれでよい。**縦並びの矢印ラベルだけ左寄せ**にする（Issue #176）
    ——枠幅は字幅の見積もりなので、実際の字がそれを超えると中央揃えでは左右へ
    等しくはみ出し、左側が矢印に掛かる。左寄せなら余りは矢印から離れる側へ出る。
    """
    text: str
    rect: Rect
    align: str = "center"


@dataclass(frozen=True)
class PlacedArrow:
    """要素どうしを結ぶ矢印．始点から終点への線分で，太さは render が決める．"""
    x1: int
    y1: int
    x2: int
    y2: int


@dataclass(frozen=True)
class PlacedLine:
    """矢尻の無い線分（シーケンス図のライフラインなど）．

    ``PlacedArrow`` と分けてあるのは**描き分けが要る**から——矢尻の有無は
    見た目の差ではなく「向きがあるか」の差で、render はそれで図形を選ぶ。
    ``dashed`` は破線（時間の経過や省略を表す線に使う）。
    """
    x1: int
    y1: int
    x2: int
    y2: int
    dashed: bool = False
