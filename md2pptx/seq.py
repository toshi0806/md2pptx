#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""シーケンス図 DSL（``` ```seq ``` ブロック）のパーサ＋レイアウタ（md2pptx）．

DESIGN.md §5.5.1 の DSL を解釈して ir.Seq へ変換する ``parse_seq`` と，
描画用の座標プラン（純粋な EMU 計算）を返す ``plan_seq`` を提供する．
python-pptx には依存しない（描画は render.py の責務）——``flow.py`` と同じ形．

ネットワークの講義でいちばん出てくる図がこれで、**時間軸を持つ**ため flow の
格子（Issue #109）では表現できない。TCP 3-way handshake・輻輳ウィンドウ・
Fast Retransmission・HTTP・SMTP/POP・DNS 反復問い合わせと、ほぼ毎回出てくる。

DSL 例::

    lifelines: 送信側, 受信側
    note(top): コネクションの確立
    送信側 -> 受信側: SYN
    受信側 -> 送信側: SYN+ACK
    note: ここで確立
    送信側 -> 受信側: ACK
    caption: 3-way handshake

**文法は3要素だけ**に絞ってある。activation box（実行中を表す縦長の箱）・
``alt`` / ``loop`` などの制御構造・参加者の生成消滅は入れない——
入れ始めると際限が無く、講義スライドで要るのは往復の線とラベルだけだった。
"""
from __future__ import annotations

import re
from dataclasses import dataclass, field

from .ir import Seq, SeqMessage, SeqNote
from .layout import PlacedArrow, PlacedLine, PlacedText, Rect, emu as _emu

# 設定行．``caption`` / ``note(top)`` / ``note(bottom)`` は flow と共通の書き方に
# 揃えてある（同じものを2通りで書けるようにしない）．``lifelines`` はこの DSL 固有．
_RE_SETTING = re.compile(
    r"^(lifelines|direction|caption|note\(top\)|note\(bottom\))\s*:\s*(.*)$")
# 図の中の注記．設定行と紛れないよう ``note(...)`` より後に判定する．
_RE_NOTE = re.compile(r"^note\s*:\s*(.*)$")
# メッセージ．``A -> B`` ／ ``A -> B: ラベル``．
# ラベルは**最初の ':' だけで割る**——"ack 1: seq 100" のような書き方が普通に出る．
_RE_MESSAGE = re.compile(r"^(.+?)\s*->\s*([^:]+?)\s*(?::\s*(.*))?$")

# --- 配置の定数（すべて EMU）．
_HEAD_H = _emu(0.62)          # ライフラインの頭（名前の箱）の高さ
_HEAD_GAP = _emu(0.14)        # 頭と縦線の間
# ラベルの高さ．本文サイズはテーマ次第で 30pt にもなるので、**行の高さぶんは
# 確保する**——足りないと矢印の線に文字が重なって読めない（実 PowerPoint で確認）．
_LABEL_H = _emu(0.44)
_ROW_MIN = _emu(0.46)         # メッセージ1本あたりの最小の縦送り
_ROW_MAX = _emu(0.80)         # 同・最大（少ないときに間延びさせない）
# ラベル幅の見積もりに使う想定サイズ（flow と同じ理由）．
# **render もこの大きさで描くこと**——本文標準サイズで描くと、本文の大きい
# テーマでは見積もりの倍近い字が入り、``wrap=False`` の注記がスライドの外へ
# 出る（Issue #167）．外から使うので公開名も置く．
_LABEL_PT = LABEL_PT = 16.0


@dataclass
class SeqPlan:
    """図ひとつぶんの配置．

    ``note_top`` / ``note_bottom`` はここに入らない——地の文なので本文
    プレースホルダへ入れるのが render の仕事で、図の座標を持たない（flow と同じ）。
    """

    heads: list[PlacedText] = field(default_factory=list)
    lines: list[PlacedLine] = field(default_factory=list)
    arrows: list[PlacedArrow] = field(default_factory=list)
    labels: list[PlacedText] = field(default_factory=list)
    notes: list[PlacedText] = field(default_factory=list)
    captions: list[PlacedText] = field(default_factory=list)


# ---------------------------------------------------------------- パース

def parse_seq(text: str) -> Seq:
    """``` ```seq ``` ブロック本文を Seq（IR）へ変換する．"""
    seq = Seq()
    declared = False
    after_step = False        # 直前に ``@step`` を読んだか（注記の since 用）
    names: list[str] = []

    def index_of(name: str) -> int:
        """名前をライフラインの添字にする（未宣言なら出てきた順に足す）．"""
        if name in names:
            return names.index(name)
        if declared:
            known = ", ".join(names) or "(none)"
            raise ValueError(
                f"unknown lifeline: {name!r} (declared: {known})")
        names.append(name)
        return len(names) - 1

    for raw in text.split("\n"):
        line = raw.strip()
        if not line:
            continue

        m = _RE_SETTING.match(line)
        if m:
            key, val = m.group(1), m.group(2).strip()
            if key == "lifelines":
                names = [p.strip() for p in val.split(",") if p.strip()]
                declared = True
            elif key == "caption":
                seq.caption = val or None
            elif key == "note(top)":
                seq.note_top = val or None
            elif key == "note(bottom)":
                seq.note_bottom = val or None
            continue

        if line == "@step":
            # 図の中の段階（Issue #125）．**その時点までの本数**を覚えておき、
            # parser がスライドの段へ展開する．図の座標には関係しない．
            seq.steps.append(len(seq.messages))
            after_step = True
            continue

        mn = _RE_NOTE.match(line)
        if mn:
            # 図の中の注記．**その時点までに引いた矢印の本数**を覚えておき、
            # 配置のときに時間軸のその位置へ置く．
            # 置く高さ（after）と出す段（since）は別．``@step`` の後ろに
            # 書いた注記は**次の矢印が出る段**から見せる——その注記は次に
            # 起きることの説明なので、先に出ると答えが見えてしまう．
            at = len(seq.messages)
            seq.notes.append(SeqNote(
                text=mn.group(1).strip(), after=at,
                since=at + 1 if after_step else at))
            continue

        mm = _RE_MESSAGE.match(line)
        if mm:
            after_step = False
            src_name = mm.group(1).strip()
            dst_name = mm.group(2).strip()
            label = (mm.group(3) or "").strip() or None
            if src_name == dst_name:
                raise ValueError(
                    f"a lifeline cannot send to itself: {src_name!r} "
                    "(self-messages are out of scope)")
            src = index_of(src_name)
            dst = index_of(dst_name)
            seq.messages.append(SeqMessage(src=src, dst=dst, label=label))
            continue

        raise ValueError(
            f"invalid seq syntax: {line!r} (expected 'A -> B: label', "
            "'lifelines: …', 'note: …', 'caption: …', "
            "or 'note(top): …' / 'note(bottom): …')")

    seq.lifelines = names
    return seq


# ---------------------------------------------------------------- 配置

def _label_width(text: str) -> int:
    """ラベルの想定幅（EMU）．全角は半角の2倍で数える（flow と同じ見積もり）．

    render はラベルを折り返さない設定で描くので、**この幅が効くのは中心の位置**
    であって読めるかどうかではない（Issue #111 と同じ理由）．
    """
    # 0x2E80 より上を全角とみなす．CJK の記号・かな・漢字（U+3000〜）だけでなく、
    # 全角英数と全角記号（U+FF01〜）も上側に入るので、実用上これで足りる
    # （"（" U+FF08 も "Ａ" U+FF21 も 2 文字ぶんとして数えられることを確認済み）．
    units = sum(2 if ord(c) > 0x2E80 else 1 for c in text)
    return int(units * _LABEL_PT * 12700 * 0.55) + _emu(0.12)


def plan_seq(seq: Seq, left: int, top: int, width: int, height: int) -> SeqPlan:
    """Seq を矩形領域 (left, top, width, height) に配置する．"""
    plan = SeqPlan()
    n = len(seq.lifelines)
    if n == 0:
        return plan

    cap_h = _emu(0.42) if seq.caption else 0
    cap_gap = _emu(0.10) if seq.caption else 0
    body_h = height - cap_h - cap_gap

    # ライフラインは幅いっぱいに等間隔．端は頭の箱がはみ出さないよう内側へ寄せる．
    col_w = width // n
    xs = [left + col_w * i + col_w // 2 for i in range(n)]

    for i, name in enumerate(seq.lifelines):
        # 列幅いっぱいまで使う．狭めると名前が縦に潰れて読めなくなる．
        w = min(col_w - _emu(0.12), max(_emu(1.0), _label_width(name)))
        plan.heads.append(PlacedText(
            name, Rect(xs[i] - w // 2, top, w, _HEAD_H)))

    # 縦線の範囲．メッセージ数で送り幅を決め、**多いときは詰め、少ないときは
    # 間延びさせない**（Fast Retransmission は矢印25本超、handshake は3本）．
    line_top = top + _HEAD_H + _HEAD_GAP
    avail = max(_emu(0.5), body_h - _HEAD_H - _HEAD_GAP)
    rows = len(seq.messages)
    if rows:
        # 帯を rows+1 で割った高さが基準．少ないときに間延びしないよう上限で
        # 抑え、多いときは下限で広げる．
        fit = avail // (rows + 1)
        row_h = max(_ROW_MIN, min(_ROW_MAX, fit))
        # ただし**下限より帯に収めるほうを優先する**．本数が多いと下限では帯を
        # 超え、図が下のスライドへはみ出す．読みにくくなっても切れるよりよい——
        # 読みにくさは見れば分かるが、はみ出しは PowerPoint を開くまで気づけない．
        row_h = min(row_h, fit) if fit else _ROW_MIN
    else:
        row_h = _ROW_MIN
    line_bottom = min(top + body_h, line_top + row_h * (rows + 1))

    for x in xs:
        plan.lines.append(PlacedLine(x, line_top, x, line_bottom))

    for k, msg in enumerate(seq.messages):
        y = line_top + row_h * (k + 1)
        x1, x2 = xs[msg.src], xs[msg.dst]
        plan.arrows.append(PlacedArrow(x1, y, x2, y))
        if msg.label:
            w = max(_emu(0.5), _label_width(msg.label))
            plan.labels.append(PlacedText(msg.label, Rect(
                (x1 + x2) // 2 - w // 2, y - _LABEL_H, w, _LABEL_H)))

    for note in seq.notes:
        # 注記は「その本数まで引いた直後」の高さへ．矢印とラベルは中央寄りに
        # あるので、注記は**右端に寄せて**重なりを避ける．
        y = line_top + row_h * note.after + row_h // 2
        w = max(_emu(0.6), _label_width(note.text))
        plan.notes.append(PlacedText(
            note.text, Rect(min(left + width - w, xs[-1] + _emu(0.12)),
                            y - _LABEL_H // 2, w, _LABEL_H)))

    if seq.caption:
        plan.captions.append(PlacedText(
            seq.caption, Rect(left, line_bottom + cap_gap, width, cap_h)))
    return plan
