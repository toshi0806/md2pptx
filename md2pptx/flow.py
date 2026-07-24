#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""フロー図 DSL（``` ```flow ``` ブロック）のパーサ＋レイアウタ（md2pptx Phase 3）．

DESIGN.md §5.5 の独自 DSL を解釈し，ir.Flow（FlowNode / FlowEdge）へ変換する
``parse_flow`` と，描画用の座標プラン（純粋な EMU 計算）を返す ``plan_flow`` を
提供する．python-pptx には依存しない（描画は render.py の責務）．

DSL 例::

    direction: lr
    [theme.thmx | テーマ]
    -変換-> [base.pptx | 土台]
    -描画-> [out.pptx | スライド]
    -> [… | ]
    caption: 配色・フォントはテーマ、内容は Markdown
    note(top): テーマと Markdown を入力に pptx を生成
    note(bottom): → テーマを差し替えるだけで見た目が一新できる
"""
from __future__ import annotations

import re
from typing import Literal

from .ir import Flow, FlowNode, FlowEdge


EMU = 914400  # 1 インチ = 914400 EMU

# 受理する値の集合．型付きなので "not in で弾いた残り" が Literal に絞られる
# （検証と型の単一の情報源にもなる）．
_DIRECTIONS: tuple[Literal["lr", "tb"], ...] = ("lr", "tb")
_NODE_KINDS: tuple[Literal["box", "ellipsis"], ...] = ("box", "ellipsis")


def _emu(inch):
    return int(inch * EMU)


# 省略記号として扱うラベル．
_ELLIPSIS = {"…", "..."}

# 設定行（key: value 形式）．
_RE_SETTING = re.compile(r"^(direction|caption|note\(top\)|note\(bottom\))\s*:\s*(.*)$")
# ノード "[ラベル | サブ]" の直後に "{color}" を許す．
_RE_NODE = re.compile(r"\[([^\]]*)\](?:\{([\w-]+)\})?")
# エッジ "->" / "-PR->"．先頭の '-' から '->' まで．
_RE_EDGE = re.compile(r"-(?:([^>]+?)-)?>")


# ---------------------------------------------------------------- パース

def parse_flow(text: str) -> Flow:
    """``` ```flow ``` ブロック本文を Flow（IR）へ変換する．"""
    flow = Flow()
    body_parts = []  # ノード／エッジを含む行（設定行以外）

    for raw in text.split("\n"):
        line = raw.strip()
        if not line:
            continue
        m = _RE_SETTING.match(line)
        if m:
            key, val = m.group(1), m.group(2).strip()
            if key == "direction":
                d = val.lower()
                if d not in _DIRECTIONS:
                    raise ValueError(
                        f"invalid flow direction: {val!r} (lr|tb)")
                flow.direction = d
            elif key == "caption":
                flow.caption = val or None
            elif key == "note(top)":
                flow.note_top = val or None
            elif key == "note(bottom)":
                flow.note_bottom = val or None
            continue
        body_parts.append(line)

    tokens = _tokenize(" ".join(body_parts))
    _build(flow, tokens)
    return flow


def _tokenize(s: str):
    """ノード／エッジのトークン列を返す（出現順）．

    ノード "[…]"／エッジ "->" 以外の文字列はタイポの可能性が高いので
    黙殺せずエラーにする（設定行は parse_flow が先に取り除いている）．
    """
    # ("node", label, color) と ("edge", label) が混在するので要素長は一定でない．
    tokens: list[tuple] = []
    i, n = 0, len(s)
    while i < n:
        c = s[i]
        if c.isspace():
            i += 1
            continue
        if c == "[":
            j = s.find("]", i)
            if j < 0:
                raise ValueError(
                    f"unclosed flow node (missing ']'): {s[i:i + 30]!r}")
            inner = s[i + 1:j]
            i = j + 1
            color = None
            if i < n and s[i] == "{":
                k = s.find("}", i)
                if k < 0:
                    raise ValueError(
                        f"unclosed flow node color (missing '}}'): "
                        f"{s[i:i + 30]!r}")
                color = s[i + 1:k]
                i = k + 1
            tokens.append(("node", inner, color))
            continue
        if c == "-":
            m = _RE_EDGE.match(s, i)
            if m:
                tokens.append(("edge", (m.group(1) or "").strip() or None))
                i = m.end()
                continue
        raise ValueError(
            f"invalid flow syntax near {s[i:i + 30]!r} "
            "(expected '[label | sublabel]' or '->' / '-label->')")
    return tokens


def _build(flow: Flow, tokens) -> None:
    """トークン列（node / edge の交互）から nodes / edges を構築する．"""
    pending_label = None
    have_pending_edge = False
    prev_idx = None

    for tok in tokens:
        if tok[0] == "node":
            node = _make_node(tok[1], tok[2])
            flow.nodes.append(node)
            idx = len(flow.nodes) - 1
            if have_pending_edge and prev_idx is not None:
                flow.edges.append(FlowEdge(src=prev_idx, dst=idx, label=pending_label))
            prev_idx = idx
            have_pending_edge = False
            pending_label = None
        else:  # edge
            have_pending_edge = True
            pending_label = tok[1]


def _make_node(inner: str, color) -> FlowNode:
    parts = inner.split("|", 1)
    label = parts[0].strip()
    sublabel = parts[1].strip() if len(parts) > 1 else None
    if sublabel == "":
        sublabel = None
    kind: Literal["box", "ellipsis"] = (
        "ellipsis" if label in _ELLIPSIS else "box")
    return FlowNode(label=label, sublabel=sublabel, kind=kind, color=color)


# ---------------------------------------------------------------- レイアウト

def plan_flow(flow: Flow, left, top, width, height):
    """Flow を矩形領域 (left, top, width, height) に配置する座標プランを返す．

    戻り値は描画指示の dict（座標はすべて EMU 整数）::

        {
          "boxes":  [(FlowNode, l, t, w, h), ...],   # 角丸四角ノード
          "ellipses": [(label, l, t, w, h), ...],    # "…" ノード
          "arrows": [(x1, y1, x2, y2), ...],         # 矢印
          "labels": [(text, l, t, w, h), ...],       # 矢印ラベル
          "captions": [(text, l, t, w, h, role), ...],  # note_top/caption/note_bottom
        }
    """
    plan: dict[str, list] = {
        "boxes": [], "ellipses": [], "arrows": [], "labels": [], "captions": []}
    nodes = flow.nodes
    if not nodes:
        return plan

    # note_top / note_bottom（地の文）は本文プレースホルダ側で描く（render 側で処理）．
    # ここでは図本体＋キャプションのみを領域内に配置する．
    # 図とキャプションを 1 つのまとまりとして領域中央に置き，キャプションは
    # box の直下に付ける（box とキャプションが離れて間延びしないように）．
    cap_h = _emu(0.5) if flow.caption else 0
    cap_gap = _emu(0.12) if flow.caption else 0

    if flow.direction == "tb":
        bottom = _plan_vertical(plan, flow, left, top, width, height,
                                cap_h + cap_gap)
    else:
        bottom = _plan_horizontal(plan, flow, left, top, width, height,
                                  cap_h + cap_gap)

    if flow.caption:
        cy = bottom + cap_gap
        plan["captions"].append((flow.caption, left, cy, width, cap_h, "caption"))
    return plan


def _plan_horizontal(plan, flow, left, top, width, height, cap_reserve):
    """横並び（lr）に配置し，box 帯の下端 y（キャプション基準）を返す．"""
    nodes = flow.nodes
    n = len(nodes)
    gx = _emu(0.65)
    # 省略記号は「…」1 文字の注記なので box と同じ幅は不要．固定幅で確保し，
    # 残りをすべて box に配分する（省略記号を挟んでも box が縮まないように）．
    ne = sum(1 for node in nodes if node.kind == "ellipsis")
    nb = n - ne
    ew = _emu(0.4)
    if nb:
        bw = (width - (n - 1) * gx - ne * ew) // nb
        bw = max(_emu(1.1), min(_emu(2.4), bw))
    else:
        bw = ew
    bh = min(_emu(1.4), int((height - cap_reserve) * 0.7))
    bh = max(_emu(0.6), bh)
    total = nb * bw + ne * ew + (n - 1) * gx
    startx = left + (width - total) // 2
    # box＋キャプションのまとまりを縦中央に置く．
    group_h = bh + cap_reserve
    by = top + max(0, (height - group_h) // 2)

    centers = []
    bl = startx
    for node in nodes:
        w = ew if node.kind == "ellipsis" else bw
        if node.kind == "ellipsis":
            plan["ellipses"].append((node.label or "…", bl, by, w, bh))
        else:
            plan["boxes"].append((node, bl, by, w, bh))
        centers.append((bl, bl + w // 2, bl + w, by + bh // 2))
        bl += w + gx

    for e in flow.edges:
        if not (0 <= e.src < n and 0 <= e.dst < n):
            continue
        a, b = centers[e.src], centers[e.dst]
        ay = a[3]
        plan["arrows"].append((a[2], ay, b[0], ay))
        if e.label:
            mx = (a[2] + b[0]) // 2
            plan["labels"].append(
                (e.label, mx - _emu(0.5), by - _emu(0.5), _emu(1.0), _emu(0.45)))
    return by + bh


def _plan_vertical(plan, flow, left, top, width, height, cap_reserve):
    """縦並び（tb）に配置し，box 列の下端 y（キャプション基準）を返す．"""
    nodes = flow.nodes
    n = len(nodes)
    gy = _emu(0.35)
    avail = height - cap_reserve
    # 横並びと同じく，省略記号は 1 行分の固定高で確保して残りを box に配分する．
    ne = sum(1 for node in nodes if node.kind == "ellipsis")
    nb = n - ne
    eh = _emu(0.35)
    if nb:
        bh = (avail - (n - 1) * gy - ne * eh) // nb
        bh = max(_emu(0.6), min(_emu(1.2), bh))
    else:
        bh = eh
    bw = min(_emu(3.2), int(width * 0.5))
    bx = left + (width - bw) // 2
    total = nb * bh + ne * eh + (n - 1) * gy
    starty = top + max(0, (avail - total) // 2)

    centers = []
    bt = starty
    for node in nodes:
        h = eh if node.kind == "ellipsis" else bh
        if node.kind == "ellipsis":
            plan["ellipses"].append((node.label or "…", bx, bt, bw, h))
        else:
            plan["boxes"].append((node, bx, bt, bw, h))
        centers.append((bx + bw // 2, bt, bt + h))
        bt += h + gy

    for e in flow.edges:
        if not (0 <= e.src < n and 0 <= e.dst < n):
            continue
        a, b = centers[e.src], centers[e.dst]
        cx = a[0]
        plan["arrows"].append((cx, a[2], cx, b[1]))
        if e.label:
            my = (a[2] + b[1]) // 2
            plan["labels"].append(
                (e.label, cx + _emu(0.2), my - _emu(0.22), _emu(1.2), _emu(0.45)))
    return starty + total
