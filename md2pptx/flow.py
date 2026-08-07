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
from dataclasses import dataclass, field
from typing import Literal

from .ir import Flow, FlowNode, FlowEdge
from .layout import EMU, PlacedArrow, PlacedText, Rect, emu as _emu


# 受理する値の集合．型付きなので "not in で弾いた残り" が Literal に絞られる
# （検証と型の単一の情報源にもなる）．
_DIRECTIONS: tuple[Literal["lr", "tb"], ...] = ("lr", "tb")
_NODE_KINDS: tuple[Literal["box", "ellipsis"], ...] = ("box", "ellipsis")


# 省略記号として扱うラベル．
_ELLIPSIS = {"…", "..."}

# 設定行（key: value 形式）．
_RE_SETTING = re.compile(r"^(direction|caption|note\(top\)|note\(bottom\))\s*:\s*(.*)$")
# ノード "[ラベル | サブ]" の直後に "{color}" を許す．
_RE_NODE = re.compile(r"\[([^\]]*)\](?:\{([\w-]+)\})?")
# エッジ "->" / "-PR->"．先頭の '-' から '->' まで．
_RE_EDGE = re.compile(r"-(?:([^>]+?)-)?>")


# ---------------------------------------------------------------- トークン

# 字句解析の途中結果．**ir.FlowNode / FlowEdge とは別物**で，こちらは構文解析の
# 作業用（``_make_node`` がこれを IR へ変換する）．混同を避けるため私有名にしてある．
#
# 種類ごとに別の型にしているのは，持っている情報が違うから．ノードは色を持ち，
# エッジは持たない——これを 1 つのタプルに詰めると，読む側が「今どちらなのか」を
# 憶えていないと添字を間違える（それを型で言えないのが Issue #34）．

@dataclass(frozen=True)
class _NodeTok:
    """``[ラベル | サブ]`` ／ ``[…]{色}``．``label`` は ``|`` で割る前の生の中身．"""
    label: str
    color: str | None = None


@dataclass(frozen=True)
class _EdgeTok:
    """``->`` ／ ``-ラベル->``．``label`` が None なら文字の無い矢印．"""
    label: str | None = None


# requires-python が 3.11 なので実行時に評価される ``|`` を使える．
_Token = _NodeTok | _EdgeTok


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


def _tokenize(s: str) -> list[_Token]:
    """ノード／エッジのトークン列を返す（出現順）．

    ノード "[…]"／エッジ "->" 以外の文字列はタイポの可能性が高いので
    黙殺せずエラーにする（設定行は parse_flow が先に取り除いている）．
    """
    tokens: list[_Token] = []
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
            color: str | None = None
            if i < n and s[i] == "{":
                k = s.find("}", i)
                if k < 0:
                    raise ValueError(
                        f"unclosed flow node color (missing '}}'): "
                        f"{s[i:i + 30]!r}")
                color = s[i + 1:k]
                i = k + 1
            tokens.append(_NodeTok(inner, color))
            continue
        if c == "-":
            m = _RE_EDGE.match(s, i)
            if m:
                tokens.append(_EdgeTok((m.group(1) or "").strip() or None))
                i = m.end()
                continue
        raise ValueError(
            f"invalid flow syntax near {s[i:i + 30]!r} "
            "(expected '[label | sublabel]' or '->' / '-label->')")
    return tokens


def _build(flow: Flow, tokens: list[_Token]) -> None:
    """トークン列（node / edge の交互）から nodes / edges を構築する．

    エッジは 2 つのノードに挟まれて初めて線になるので，``->`` を読んだ時点では
    まだ引けない——次のノードが来るまで「保留中の矢印」として持ち越す．
    相手が現れないまま終われば（``-> [A]`` の先頭や ``[A] ->`` の末尾），
    その矢印は捨てる．
    """
    pending: _EdgeTok | None = None
    prev_idx: int | None = None

    for tok in tokens:
        if isinstance(tok, _NodeTok):
            flow.nodes.append(_make_node(tok.label, tok.color))
            idx = len(flow.nodes) - 1
            if pending is not None and prev_idx is not None:
                flow.edges.append(
                    FlowEdge(src=prev_idx, dst=idx, label=pending.label))
            prev_idx = idx
            pending = None
        else:
            # ``->`` が続いた場合は上書き（覚えておけるのは直前の 1 本だけ）．
            pending = tok


def _make_node(inner: str, color: str | None) -> FlowNode:
    parts = inner.split("|", 1)
    label = parts[0].strip()
    sublabel = parts[1].strip() if len(parts) > 1 else None
    if sublabel == "":
        sublabel = None
    kind: Literal["box", "ellipsis"] = (
        "ellipsis" if label in _ELLIPSIS else "box")
    return FlowNode(label=label, sublabel=sublabel, kind=kind, color=color)


# ---------------------------------------------------------------- 配置プラン

# 描画指示（座標はすべて EMU 整数）．render はこれを読んで図形を置くだけで，
# 位置の計算はここで終わっている．
#
@dataclass(frozen=True)
class PlacedNode:
    """角丸四角ノード．色・サブラベルを持つので ``node`` ごと渡す．

    **``layout.py`` ではなくここに置く**——``FlowNode`` を抱えており flow 固有だから．
    共通の入れ物に 1 つの DSL の型を持ち込むと、次の DSL が再利用できずに詰む．
    """
    node: FlowNode
    rect: Rect


@dataclass
class FlowPlan:
    """図ひとつぶんの配置．

    ``note_top`` / ``note_bottom`` はここに入らない——地の文なので本文
    プレースホルダへ入れるのが render の仕事で，図の座標を持たない．

    **ここだけ ``frozen`` にしていない．** 中の要素（``Rect`` や ``Placed*``）は
    作ったら変わらないが，この入れ物は ``_plan_horizontal`` / ``_plan_vertical``
    が置きながら追記していくため．
    """
    boxes: list[PlacedNode] = field(default_factory=list)
    ellipses: list[PlacedText] = field(default_factory=list)
    arrows: list[PlacedArrow] = field(default_factory=list)
    labels: list[PlacedText] = field(default_factory=list)
    captions: list[PlacedText] = field(default_factory=list)


# ---------------------------------------------------------------- レイアウト

def plan_flow(flow: Flow, left: int, top: int, width: int,
              height: int) -> FlowPlan:
    """Flow を矩形領域 (left, top, width, height) に配置する．"""
    plan = FlowPlan()
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
        plan.captions.append(
            PlacedText(flow.caption, Rect(left, cy, width, cap_h)))
    return plan


def _plan_horizontal(plan: FlowPlan, flow: Flow, left: int, top: int,
                     width: int, height: int, cap_reserve: int) -> int:
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

    # 置いた矩形はエッジを引くときに端と中心が要るので，そのまま取っておく
    # （ノードの並び順で引ける）．
    rects = []
    bl = startx
    for node in nodes:
        w = ew if node.kind == "ellipsis" else bw
        rect = Rect(bl, by, w, bh)
        if node.kind == "ellipsis":
            plan.ellipses.append(PlacedText(node.label or "…", rect))
        else:
            plan.boxes.append(PlacedNode(node, rect))
        rects.append(rect)
        bl += w + gx

    for e in flow.edges:
        if not (0 <= e.src < n and 0 <= e.dst < n):
            continue
        a, b = rects[e.src], rects[e.dst]
        ay = a.center_y
        plan.arrows.append(PlacedArrow(a.right, ay, b.left, ay))
        if e.label:
            mx = (a.right + b.left) // 2
            plan.labels.append(PlacedText(e.label, Rect(
                mx - _emu(0.5), by - _emu(0.5), _emu(1.0), _emu(0.45))))
    return by + bh


def _plan_vertical(plan: FlowPlan, flow: Flow, left: int, top: int,
                   width: int, height: int, cap_reserve: int) -> int:
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

    rects = []
    bt = starty
    for node in nodes:
        h = eh if node.kind == "ellipsis" else bh
        rect = Rect(bx, bt, bw, h)
        if node.kind == "ellipsis":
            plan.ellipses.append(PlacedText(node.label or "…", rect))
        else:
            plan.boxes.append(PlacedNode(node, rect))
        rects.append(rect)
        bt += h + gy

    for e in flow.edges:
        if not (0 <= e.src < n and 0 <= e.dst < n):
            continue
        a, b = rects[e.src], rects[e.dst]
        cx = a.center_x
        plan.arrows.append(PlacedArrow(cx, a.bottom, cx, b.top))
        if e.label:
            my = (a.bottom + b.top) // 2
            plan.labels.append(PlacedText(e.label, Rect(
                cx + _emu(0.2), my - _emu(0.22), _emu(1.2), _emu(0.45))))
    return starty + total
