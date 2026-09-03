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
from .layout import (
    EMU, PlacedArrow, PlacedLine, PlacedText, Rect, emu as _emu,
)


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
# ノード名 "[#pc PC | …]"．**"#" で書く**——"[id: ラベル]" だと
# "[HTTP: ハイパーテキスト転送プロトコル]" を名前付きと読んでしまう（Issue #109）．
_RE_NODE_ID = re.compile(r"^#([A-Za-z_][\w-]*)\s+(.*)$", re.S)
# エッジ行でノードを指す名前（トークン列では _RefTok）．
_RE_REF = re.compile(r"[A-Za-z_][\w-]*")
# 段区切り．この行だけで次の段へ移る．
_RE_ROW_BREAK = re.compile(r"^-{2,}$")


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


@dataclass(frozen=True)
class _RefTok:
    """既に置いたノードを名前で指す（``pc -> srv`` の ``pc``）．

    ``_NodeTok`` と別にしてあるのは**置くのか指すのかが違う**から．同じ型に
    詰めると ``_build`` が「今つくるのか探すのか」を憶えていないと間違える．
    """
    name: str


# requires-python が 3.11 なので実行時に評価される ``|`` を使える．
_Token = _NodeTok | _EdgeTok | _RefTok


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

    # **段ごとにトークン化する**（従来は全行を空白で連結していた）．
    # 段区切りは行の構造そのものなので、連結してしまうと復元できない．
    rows_of_tokens: list[list[_Token]] = [[]]
    # 図の中の段階（Issue #125）．``@step`` までに置いたノードの**数**を覚える．
    # 行の位置ではなく数で持つのは、parser が Flow.upto(n) で切るため．
    step_at: list[int] = []
    seen_nodes = 0
    for line in body_parts:
        if line == "@step":
            step_at.append(seen_nodes)
            continue
        if _RE_ROW_BREAK.match(line):
            rows_of_tokens.append([])
            continue
        toks = _tokenize(line)
        seen_nodes += sum(1 for t in toks if isinstance(t, _NodeTok))
        rows_of_tokens[-1].extend(toks)
    has_break = len(rows_of_tokens) > 1
    _build(flow, rows_of_tokens, has_break)
    flow.steps = step_at
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
        mr = _RE_REF.match(s, i)
        if mr:
            tokens.append(_RefTok(mr.group(0)))
            i = mr.end()
            continue
        raise ValueError(
            f"invalid flow syntax near {s[i:i + 30]!r} "
            "(expected '[label | sublabel]' or '->' / '-label->')")
    return tokens


def _build(flow: Flow, rows: list[list[_Token]], has_break: bool) -> None:
    """段ごとのトークン列から nodes / edges / rows を構築する．

    エッジは 2 つのノードに挟まれて初めて線になるので，``->`` を読んだ時点では
    まだ引けない——次のノード（か名前）が来るまで「保留中の矢印」として持ち越す．
    相手が現れないまま終われば（``-> [A]`` の先頭や ``[A] ->`` の末尾），
    その矢印は捨てる．

    **保留は段をまたがない．** 段の終わりで持ち越すと、書いていない縦の線が
    生えることになる（段をまたぐ線は名前で明示的に書く）．

    ``has_break`` は原稿に ``--`` があったか．無ければ ``flow.rows`` は空のまま
    にする——「段の指定なし＝一列」を保ち、従来の原稿の配置を変えないため．
    """
    names: dict[str, int] = {}

    def resolve(tok: _NodeTok | _RefTok) -> int:
        """トークンをノード index にする（``_NodeTok`` は置き、``_RefTok`` は探す）．"""
        if isinstance(tok, _RefTok):
            if tok.name not in names:
                known = ", ".join(sorted(names)) or "(none)"
                raise ValueError(
                    f"unknown flow node name: {tok.name!r} (known: {known})")
            return names[tok.name]
        node = _make_node(tok.label, tok.color)
        if node.node_id is not None:
            if node.node_id in names:
                raise ValueError(
                    f"duplicate flow node name: {node.node_id!r}")
            names[node.node_id] = len(flow.nodes)
        flow.nodes.append(node)
        return len(flow.nodes) - 1

    for row in rows:
        placed: list[int] = []          # この段に**新しく置いた**ノード
        pending: _EdgeTok | None = None
        prev_idx: int | None = None
        for tok in row:
            if isinstance(tok, _EdgeTok):
                # ``->`` が続いた場合は上書き（覚えておけるのは直前の 1 本だけ）．
                pending = tok
                continue
            new_node = isinstance(tok, _NodeTok)
            idx = resolve(tok)
            if new_node:
                placed.append(idx)
            if pending is not None and prev_idx is not None:
                flow.edges.append(
                    FlowEdge(src=prev_idx, dst=idx, label=pending.label))
            prev_idx = idx
            pending = None
        if has_break and placed:
            flow.rows.append(placed)


def _make_node(inner: str, color: str | None) -> FlowNode:
    node_id: str | None = None
    mid = _RE_NODE_ID.match(inner.strip())
    if mid:
        node_id, inner = mid.group(1), mid.group(2)
    parts = inner.split("|", 1)
    label = parts[0].strip()
    sublabel = parts[1].strip() if len(parts) > 1 else None
    if sublabel == "":
        sublabel = None
    kind: Literal["box", "ellipsis"] = (
        "ellipsis" if label in _ELLIPSIS else "box")
    return FlowNode(label=label, sublabel=sublabel, kind=kind, color=color,
                    node_id=node_id)


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
    # 隣り合わないノードを結ぶ細い矢印（Issue #109）．``arrows`` の塗り矢印は
    # **すき間を埋める**ための形なので、離れた 2 点を結ぶと box に食い込む．
    lines: list[PlacedLine] = field(default_factory=list)
    labels: list[PlacedText] = field(default_factory=list)
    captions: list[PlacedText] = field(default_factory=list)


# ---------------------------------------------------------------- レイアウト

# 矢印ラベルの枠を見積もるときの既定の文字サイズ．
#
# **呼び出し側が実寸を渡す**（``plan_flow(label_pt=…)``）．render は自分が
# どのサイズで描くかを知っているので、渡してもらえば枠は実寸で決まる．
#
# 以前はここを 16pt 固定にして「実寸と合っている必要はない、折り返さないから
# 字は割れない」と説明していた．**横はそのとおりだが縦は成り立たない**——
# 枠は ``anchor=MIDDLE`` なので、大きい字は枠の上下へ等しくはみ出し、
# 下側が box に乗る（Issue #178．cn2026-12 p.42 で「暗号化」が 9pt 食い込んでいた）．
_LABEL_PT = 16.0

# 行送り．枠の高さは「文字サイズ × これ」に上下のアキを足したもの．
_LABEL_LINE = 1.2
_LABEL_PAD = _emu(0.08)


def _label_height(label_pt: float = _LABEL_PT) -> int:
    """矢印ラベルの枠の高さ（EMU）．**描くサイズで決まる**（Issue #178）．"""
    return int(label_pt * _LABEL_LINE * 12700) + _LABEL_PAD


def _label_width(text: str, label_pt: float = _LABEL_PT) -> int:
    """矢印ラベルの想定幅（EMU）．全角は半角の2倍で数える．

    固定幅だと**中心がずれる**（長いラベルほど左に寄って矢印から離れる）．
    Issue #111 では折り返しと重なって "NAMEPREP" が "NAM / EPRE / P" と
    3 行に割れていた．折り返しは render 側で止め，ここでは中心を合わせる．
    """
    # 0x2E80 より上を全角とみなす．CJK の記号・かな・漢字（U+3000〜）だけでなく、
    # 全角英数と全角記号（U+FF01〜）も上側に入るので、実用上これで足りる
    # （"（" U+FF08 も "Ａ" U+FF21 も 2 文字ぶんとして数えられることを確認済み）．
    units = sum(2 if ord(c) > 0x2E80 else 1 for c in text)
    return int(units * label_pt * 12700 * 0.55) + _emu(0.12)


def _label_rect(text: str, center_x: int, bottom_y: int,
                label_pt: float = _LABEL_PT) -> Rect:
    """矢印ラベルを box の上へ、文字が入る幅と高さで置く．"""
    w = max(_emu(0.5), _label_width(text, label_pt))
    h = _label_height(label_pt)
    return Rect(center_x - w // 2, bottom_y - h - _emu(0.04), w, h)


def plan_flow(flow: Flow, left: int, top: int, width: int,
              height: int, label_pt: float = _LABEL_PT) -> FlowPlan:
    """Flow を矩形領域 (left, top, width, height) に配置する．

    ``label_pt`` は矢印ラベルを**実際に描く**文字サイズ．枠の幅と高さをこれで
    決めるので、render は自分が使うサイズをそのまま渡す（Issue #178）．
    """
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

    if flow.rows:
        bottom = _plan_grid(plan, flow, left, top, width, height,
                            cap_h + cap_gap, label_pt)
    elif flow.direction == "tb":
        bottom = _plan_vertical(plan, flow, left, top, width, height,
                                cap_h + cap_gap, label_pt)
    else:
        bottom = _plan_horizontal(plan, flow, left, top, width, height,
                                  cap_h + cap_gap, label_pt)

    if flow.caption:
        cy = bottom + cap_gap
        plan.captions.append(
            PlacedText(flow.caption, Rect(left, cy, width, cap_h)))
    return plan


def _plan_grid(plan: FlowPlan, flow: Flow, left: int, top: int,
               width: int, height: int, reserve: int,
               label_pt: float = _LABEL_PT) -> int:
    """段（``--`` 区切り）を持つ図を格子に置く（Issue #109）．

    段は上から下へ、段の中は左から右。**列は全段で揃える**——段ごとに幅を
    変えると、同じ列にあるはずのノードがずれて「格子」に見えない。
    列数はいちばん要素の多い段に合わせる。

    エッジは 2 通りに描き分ける:

    - **同じ段で隣り合う** → ``arrows``（塗り矢印）．すき間にちょうど収まる形で、
      一列のフローと同じ見た目になる
    - **それ以外**（段をまたぐ・飛び越す） → ``lines``（細い矢印）．
      塗り矢印は「すき間を埋める」ための形なので、離れた 2 点を結ぶと box に食い込む

    戻り値は図の下端（キャプションを置く位置）．
    """
    rows = flow.rows
    ncol = max(len(r) for r in rows)
    nrow = len(rows)
    gap_x, gap_y = _emu(0.5), _emu(0.45)
    avail_w = width - gap_x * (ncol - 1)
    avail_h = height - reserve - gap_y * (nrow - 1)
    bw = max(_emu(0.9), min(_emu(2.4), avail_w // ncol))
    bh = max(_emu(0.5), min(_emu(1.2), avail_h // nrow))

    grid_w = bw * ncol + gap_x * (ncol - 1)
    grid_h = bh * nrow + gap_y * (nrow - 1)
    x0 = left + max(0, (width - grid_w) // 2)
    y0 = top + max(0, (height - reserve - grid_h) // 2)

    rects: dict[int, Rect] = {}
    for r, row in enumerate(rows):
        y = y0 + r * (bh + gap_y)
        for c, idx in enumerate(row):
            rect = Rect(x0 + c * (bw + gap_x), y, bw, bh)
            rects[idx] = rect
            node = flow.nodes[idx]
            if node.kind == "ellipsis":
                plan.ellipses.append(PlacedText(node.label, rect))
            else:
                plan.boxes.append(PlacedNode(node, rect))

    # 同じ段で隣り合うか（塗り矢印にしてよいか）を引けるようにしておく．
    pos = {idx: (r, c) for r, row in enumerate(rows) for c, idx in enumerate(row)}
    for e in flow.edges:
        a, b = rects.get(e.src), rects.get(e.dst)
        if a is None or b is None:
            continue
        pa, pb = pos.get(e.src), pos.get(e.dst)
        if pa is None or pb is None:
            # **構造上ここへは来ない**——``_build`` はノードを置いたら必ず
            # その段に入れ，段は必ず ``flow.rows`` へ入るので，どのノードも
            # ちょうど 1 つの段に属する．黙って線を落とすのは避けたい挙動なので，
            # 万一この不変条件が崩れたときに落ちないための保険として残す．
            continue
        adjacent = pa[0] == pb[0] and abs(pa[1] - pb[1]) == 1
        if adjacent:
            x1 = a.right if pb[1] > pa[1] else a.left
            x2 = b.left if pb[1] > pa[1] else b.right
            plan.arrows.append(PlacedArrow(x1, a.center_y, x2, b.center_y))
            if e.label:
                plan.labels.append(PlacedText(
                    e.label, _label_rect(e.label, (x1 + x2) // 2, a.top,
                                         label_pt)))
        else:
            x1, y1, x2, y2 = _edge_points(a, b)
            plan.lines.append(PlacedLine(x1, y1, x2, y2))
            if e.label:
                plan.labels.append(PlacedText(
                    e.label, _label_rect(e.label, (x1 + x2) // 2,
                                         (y1 + y2) // 2
                                         + _label_height(label_pt) // 2,
                                         label_pt)))
    return y0 + grid_h


def _edge_points(a: Rect, b: Rect) -> tuple[int, int, int, int]:
    """2 つの矩形を結ぶ線分の端点（互いに向いた辺の中点）を返す．

    中心どうしを結ぶと線が box の中を通ってしまうので、**縦横どちらに離れて
    いるかで辺を選ぶ**．斜めに離れている場合は離れ方の大きいほうを採る．
    """
    dx = b.center_x - a.center_x
    dy = b.center_y - a.center_y
    if abs(dy) >= abs(dx):                      # 主に上下に離れている
        if dy >= 0:
            return a.center_x, a.bottom, b.center_x, b.top
        return a.center_x, a.top, b.center_x, b.bottom
    if dx >= 0:                                 # 主に左右に離れている
        return a.right, a.center_y, b.left, b.center_y
    return a.left, a.center_y, b.right, b.center_y


def _plan_horizontal(plan: FlowPlan, flow: Flow, left: int, top: int,
                     width: int, height: int, cap_reserve: int,
                     label_pt: float = _LABEL_PT) -> int:
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
            plan.labels.append(PlacedText(
                e.label, _label_rect(e.label, mx, by, label_pt)))
    return by + bh


def _plan_vertical(plan: FlowPlan, flow: Flow, left: int, top: int,
                   width: int, height: int, cap_reserve: int,
                   label_pt: float = _LABEL_PT) -> int:
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
            # 縦並びではラベルを矢印の**右横**に置く（上下は box で埋まっている）．
            # **左寄せ**にするのは、枠幅が ``_label_width`` の見積もりだから
            # （Issue #176）——中央揃えだと見積もりを超えたぶんが左右へ等しく
            # はみ出し、左側が矢印に掛かって先頭の字が読めなくなる。
            my = (a.bottom + b.top) // 2
            r = _label_rect(e.label, cx, my + _label_height(label_pt) // 2,
                            label_pt)
            plan.labels.append(PlacedText(
                e.label, Rect(cx + _emu(0.18), r.top, r.width, r.height),
                align="left"))
    return starty + total
