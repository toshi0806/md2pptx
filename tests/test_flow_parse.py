#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""フロー図 DSL の解釈を固定するテスト（Issue #34 の足場）．

``flow.py`` の中間表現をタプルから dataclass へ移すにあたり，**移す前に**書いた．
中身を入れ替える変更なので，見るべきは「入力に対して同じ Flow が出るか」であって
トークンの持ち方ではない——だからこのファイルは ``parse_flow`` / ``plan_flow`` という
**外から見える 2 つの関数だけ**を叩く．``_tokenize`` や ``_NodeTok`` には触れない
（触ると，構造を変えるたびにテストごと書き換えることになり，守るものが無くなる）．

DSL の仕様として DESIGN.md に書いてあるのは主要な形だけで，**端の挙動は
どこにも書かれていない**．書かれていないものは変わっても気づけないので，
現状の挙動をここに写しておく（「こうあるべき」ではなく「いまこうである」）．
"""
from __future__ import annotations

import re

import pytest

from md2pptx import flow as F


# ------------------------------------------------------------ 基本の形

def test_it_reads_the_pipeline_from_example_md():
    """example.md の「生成パイプライン」——DSL の要素が一通り出てくる唯一の実例．"""
    f = F.parse_flow(
        "direction: lr\n"
        "note(top): テーマと Markdown を入力に pptx を生成\n"
        "[theme.thmx | テーマ]\n"
        "-変換-> [base.pptx | 土台]\n"
        "-描画-> [out.pptx | スライド]\n"
        "caption: 配色・フォントはテーマ、内容は Markdown\n"
        "note(bottom): → テーマを差し替えるだけで見た目が一新できる")

    assert f.direction == "lr"
    assert f.caption == "配色・フォントはテーマ、内容は Markdown"
    assert f.note_top == "テーマと Markdown を入力に pptx を生成"
    assert f.note_bottom == "→ テーマを差し替えるだけで見た目が一新できる"

    assert [n.label for n in f.nodes] == ["theme.thmx", "base.pptx", "out.pptx"]
    assert [n.sublabel for n in f.nodes] == ["テーマ", "土台", "スライド"]
    assert [(e.src, e.dst, e.label) for e in f.edges] == [
        (0, 1, "変換"), (1, 2, "描画")]


def test_nodes_and_edges_may_share_a_line():
    """1 行に詰めても行ごとに分けても同じ——行は連結されてから解釈される．"""
    packed = F.parse_flow("[A] -> [B] -> [C]")
    spread = F.parse_flow("[A]\n-> [B]\n-> [C]")

    assert [n.label for n in packed.nodes] == [n.label for n in spread.nodes]
    assert [(e.src, e.dst) for e in packed.edges] == [(0, 1), (1, 2)]


def test_an_edge_without_a_label_has_none():
    """``->`` と ``-ラベル->`` の違いはラベルの有無だけ（空文字ではなく None）．"""
    f = F.parse_flow("[A] -> [B] -x-> [C]")

    assert [e.label for e in f.edges] == [None, "x"]


@pytest.mark.parametrize("src, label, sublabel", [
    ("[A | sub]", "A", "sub"),
    ("[A]", "A", None),
    ("[A|]", "A", None),              # 区切りだけ書いても空のサブラベルは持たない
    ("[ | ]", "", None),
    ("[  ]", "", None),
    ("[|B]", "", "B"),                # ラベルが空でもサブラベルは生きる
    ("[A | b | c]", "A", "b | c"),    # 区切りは最初の 1 個だけ
])
def test_how_a_node_label_splits(src, label, sublabel):
    """``|`` の扱い．空白は落とし，空になったサブラベルは None にする．"""
    node = F.parse_flow(src).nodes[0]

    assert (node.label, node.sublabel) == (label, sublabel)


def test_a_color_can_follow_a_node():
    """``{...}`` はノードの色．付いていないノードは None のまま．"""
    f = F.parse_flow("[A]{accent2} -> [B]")

    assert [n.color for n in f.nodes] == ["accent2", None]


def test_a_color_does_not_need_a_space_before_the_next_node():
    """``[A]{c}[B]`` のように詰めても 2 ノードに割れる（色の閉じで区切れる）．"""
    f = F.parse_flow("[A]{c1}[B]")

    assert [(n.label, n.color) for n in f.nodes] == [("A", "c1"), ("B", None)]


@pytest.mark.parametrize("label", ["…", "..."])
def test_an_ellipsis_node_is_a_kind_of_its_own(label):
    """``…`` / ``...`` は「途中を省いた」印で，ふつうの箱とは描き方が違う．"""
    f = F.parse_flow(f"[A] -> [{label}] -> [B]")

    assert [n.kind for n in f.nodes] == ["box", "ellipsis", "box"]


# ------------------------------------------------------- 端の挙動（記録）

def test_an_edge_with_nothing_on_one_side_is_dropped():
    """繋ぐ先が無いエッジは黙って消える．

    ``-> [A] ->`` は前後に相手がいないので線が引けない．**エラーにはしない**——
    書きかけの原稿でよく通る形で，ここで止めると編集しながらのプレビューが使えない．
    """
    f = F.parse_flow("-> [A] ->")

    assert [n.label for n in f.nodes] == ["A"]
    assert f.edges == []


def test_nodes_side_by_side_are_not_connected():
    """矢印を書かずに並べたノードは繋がらない．

    ``[A] -> [B] [C]`` で B–C 間に線が出ないこと——保留中の矢印を使い切ったら
    忘れる，という一点で決まっている．忘れ損ねると，1 本書いた矢印が以降の
    ノード全部に伝染して，**書いていない線が図に増える**．
    """
    f = F.parse_flow("[A] -x-> [B] [C]")

    assert [n.label for n in f.nodes] == ["A", "B", "C"]
    assert [(e.src, e.dst, e.label) for e in f.edges] == [(0, 1, "x")]


def test_only_the_last_of_consecutive_edges_survives():
    """``->`` が続いたら線は 1 本，ラベルは**最後のものが残る**．

    ``-先-> -後-> [B]`` で「先」が消えるのは意図した設計ではなく，
    次のノードを待つ間ラベルを 1 つしか覚えていないことの帰結．変えるなら
    それは仕様変更であって，リファクタリングで静かに変わってよいものではない．
    """
    f = F.parse_flow("[A] -先-> -後-> [B]")

    assert [(e.src, e.dst, e.label) for e in f.edges] == [(0, 1, "後")]


def test_an_empty_diagram_is_not_an_error():
    """ノードが 0 個でも通る（``` ```flow ``` を開いた直後の状態）．"""
    assert F.parse_flow("").nodes == []
    assert F.parse_flow("->").nodes == []


def test_settings_are_taken_from_anywhere_in_the_block():
    """設定行は図の途中に挟んでもよい（行の位置で意味が変わらない）．"""
    f = F.parse_flow("[A]\ndirection: tb\n-> [B]\ncaption: あとがき")

    assert (f.direction, f.caption) == ("tb", "あとがき")
    assert [(e.src, e.dst) for e in f.edges] == [(0, 1)]


@pytest.mark.parametrize("key, attr", [
    ("caption", "caption"), ("note(top)", "note_top"),
    ("note(bottom)", "note_bottom"),
])
def test_a_setting_with_an_empty_value_stays_none(key, attr):
    """``caption:`` と書いて中身が無いのは「指定しなかった」と同じ扱い．"""
    assert getattr(F.parse_flow(f"[A]\n{key}:"), attr) is None


# --------------------------------------------------------- 誤りの伝え方

@pytest.mark.parametrize("src, message", [
    ("[unclosed",     "unclosed flow node (missing ']')"),
    ("[A]{unclosed",  "unclosed flow node color (missing '}')"),
    # 裸の語は**ノード名の参照**になった（Issue #109）．知らない名前は
    # 「文法が違う」ではなく「その名前が無い」と言うほうが直しやすい．
    ("junk",          "unknown flow node name"),
    ("[A] ~ [B]",     "invalid flow syntax"),
    ("[A] - [B]",     "invalid flow syntax"),   # '-' だけでは矢印にならない
    ("[A] -x- [B]",   "invalid flow syntax"),   # '>' で閉じていない
    ("direction: ne", "invalid flow direction"),
])
def test_it_refuses_rather_than_ignores(src, message):
    """読めない字はタイポの可能性が高いので黙って捨てない．

    捨てると，矢印を書いたつもりの図が線の無い図として出てきて，原因が分からない．
    """
    with pytest.raises(ValueError, match=re.escape(message)):
        F.parse_flow(src)


def test_the_error_shows_where_it_gave_up():
    """メッセージに問題の箇所を入れる（どこを直すかが分からないと意味がない）．"""
    with pytest.raises(ValueError, match="~ここ"):
        F.parse_flow("[A] ~ここ~ [B]")


# --------------------------------------------------------- 矩形（#71）

def test_a_rect_knows_its_own_edges_and_centre():
    """``Rect`` の端と中心．**実測値を直に書く**．

    他のテストが ``a.right`` のような属性で期待値を組み立てているので，
    ここだけは属性を使わずに数で押さえる——さもないと実装が壊れたとき
    期待値も同じように壊れて，両辺が一致してしまう（実際にそうなった）．
    """
    r = F.Rect(left=100, top=200, width=30, height=41)

    assert (r.left, r.top, r.width, r.height) == (100, 200, 30, 41)
    assert (r.right, r.bottom) == (130, 241)
    assert (r.center_x, r.center_y) == (115, 220)   # 端数は切り捨て


# --------------------------------------------------------- 配置（#26）

def _slots(plan, direction):
    """箱と省略記号を並び順に戻し，(省略記号か, 主軸方向の大きさ) の列を返す．"""
    rects = ([(b.rect, False) for b in plan.boxes]
             + [(e.rect, True) for e in plan.ellipses])
    if direction == "lr":
        rects.sort(key=lambda r: r[0].left)
        return [(is_ellipsis, r.width) for r, is_ellipsis in rects]
    rects.sort(key=lambda r: r[0].top)
    return [(is_ellipsis, r.height) for r, is_ellipsis in rects]


@pytest.mark.parametrize("direction", ["lr", "tb"])
def test_an_ellipsis_takes_a_narrower_slot_than_a_box(direction):
    """省略記号は固定の小さい枠を取り，浮いた分を箱が使う（Issue #26）．

    「…」は 1 文字の注記なので箱と同じ幅は要らない．等分に配ると，省略を 1 つ
    挟んだだけで箱が 1 つ増えたのと同じだけ痩せる——書いている側から見ると
    理由の分からない縮み方になる．

    「箱がまったく縮まない」わけではない（枠 1 つ分と間隔は増える）．
    確かめたいのは**省略記号のほうが箱より狭く，その差だけ箱が得をする**こと．
    """
    src = f"direction: {direction}\n" + "{}"
    area = (0, 0, F.EMU * 6, F.EMU * 4)   # 上下限に張り付かない大きさを選ぶ
    with_box = _slots(
        F.plan_flow(F.parse_flow(src.format("[A] -> [X] -> [B]")), *area),
        direction)
    with_gap = _slots(
        F.plan_flow(F.parse_flow(src.format("[A] -> [… | ] -> [B]")), *area),
        direction)

    assert [s[0] for s in with_gap] == [False, True, False], "真ん中が省略記号"
    assert with_gap[1][1] < with_gap[0][1], "省略記号の枠は箱より狭い"
    assert with_gap[0][1] > with_box[0][1], "その分だけ箱が広く取れている"


@pytest.mark.parametrize("direction", ["lr", "tb"])
def test_it_draws_one_arrow_per_edge(direction):
    """エッジの本数だけ矢印が出る．ラベル付きのものにだけ文字が付く．"""
    plan = F.plan_flow(
        F.parse_flow(f"direction: {direction}\n[A] -> [B] -付き-> [C]"),
        0, 0, F.EMU * 8, F.EMU * 4)

    assert len(plan.arrows) == 2
    assert [lb.text for lb in plan.labels] == ["付き"]


@pytest.mark.parametrize("direction", ["lr", "tb"])
def test_an_arrow_runs_between_the_two_nodes(direction):
    """矢印は前のノードの端から次のノードの端まで（中を突き抜けない）．

    lr なら右端→左端で高さは中心，tb なら下端→上端で左右は中心．
    ここを取り違えても座標は同じ整数なので，出力を見るまで分からない．
    """
    plan = F.plan_flow(F.parse_flow(f"direction: {direction}\n[A] -> [B]"),
                       0, 0, F.EMU * 8, F.EMU * 4)
    (a, b), (arrow,) = [x.rect for x in plan.boxes], plan.arrows

    # 期待値は ``right`` / ``center_y`` といった**導出プロパティを経由せず**，
    # 素の値（left / top / width / height）の算術で組む．プロパティを使うと，
    # それが壊れたとき期待値まで同じように壊れて一致してしまう（実際そうなった）．
    # 素の値そのものは test_a_rect_knows_its_own_edges_and_centre が実測値で
    # 押さえているので，ここはその上に乗る．
    if direction == "lr":
        assert (arrow.x1, arrow.x2) == (a.left + a.width, b.left)
        assert arrow.y1 == arrow.y2 == a.top + a.height // 2
    else:
        assert (arrow.y1, arrow.y2) == (a.top + a.height, b.top)
        assert arrow.x1 == arrow.x2 == a.left + a.width // 2


def test_a_caption_sits_below_the_boxes():
    """キャプションは図の下（離れて間延びしないよう箱の直下に付ける）．"""
    plan = F.plan_flow(F.parse_flow("[A] -> [B]\ncaption: 説明"),
                       0, 0, F.EMU * 8, F.EMU * 4)

    (caption,) = plan.captions
    box_bottom = max(b.rect.bottom for b in plan.boxes)

    assert caption.text == "説明"
    assert caption.rect.top >= box_bottom


def test_notes_are_not_drawn_here():
    """note(top) / note(bottom) は地の文なので，図の座標プランには出てこない．

    本文プレースホルダへ入れるのは render の仕事（DESIGN.md の分担）．
    ここで箱を作ってしまうと二重に描かれる．
    """
    plan = F.plan_flow(
        F.parse_flow("[A]\nnote(top): 上\nnote(bottom): 下\ncaption: 中"),
        0, 0, F.EMU * 8, F.EMU * 4)

    assert [c.text for c in plan.captions] == ["中"]


def test_an_empty_flow_plans_nothing():
    """ノードが無ければ何も置かない（空の枠だけが残らないように）．"""
    plan = F.plan_flow(F.parse_flow(""), 0, 0, F.EMU * 8, F.EMU * 4)

    assert plan == F.FlowPlan()


def test_object_blocks_is_populated():
    """``OBJECT_BLOCKS`` が空でない（Issue: PR #186 のレビュー）．

    ``get_args`` で Union から実行時のタプルを導いているので、型の書き方を
    変えると**黙って空になる**ことがありうる。空のタプルを ``isinstance`` に
    渡すと常に偽で、オブジェクトブロックの判定が全部すり抜ける——
    エラーにならないぶん気づきにくい。``ARROW_DIRECTIONS`` と同じ要領で固定する。
    """
    from md2pptx.ir import OBJECT_BLOCKS, Arrow, Flow, Image, Seq, Table
    assert set(OBJECT_BLOCKS) == {Table, Flow, Image, Seq, Arrow}


def test_every_object_block_with_steps_can_be_sliced():
    """図の中で ``@step`` を持てる型は ``upto`` を実装している．

    ``_expand_steps`` は ``isinstance(last, (Flow, Seq))`` で判定するので、
    ここが実装と食い違うと段が黙って刻まれなくなる。
    """
    from md2pptx.ir import Flow, Seq
    for t in (Flow, Seq):
        assert callable(getattr(t, "upto", None)), t.__name__
