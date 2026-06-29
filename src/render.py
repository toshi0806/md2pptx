#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""IR（ir.py）→ pptx 描画（md2pptx ステージ2 / Phase 1）．

ステージ0（thmx2pptx.py）が生成した base pptx を土台に開き，IR の Deck を
走査して 1 つのプレゼンテーションへ描画する．配色・フォントはテーマ
（thmx）任せで，スクリプト側で色・フォントをハードコードしない（図形のみ
テーマのアクセント色を参照する）．

``参照スクリプト`` のモジュール大域に依存したヘルパ群（no_bullet /
add_slide_number / add_bullets / set_autonum / enum_items / fit_body /
content_slide）を Renderer のメソッドへ移植し，self.prs / self.layouts /
テーマ色エイリアスから状態を解決する（DESIGN.md §6）．

使い方::

    from thmx2pptx import thmx_to_pptx
    from ir import Deck
    from render import build

    base = thmx_to_pptx("theme.thmx")
    build(deck, base, "out.pptx")
"""
from __future__ import annotations

import copy
import sys

from pptx import Presentation
from pptx.enum.text import MSO_AUTO_SIZE
from pptx.enum.dml import MSO_THEME_COLOR
from pptx.oxml.ns import qn

try:  # パッケージ実行・単体実行のどちらでも import できるように
    from .ir import Deck, Slide, TitleSlide, Line, Table, Flow
except ImportError:  # pragma: no cover - 単体実行時のフォールバック
    from ir import Deck, Slide, TitleSlide, Line, Table, Flow


class Renderer:
    """IR を pptx へ描画するレンダラ．

    base pptx（テーマのみを持つ 0 枚構成）を開き，レイアウトとテーマ色
    エイリアスを初期化する．スライドはすべて新規追加で生成する
    （thmx 由来の base は本文スライドを持たない）．
    """

    def __init__(self, base_pptx_path):
        self.prs = Presentation(base_pptx_path)
        self.SW = self.prs.slide_width
        self.SH = self.prs.slide_height

        # テーマのアクセント色（図形用）．テキスト色・フォントはテーマ任せ．
        # Phase 1 では box/arrow/note/table（Phase 2/3）が使うため保持のみ．
        self.A2 = MSO_THEME_COLOR.ACCENT_2       # 緑
        self.A6 = MSO_THEME_COLOR.ACCENT_6       # 緑（濃）
        self.T2 = MSO_THEME_COLOR.TEXT_2         # 緑系テキスト
        self.GOLD = MSO_THEME_COLOR.ACCENT_1     # 金
        self.BG = MSO_THEME_COLOR.BACKGROUND_1   # 背景（白）
        self.TX = MSO_THEME_COLOR.TEXT_1         # 本文色（黒）

        # レイアウト解決．title=0 / content=1 / section=2．
        layouts = self.prs.slide_layouts
        self.layouts = layouts
        self.title_layout = layouts[0]
        self.L1 = layouts[1] if len(layouts) > 1 else layouts[0]
        self.section_layout = layouts[2] if len(layouts) > 2 else self.L1

    # ------------------------------------------------------------ helpers
    def no_bullet(self, para):
        """段落の行頭記号を消す（結論行など）．"""
        pPr = para._p.get_or_add_pPr()
        for tag in ("a:buChar", "a:buAutoNum"):
            e = pPr.find(qn(tag))
            if e is not None:
                pPr.remove(e)
        if pPr.find(qn("a:buNone")) is None:
            pPr.append(pPr.makeelement(qn("a:buNone"), {}))

    def add_slide_number(self, slide):
        """スライド自身のレイアウトの番号プレースホルダ（idx==12）を複製して有効化．

        セクションスライド（レイアウト2）など L1 以外の上でも，そのスライドの
        実レイアウトから番号プレースホルダを取得する．
        """
        for lph in slide.slide_layout.placeholders:
            if lph.placeholder_format.idx == 12:
                slide.shapes._spTree.append(copy.deepcopy(lph._element))
                return

    def add_bullets(self, tf, items):
        """(level, text) の列を本文 text_frame に箇条書きとして流し込む．"""
        first = True
        for lvl, txt in items:
            p = tf.paragraphs[0] if first else tf.add_paragraph()
            first = False
            p.level = lvl
            p.text = txt

    def set_autonum(self, p, fmt="arabicPeriod", color=None):
        """段落の行頭記号を自動採番（1. 2. 3. …）に切り替える（enumerate 相当）．
           color にテーマ色名（例 "tx1"）を渡すと採番記号の色を指定する．"""
        pPr = p._p.get_or_add_pPr()
        if color:
            for tag in ("a:buClrTx", "a:buClr"):
                el = pPr.find(qn(tag))
                if el is not None:
                    pPr.remove(el)
            buClr = pPr.makeelement(qn("a:buClr"), {})
            buClr.append(buClr.makeelement(qn("a:schemeClr"), {"val": color}))
            pPr.insert(0, buClr)  # buClr は採番記号より前に置く
        bu = pPr.makeelement(qn("a:buAutoNum"), {"type": fmt})
        for tag in ("a:buChar", "a:buNone", "a:buAutoNum"):
            el = pPr.find(qn(tag))
            if el is not None:
                pPr.replace(el, bu)
                break
        else:
            pPr.append(bu)

    def enum_items(self, tf, items):
        """(見出し, 説明) を，見出し=自動採番(level0)・説明=通常箇条書き(level1) で並べる．"""
        first = True
        for head, desc in items:
            p = tf.paragraphs[0] if first else tf.add_paragraph()
            first = False
            p.level = 0
            p.text = head
            self.set_autonum(p)
            d = tf.add_paragraph()
            d.level = 1
            d.text = desc

    def fit_body(self, tf, scale=None):
        """本文プレースホルダに自動調整（normAutofit）を設定する．
        scale を与えると縮小率（%）を明示的に焼き込む．"""
        tf.word_wrap = True
        tf.auto_size = MSO_AUTO_SIZE.TEXT_TO_FIT_SHAPE
        if scale is not None:
            bodyPr = tf._txBody.bodyPr
            na = bodyPr.find(qn("a:normAutofit"))
            if na is None:
                # auto_size 設定で生成されない環境向けのフォールバック
                na = bodyPr.makeelement(qn("a:normAutofit"), {})
                bodyPr.append(na)
            na.set("fontScale", str(int(scale * 1000)))

    def _body_placeholder(self, slide):
        """本文プレースホルダ（idx==1）を返す．無ければ None．"""
        for ph in slide.placeholders:
            if ph.placeholder_format.idx == 1:
                return ph
        return None

    def content_slide(self, title, items):
        """タイトル＋箇条書きの基本スライドを 1 枚追加して返す．"""
        s = self.prs.slides.add_slide(self.L1)
        s.shapes.title.text = title
        body = self._body_placeholder(s)
        self.add_bullets(body.text_frame, items)
        self.fit_body(body.text_frame)
        self.add_slide_number(s)
        return s

    # --------------------------------------------------------- title slide
    def render_title_slide(self, ts: TitleSlide):
        """タイトルスライドを 1 枚目として追加する（番号は付けない）．

        base は 0 枚構成のため，タイトルレイアウト上に新規スライドを追加し，
        CENTER_TITLE（idx==0）に title を，SUBTITLE（idx==1）に
        subtitle＋author＋affiliation を流し込む．
        """
        s = self.prs.slides.add_slide(self.title_layout)

        title_ph = self._find_placeholder(s, 0)
        if title_ph is not None and ts.title:
            tf = title_ph.text_frame
            lines = ts.title.split("\n")
            tf.paragraphs[0].text = lines[0]
            for ln in lines[1:]:
                tf.add_paragraph().text = ln

        sub_ph = self._find_placeholder(s, 1)
        if sub_ph is not None:
            tf = sub_ph.text_frame
            sub_lines = []
            if ts.subtitle:
                sub_lines.append(ts.subtitle)
            if ts.author:
                sub_lines.append(ts.author)
            sub_lines.extend(ts.affiliation or [])
            if sub_lines:
                tf.paragraphs[0].text = sub_lines[0]
                for ln in sub_lines[1:]:
                    tf.add_paragraph().text = ln
        return s

    def _find_placeholder(self, slide, idx):
        for ph in slide.placeholders:
            if ph.placeholder_format.idx == idx:
                return ph
        return None

    # ------------------------------------------------------- content slide
    def render_slide(self, slide: Slide, slide_number=True, default_autofit=True):
        """コンテンツスライドを 1 枚追加して返す．

        ``slide.blocks`` を出現順に処理する．Phase 1 では Line ブロックのみを
        本文プレースホルダへ描画し，Table / Flow はスキップする（TODO）．
        """
        directives = slide.directives or {}
        layout_idx = directives.get("layout", slide.layout)
        try:
            layout = self.layouts[layout_idx]
        except (IndexError, TypeError):
            layout = self.L1

        s = self.prs.slides.add_slide(layout)
        if slide.title is not None and s.shapes.title is not None:
            s.shapes.title.text = slide.title

        # スライド既定の採番色（@autonum-color）．Line.num_color が優先．
        default_num_color = directives.get("autonum_color")

        body = self._body_placeholder(s)
        line_blocks = [b for b in slide.blocks if isinstance(b, Line)]

        if body is not None and line_blocks:
            tf = body.text_frame
            first = True
            for blk in line_blocks:
                p = tf.paragraphs[0] if first else tf.add_paragraph()
                first = False
                p.level = blk.level
                p.text = blk.text
                if blk.kind == "autonum":
                    fmt = blk.num_style or "arabicPeriod"
                    color = blk.num_color or default_num_color
                    self.set_autonum(p, fmt, color=color)
                elif blk.kind == "plain":
                    self.no_bullet(p)
                # kind == "bullet" はテーマ既定のまま

            # autofit 適用．数値でない @autofit は警告して既定挙動へフォールバック．
            autofit = directives.get("autofit")
            scale = None
            if autofit is not None:
                try:
                    scale = float(autofit)
                except (TypeError, ValueError):
                    sys.stderr.write(
                        f"md2pptx: warning: ignoring non-numeric @autofit "
                        f"value {autofit!r}\n"
                    )
            if scale is not None:
                self.fit_body(tf, scale=scale)
            elif default_autofit:
                self.fit_body(tf)

        # Table / Flow ブロックは Phase 2 / 3 で対応（ここではスキップ）．
        for blk in slide.blocks:
            if isinstance(blk, (Table, Flow)):
                # TODO(Phase 2/3): Table=add_table（ヘッダ着色）, Flow=flow.py レイアウタ
                continue

        if slide_number:
            self.add_slide_number(s)
        return s

    # ------------------------------------------------------------- deck
    def render(self, deck: Deck):
        """Deck 全体を描画し，Presentation を返す．"""
        meta = deck.meta or {}
        slide_number = meta.get("slide_number", True)
        default_autofit = meta.get("default_autofit", True)

        if deck.title_slide is not None:
            self.render_title_slide(deck.title_slide)

        for sl in deck.slides:
            self.render_slide(
                sl,
                slide_number=slide_number,
                default_autofit=default_autofit,
            )
        return self.prs

    def save(self, path):
        """現在の Presentation を保存する．"""
        self.prs.save(path)
        return path


def build(deck, base_pptx_path, out_path):
    """Deck を base pptx 上に描画して out_path に保存する（CLI 用エントリ）．

    Returns:
        out_path（保存先パス）．
    """
    r = Renderer(base_pptx_path)
    r.render(deck)
    r.save(out_path)
    return out_path
