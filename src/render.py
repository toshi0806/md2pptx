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
from pptx.enum.text import MSO_AUTO_SIZE, MSO_ANCHOR
from pptx.enum.dml import MSO_THEME_COLOR
from pptx.oxml.ns import qn
from pptx.util import Inches, Pt

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

    def render_table(self, slide, table, left, top, width, height, col_ratios=None):
        """Table ブロックを座標指定で 1 つ描画する（ヘッダ行をアクセント色で着色）．

        ``参照スクリプト`` の表描画を移植・一般化したもの．列幅は既定で均等，
        ``col_ratios`` を与えると比率配分する．配色はテーマ任せ（ヘッダのみ
        アクセント色 A2＋背景色 BG の文字）．
        """
        nrows = len(table.rows) + (1 if table.header else 0)
        ncols = max(
            len(table.header) if table.header else 0,
            max((len(r) for r in table.rows), default=0),
        )
        if nrows == 0 or ncols == 0:
            return None

        gf = slide.shapes.add_table(nrows, ncols, left, top, width, height)
        tbl = gf.table

        # 列幅：均等 or 比率指定．
        if col_ratios and len(col_ratios) == ncols and sum(col_ratios) > 0:
            tot = float(sum(col_ratios))
            for ci, r in enumerate(col_ratios):
                tbl.columns[ci].width = int(width * r / tot)
        else:
            cw = int(width / ncols)
            for ci in range(ncols):
                tbl.columns[ci].width = cw

        data = ([table.header] if table.header else []) + list(table.rows)
        # 行数に応じてフォントサイズを調整（多いほど小さく）．
        fsize = 24 if nrows <= 4 else (18 if nrows <= 7 else 14)

        for ri, row in enumerate(data):
            is_header = bool(table.header) and ri == 0
            for ci in range(ncols):
                cell = tbl.cell(ri, ci)
                cell.vertical_anchor = MSO_ANCHOR.MIDDLE
                cell.margin_left = Pt(10)
                cell.margin_right = Pt(6)
                cell.margin_top = Pt(2)
                cell.margin_bottom = Pt(2)
                pa = cell.text_frame.paragraphs[0]
                pa.text = row[ci] if ci < len(row) else ""
                for run in pa.runs:
                    run.font.size = Pt(fsize)
                    if is_header:
                        run.font.bold = True
                        run.font.color.theme_color = self.BG
                if is_header:
                    cell.fill.solid()
                    cell.fill.fore_color.theme_color = self.A2
        return gf

    # ------------------------------------------------------- content slide
    def render_slide(self, slide: Slide, slide_number=True, default_autofit=True):
        """コンテンツスライドを 1 枚追加して返す．

        ``slide.blocks`` を出現順に処理する．表を含まないスライドは Phase 1 と
        同じく本文プレースホルダへ Line を流し込み，表を含むスライドは座標スタック
        配置（テキスト→表→テキスト …）で描画する．Flow は Phase 3 でスキップ．
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
        scale = self._autofit_scale(directives)
        blocks = slide.blocks or []

        if any(isinstance(b, Table) for b in blocks):
            self._render_stacked(s, blocks, default_num_color, scale, default_autofit,
                                 self._col_ratios(directives))
        else:
            line_blocks = [b for b in blocks if isinstance(b, Line)]
            body = self._body_placeholder(s)
            if body is not None and line_blocks:
                tf = body.text_frame
                self._fill_lines(tf, line_blocks, default_num_color)
                self._apply_autofit(tf, scale, default_autofit)

        if slide_number:
            self.add_slide_number(s)
        return s

    # ----------------------------------------------------- 描画ユーティリティ
    def _fill_lines(self, tf, line_blocks, default_num_color):
        """Line 列を text_frame の段落として流し込む（採番／no_bullet を適用）．"""
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

    def _autofit_scale(self, directives):
        """@autofit ディレクティブを縮小率へ解釈する（非数値は警告して None）．"""
        autofit = directives.get("autofit")
        if autofit is None:
            return None
        try:
            return float(autofit)
        except (TypeError, ValueError):
            sys.stderr.write(
                f"md2pptx: warning: ignoring non-numeric @autofit value "
                f"{autofit!r}\n"
            )
            return None

    def _apply_autofit(self, tf, scale, default_autofit):
        """縮小率指定があれば焼き込み，無ければ既定の自動調整を設定する．"""
        if scale is not None:
            self.fit_body(tf, scale=scale)
        elif default_autofit:
            self.fit_body(tf)

    def _col_ratios(self, directives):
        """@col-widths ディレクティブ（"45,55" 等）を比率リストへ解釈する．"""
        v = directives.get("col_widths")
        if not v:
            return None
        try:
            return [float(x) for x in str(v).replace("，", ",").split(",")]
        except ValueError:
            return None

    def _content_rect(self, slide):
        """本文領域の矩形 (left, top, width, height) を返す（座標配置の基準）．"""
        ph = self._body_placeholder(slide)
        if ph is not None and None not in (ph.left, ph.top, ph.width, ph.height):
            return (ph.left, ph.top, ph.width, ph.height)
        try:
            for lph in slide.slide_layout.placeholders:
                if lph.placeholder_format.idx == 1 and None not in (
                    lph.left, lph.top, lph.width, lph.height
                ):
                    return (lph.left, lph.top, lph.width, lph.height)
        except Exception:
            pass
        # 既定：タイトル下の本文相当領域．
        return (Inches(0.6), Inches(1.7), self.SW - Inches(1.2),
                self.SH - Inches(2.3))

    def _render_text_box(self, slide, lines, left, top, width, height,
                         default_num_color, scale, default_autofit):
        """テキストセグメントを座標指定のテキストボックスへ描画する．"""
        tb = slide.shapes.add_textbox(left, top, width, height)
        tf = tb.text_frame
        self._fill_lines(tf, lines, default_num_color)
        self._apply_autofit(tf, scale, default_autofit)
        return tb

    def _render_stacked(self, slide, blocks, default_num_color, scale,
                        default_autofit, col_ratios):
        """表を含むスライドを，テキストと表のセグメントへ分けて縦に積む．

        本文プレースホルダの矩形を内容領域とし，テキストはテキストボックス，
        表は座標指定のテーブルとして重ならないように上から配置する．
        """
        left, top, width, height = self._content_rect(slide)

        # 座標配置するため空の本文プレースホルダは取り除く．
        body = self._body_placeholder(slide)
        if body is not None:
            body._element.getparent().remove(body._element)

        # ブロックをテキスト／表のセグメントへ（出現順を保つ）．
        segments = []
        cur_lines = []
        for b in blocks:
            if isinstance(b, Line):
                cur_lines.append(b)
            elif isinstance(b, Table):
                if cur_lines:
                    segments.append(("text", cur_lines))
                    cur_lines = []
                segments.append(("table", b))
            # Flow は Phase 3 で対応（ここでは無視）．
        if cur_lines:
            segments.append(("text", cur_lines))
        if not segments:
            return

        def weight(seg):
            if seg[0] == "text":
                return max(1, len(seg[1]))
            t = seg[1]
            return max(2, len(t.rows) + (1 if t.header else 0))

        weights = [weight(s) for s in segments]
        total = float(sum(weights))
        gap = Pt(6)
        avail = height - gap * (len(segments) - 1)
        y = top
        for seg, w in zip(segments, weights):
            seg_h = int(avail * w / total)
            if seg[0] == "text":
                self._render_text_box(slide, seg[1], left, y, width, seg_h,
                                      default_num_color, scale, default_autofit)
            else:
                self.render_table(slide, seg[1], left, y, width, seg_h, col_ratios)
            y += seg_h + gap

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
