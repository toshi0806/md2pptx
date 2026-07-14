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
import math
import os
import struct
import sys

from pptx import Presentation
from pptx.enum.text import MSO_AUTO_SIZE, MSO_ANCHOR, PP_ALIGN
from pptx.enum.dml import MSO_THEME_COLOR
from pptx.enum.shapes import MSO_SHAPE
from pptx.oxml.ns import qn
from pptx.util import Inches, Pt

try:  # パッケージ実行・単体実行のどちらでも import できるように
    from .ir import Deck, Slide, TitleSlide, Line, Table, Flow, Image
    from .flow import plan_flow
except ImportError:  # pragma: no cover - 単体実行時のフォールバック
    from ir import Deck, Slide, TitleSlide, Line, Table, Flow, Image
    from flow import plan_flow


# Table.aligns の寄せ名 → PowerPoint の段落水平アラインメント．
_TABLE_ALIGN = {"left": PP_ALIGN.LEFT, "center": PP_ALIGN.CENTER, "right": PP_ALIGN.RIGHT}


def _read_image_size(path):
    """画像（png / jpg）のピクセル寸法 (width, height) をヘッダ解析で取得する．

    python-pptx の内部クラスや Pillow に依存せず，標準ライブラリだけでファイル
    ヘッダを読む（対応形式は png / jpeg）．寸法が読めない場合は ValueError．
    """
    with open(path, "rb") as f:
        head = f.read(8)
        if head[:8] == b"\x89PNG\r\n\x1a\n":
            # PNG: シグネチャ(8B)の直後が IHDR チャンク．チャンク長(4B)＋チャンクタイプ
            # "IHDR"(4B) を読み飛ばすと，先頭に width/height（ビッグエンディアン 4B×2）．
            f.read(4 + 4)  # チャンク長(4) + チャンクタイプ"IHDR"(4)
            w, h = struct.unpack(">II", f.read(8))
            return w, h
        if head[:2] == b"\xff\xd8":  # JPEG（SOI）．SOF マーカーまで走査する．
            f.seek(2)
            while True:
                b = f.read(1)
                if not b:
                    break
                if b != b"\xff":
                    continue
                marker = f.read(1)
                while marker == b"\xff":  # 連続する 0xFF（フィルバイト）を読み飛ばす
                    marker = f.read(1)
                if not marker:
                    break
                m = marker[0]
                # SOF0..SOF15（0xC0..0xCF）に寸法．ただし DHT/JPG/DAC は除く．
                if 0xC0 <= m <= 0xCF and m not in (0xC4, 0xC8, 0xCC):
                    f.read(2 + 1)  # segment length(2) + precision(1)
                    h, w = struct.unpack(">HH", f.read(4))
                    return w, h
                seg = f.read(2)
                if len(seg) < 2:
                    break
                length = struct.unpack(">H", seg)[0]
                f.seek(length - 2, os.SEEK_CUR)  # このセグメントを読み飛ばす
    raise ValueError(f"cannot read image dimensions (png/jpeg only): {path}")


class Renderer:
    """IR を pptx へ描画するレンダラ．

    base pptx（テーマのみを持つ 0 枚構成）を開き，レイアウトとテーマ色
    エイリアスを初期化する．スライドはすべて新規追加で生成する
    （thmx 由来の base は本文スライドを持たない）．
    """

    # フロー box の幅見積もりに掛ける安全係数（_text_width_pt の楽観的見積もりを補正）．
    _BOX_W_SAFETY = 1.15

    # 相対フォントサイズ 1 段あたりの倍率（≈12.5%）．拡大は ×，縮小は ÷．
    # 絶対 pt はハードコードせず，テーマ既定サイズ（_body_font_levels）からの相対比のみ持つ．
    _SIZE_STEP_RATIO = 1.125
    # 相対サイズの下限・上限（極小化／巨大化を防ぐ安全クランプ．段数の暴走対策）．
    _SIZE_MIN_PT = 8.0
    _SIZE_MAX_PT = 96.0

    def __init__(self, base_pptx_path, base_dir=None):
        self.prs = Presentation(base_pptx_path)
        # 画像などの相対パスを解決する基準ディレクトリ（既定は Markdown ファイルの
        # 置き場．cli が渡す）．None なら実行時のカレントを基準にする．
        self.base_dir = base_dir
        # テーマに pptx を渡した場合，元々入っているスライド（テンプレート用の
        # プレースホルダ枚）が先頭に残らないよう，常に 0 枚から描画を始める．
        self._clear_slides()
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

        # フロー図 box の自動配色（テーマアクセント色を順番に割当）．
        self._box_palette = [self.T2, self.A6, self.A6, self.GOLD, self.A2]
        # テーマ色名（DSL の {accent6} 等）→ MSO_THEME_COLOR の対応表．
        self._theme_map = {
            "accent1": MSO_THEME_COLOR.ACCENT_1, "accent2": MSO_THEME_COLOR.ACCENT_2,
            "accent3": MSO_THEME_COLOR.ACCENT_3, "accent4": MSO_THEME_COLOR.ACCENT_4,
            "accent5": MSO_THEME_COLOR.ACCENT_5, "accent6": MSO_THEME_COLOR.ACCENT_6,
            "tx1": MSO_THEME_COLOR.TEXT_1, "tx2": MSO_THEME_COLOR.TEXT_2,
            "bg1": MSO_THEME_COLOR.BACKGROUND_1, "bg2": MSO_THEME_COLOR.BACKGROUND_2,
        }
        # 本文スタイルのレベル別フォントサイズ（pt）のキャッシュ（None=未取得）．
        self._body_levels = None

        # レイアウト解決．title=0 / content=1 / section=2．
        layouts = self.prs.slide_layouts
        self.layouts = layouts
        self.title_layout = layouts[0]
        self.L1 = layouts[1] if len(layouts) > 1 else layouts[0]
        self.section_layout = layouts[2] if len(layouts) > 2 else self.L1

    # ------------------------------------------------------------ helpers
    def _clear_slides(self):
        """base pptx に既存のスライドがあれば取り除く（0 枚から描画するため）．"""
        sldIdLst = self.prs.slides._sldIdLst
        for sldId in list(sldIdLst):
            rId = sldId.get(qn("r:id"))
            if rId:
                try:
                    self.prs.part.drop_rel(rId)
                except KeyError:
                    pass
            sldIdLst.remove(sldId)

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

    def _body_font_levels(self):
        """本文スタイルのレベル別フォントサイズ（pt）のリストを返す（lvl1 始まり）．

        スライドマスターの ``p:txStyles/p:bodyStyle/a:lvl{n}pPr/a:defRPr@sz``
        を lvl1 から順に読み取る．表・図が標準サイズで収まらないとき，下位レベルの
        小さいサイズへ段階的に切り替えるために用いる．取得できなければ既定 [18]．
        """
        if self._body_levels is not None:
            return self._body_levels
        levels = []
        try:
            master = self.prs.slide_masters[0]
            body = master.element.find(
                qn("p:txStyles") + "/" + qn("p:bodyStyle"))
            if body is not None:
                for lvl in range(1, 10):
                    el = body.find(
                        qn("a:lvl%dpPr" % lvl) + "/" + qn("a:defRPr"))
                    if el is not None and el.get("sz"):
                        levels.append(int(el.get("sz")) / 100.0)
        except Exception:
            pass
        if not levels:
            levels = [18.0]
        self._body_levels = levels
        return levels

    def _body_font_size(self):
        """本文プレースホルダの標準フォントサイズ（pt．lvl1）を返す．"""
        return self._body_font_levels()[0]

    def _apply_size_delta(self, p, level, delta):
        """段落 p に相対フォントサイズ（delta 段）を適用する．

        基点はその段落の level に対応するテーマ既定サイズ（_body_font_levels）．
        実サイズ = round(base × 1.125**delta) を [_SIZE_MIN_PT, _SIZE_MAX_PT] で
        クランプする（大きな段数指定でも極小・巨大化しない）．p.level（インデント）は
        変更しない＝段落の既定文字書式（defRPr＝p.font）にサイズを設定するため，
        run が無い空段落でも有効で，bullet/採番記号も本文と同じサイズになる．

        - delta is None（未指定）: 何もしない（スライド既定もテーマ既定もそのまま）．
        - delta == 0（テーマ既定に固定）: サイズ指定を明示的に外しテーマ継承へ戻す．
        """
        if delta is None:
            return
        if delta == 0:
            # スライド既定（@body-size）を無効化し，テーマ既定サイズへ戻す．
            p.font.size = None
            return
        levels = self._body_font_levels()
        base = levels[min(level, len(levels) - 1)]
        size = round(base * self._SIZE_STEP_RATIO ** delta)
        size = min(self._SIZE_MAX_PT, max(self._SIZE_MIN_PT, size))
        p.font.size = Pt(size)

    def _title_font_size(self):
        """タイトルプレースホルダの既定フォントサイズ（pt．lvl1）を返す（既定 42）．"""
        try:
            master = self.prs.slide_masters[0]
            el = master.element.find(
                qn("p:txStyles") + "/" + qn("p:titleStyle")
                + "/" + qn("a:lvl1pPr") + "/" + qn("a:defRPr"))
            if el is not None and el.get("sz"):
                return int(el.get("sz")) / 100.0
        except Exception:
            pass
        return 42.0

    def _subtitle_font_size(self):
        """副題プレースホルダ（idx 1）の既定フォントサイズ（pt．lvl1）を返す（既定 28）．

        著者・所属欄はこのプレースホルダに入るため，相対サイズ段数（{-1} 等）の基点にする．
        （副題自体はタイトル枠内に別サイズで入るため基点が異なる：render_title_slide 参照．）
        """
        if self.title_layout is None:
            return 28.0
        try:
            for ph in self.title_layout.placeholders:
                if ph.placeholder_format.idx == 1:
                    for dr in ph._element.iter(qn("a:defRPr")):
                        if dr.get("sz"):
                            return int(dr.get("sz")) / 100.0
        except Exception:
            pass
        return 28.0

    def _size_from_delta(self, base_pt, delta):
        """基点サイズ base_pt（pt）に相対段数 delta を適用した pt 値を返す（範囲クランプ）．

        本文の _apply_size_delta と同じ 1.125 倍/段の比率．delta が None なら base をそのまま返す．
        """
        if delta is None:
            return base_pt
        size = round(base_pt * self._SIZE_STEP_RATIO ** delta)
        return min(self._SIZE_MAX_PT, max(self._SIZE_MIN_PT, size))

    @staticmethod
    def _text_width_pt(text, font_pt):
        """テキストの概算表示幅（pt）．全角は font_pt，半角は約 0.55×で見積もる．"""
        w = 0.0
        for ch in text or "":
            w += font_pt if ord(ch) > 0x2E80 else font_pt * 0.55
        return w

    def _fit_font(self, fits_at):
        """レベル別サイズを大きい順に試し，``fits_at(size)`` が真の最大サイズを返す．

        どのレベルでも収まらなければ最小レベル（最後）のサイズを返す（ベストエフォート）．
        """
        levels = self._body_font_levels()
        for sz in levels:
            if fits_at(sz):
                return sz
        return levels[-1]

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
            # 副題はタイトル枠内に，本文より少し小さめの文字で入れる．
            if ts.subtitle:
                sp = tf.add_paragraph()
                sp.text = ts.subtitle
                sp.space_before = Pt(6)
                # 副題はタイトル枠内に入るため基点はタイトル×0.8（副題プレースホルダの
                # 既定サイズではない）．著者・所属は別プレースホルダなので基点が異なるが，
                # いずれも「その要素が本来出るサイズから delta 段」で一貫する．
                sub_sz = self._size_from_delta(
                    self._title_font_size() * 0.8, ts.subtitle_delta)
                for r in sp.runs:
                    r.font.size = Pt(sub_sz)

        # 副題プレースホルダには著者・所属のみを入れる（副題はタイトル枠へ移動）．
        sub_ph = self._find_placeholder(s, 1)
        if sub_ph is not None:
            # 各行の相対サイズ段数（{-1} 等）を行と 1 対 1 で持ち回る（None＝未指定）．
            # affiliation_deltas は TitleSlide.__post_init__ で affiliation と同長が保証される．
            sub_lines = []
            sub_deltas = []
            if ts.author:
                sub_lines.append(ts.author)
                sub_deltas.append(ts.author_delta)
            for aff, delta in zip(ts.affiliation or [], ts.affiliation_deltas):
                sub_lines.append(aff)
                sub_deltas.append(delta)
            if sub_lines:
                # 所属行の折り返しを抑えるため右方向へ枠を広げる（左位置は維持）．
                # 継承ジオメトリの場合は 4 辺すべてを実効値で明示する
                # （一部だけ設定すると top/height が 0 に落ちて枠が移動するため）．
                left, top, width, height = self._effective_geom(sub_ph, s)
                if None not in (left, top, width, height):
                    new_w = self.SW - left - Inches(0.2)
                    if new_w > width:
                        sub_ph.left = left
                        sub_ph.top = top
                        sub_ph.height = height
                        sub_ph.width = new_w
                tf = sub_ph.text_frame
                tf.paragraphs[0].text = sub_lines[0]
                for ln in sub_lines[1:]:
                    tf.add_paragraph().text = ln
                # {-1}/{+1} 指定のある行だけ副題既定サイズを基点に段階調整する．
                base = self._subtitle_font_size()
                for para, delta in zip(tf.paragraphs, sub_deltas):
                    if delta is not None:
                        para.font.size = Pt(self._size_from_delta(base, delta))
        return s

    def _effective_geom(self, ph, slide):
        """プレースホルダの実効ジオメトリ (left, top, width, height) を返す．

        スライド上で未指定（継承）の値はレイアウトの同 idx プレースホルダで補う．
        """
        left, top, width, height = ph.left, ph.top, ph.width, ph.height
        if None in (left, top, width, height):
            idx = ph.placeholder_format.idx
            try:
                for lph in slide.slide_layout.placeholders:
                    if lph.placeholder_format.idx == idx:
                        left = left if left is not None else lph.left
                        top = top if top is not None else lph.top
                        width = width if width is not None else lph.width
                        height = height if height is not None else lph.height
                        break
            except Exception:
                pass
        return left, top, width, height

    def _find_placeholder(self, slide, idx):
        for ph in slide.placeholders:
            if ph.placeholder_format.idx == idx:
                return ph
        return None

    # ----------------------------------------------------------- 図形（flow）
    def box(self, slide, l, t, w, h, text, tc, sub=None, fsize=None, ssize=None):
        """角丸四角ノードを描く（塗りはテーマ色 tc，文字は背景色 BG）．

        fsize / ssize（pt）を省略するとテーマ既定のフォントサイズを継承する．
        """
        shp = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, l, t, w, h)
        shp.fill.solid()
        shp.fill.fore_color.theme_color = tc
        shp.line.color.theme_color = self.TX
        shp.line.width = Pt(0.5)
        tf = shp.text_frame
        tf.word_wrap = True
        tf.vertical_anchor = MSO_ANCHOR.MIDDLE
        tf.margin_left = Pt(4)
        tf.margin_right = Pt(4)
        tf.margin_top = Pt(1)
        tf.margin_bottom = Pt(1)
        pa = tf.paragraphs[0]
        pa.alignment = PP_ALIGN.CENTER
        pa.text = text
        for r in pa.runs:
            r.font.color.theme_color = self.BG
            if fsize is not None:
                r.font.size = Pt(fsize)
            r.font.bold = True
        if sub:
            p2 = tf.add_paragraph()
            p2.alignment = PP_ALIGN.CENTER
            p2.text = sub
            for r in p2.runs:
                r.font.color.theme_color = self.BG
                if ssize is not None:
                    r.font.size = Pt(ssize)
        return shp

    def arrow(self, slide, x1, y1, x2, y2, w=2.0):
        """矢印（直線コネクタ＋三角の矢じり）を描く．"""
        cn = slide.shapes.add_connector(2, x1, y1, x2, y2)
        cn.line.color.theme_color = self.TX
        cn.line.width = Pt(w)
        le = cn.line._get_or_add_ln()
        le.append(le.makeelement(
            qn("a:tailEnd"), {"type": "triangle", "w": "med", "len": "med"}))
        return cn

    def block_arrow(self, slide, x1, y1, x2, y2, thickness, color=None):
        """ノード間のすき間に塗りつぶしのブロック矢印を置く．

        太い直線＋三角矢じりは box に食い込み見栄えが悪いため，すき間に収まる
        塗り矢印シェイプ（RIGHT/DOWN_ARROW）を用いる．色はテーマ任せ（既定はアクセント）．
        """
        inset = Inches(0.05)
        if abs(x2 - x1) >= abs(y2 - y1):       # 横向き
            left = min(x1, x2) + inset
            width = abs(x2 - x1) - 2 * inset
            if width <= 0:
                left, width = min(x1, x2), abs(x2 - x1)
            height = thickness
            top = y1 - height // 2
            shape = MSO_SHAPE.RIGHT_ARROW if x2 >= x1 else MSO_SHAPE.LEFT_ARROW
        else:                                   # 縦向き
            top = min(y1, y2) + inset
            height = abs(y2 - y1) - 2 * inset
            if height <= 0:
                top, height = min(y1, y2), abs(y2 - y1)
            width = thickness
            left = x1 - width // 2
            shape = MSO_SHAPE.DOWN_ARROW if y2 >= y1 else MSO_SHAPE.UP_ARROW
        shp = slide.shapes.add_shape(shape, int(left), int(top),
                                     int(width), int(height))
        shp.fill.solid()
        shp.fill.fore_color.theme_color = color or self.TX
        shp.line.fill.background()
        return shp

    def note(self, slide, l, t, w, h, text, size, tc=None, bold=False,
             align=PP_ALIGN.LEFT):
        """注記用テキストボックスを描く（キャプション・矢印ラベル・省略記号）．"""
        tb = slide.shapes.add_textbox(l, t, w, h)
        tf = tb.text_frame
        tf.word_wrap = True
        pa = tf.paragraphs[0]
        pa.alignment = align
        pa.text = text
        for r in pa.runs:
            if size is not None:
                r.font.size = Pt(size)
            if tc is not None:
                r.font.color.theme_color = tc
            if bold:
                r.font.bold = True
        return tb

    def _theme_color(self, name):
        """テーマ色名を MSO_THEME_COLOR へ解決する（未知なら None）．"""
        if not name:
            return None
        return self._theme_map.get(str(name).lower())

    def _box_fits(self, node, bw, bh, font_pt):
        """box（主ラベル＋副ラベル）が指定フォントサイズで収まるか概算判定する．

        幅見積もりは安全係数 ``_BOX_W_SAFETY`` を掛けて保守的に評価する（``theme.thmx``
        のような半角主体ラベルが実 PowerPoint で 1 字あふれて折り返すのを防ぐ）．
        """
        line_h = font_pt * 1.2
        inner_w = max(1.0, bw / 12700.0 - 8)   # 左右マージン約 Pt(4)×2
        inner_h = bh / 12700.0 - 4             # 上下マージン約
        safe = self._BOX_W_SAFETY
        lines = max(1, math.ceil(
            safe * self._text_width_pt(node.label, font_pt) / inner_w))
        if node.sublabel:
            lines += max(1, math.ceil(
                safe * self._text_width_pt(node.sublabel, font_pt) / inner_w))
        return lines * line_h <= inner_h

    def render_flow(self, slide, flow, left, top, width, height):
        """Flow ブロックを矩形領域へ描画する（flow.plan_flow の座標プランを使用）．

        図中の文字サイズは本文プレースホルダの標準サイズに揃える．
        """
        plan = plan_flow(flow, left, top, width, height)
        # box が標準サイズで収まらなければ，全 box 一律で下位レベルへ切り替える．
        boxes = plan["boxes"]
        if boxes:
            bsz = self._fit_font(
                lambda sz: all(self._box_fits(node, bw, bh, sz)
                               for node, _, _, bw, bh in boxes))
        else:
            bsz = self._body_font_size()
        bi = 0
        for node, bl, bt, bw, bh in plan["boxes"]:
            tc = self._theme_color(node.color) or \
                self._box_palette[bi % len(self._box_palette)]
            bi += 1
            self.box(slide, bl, bt, bw, bh, node.label, tc,
                     sub=node.sublabel or None, fsize=bsz, ssize=bsz)
        for label, bl, bt, bw, bh in plan["ellipses"]:
            self.note(slide, bl, bt, bw, bh, label, bsz, tc=self.T2, bold=True,
                      align=PP_ALIGN.CENTER)
        # ノード間のすき間に塗りつぶしのブロック矢印を置く（太さは box 高に比例）．
        box_h_emu = boxes[0][4] if boxes else Inches(1.0)
        thick = int(box_h_emu * 0.34)
        for x1, y1, x2, y2 in plan["arrows"]:
            self.block_arrow(slide, x1, y1, x2, y2, thick)
        for text, bl, bt, bw, bh in plan["labels"]:
            self.note(slide, bl, bt, bw, bh, text, bsz, tc=self.T2, bold=True,
                      align=PP_ALIGN.CENTER)
        for text, bl, bt, bw, bh, role in plan["captions"]:
            # caption のみ図に付随（note_top / note_bottom はプレースホルダ側で描く）．
            if role == "caption":
                self.note(slide, bl, bt, bw, bh, text, bsz, tc=self.T2,
                          align=PP_ALIGN.CENTER)

    def _table_col_widths(self, ncols, width, col_ratios):
        """表の列幅（EMU）リストを返す（均等 or 比率指定）．"""
        if col_ratios and len(col_ratios) == ncols and sum(col_ratios) > 0:
            tot = float(sum(col_ratios))
            return [int(width * r / tot) for r in col_ratios]
        cw = int(width / ncols)
        return [cw] * ncols

    def _table_height_emu(self, data, col_w, font_pt):
        """指定フォントサイズでの表の概算総高（EMU）を見積もる（折り返し考慮）．

        実際の PowerPoint レンダリングは行間・最小行高などで見積りより伸びがち
        なため，安全係数を掛けて保守的（やや大きめ）に見積もる．
        """
        line_h = font_pt * 1.32          # 行間込みの行高
        cell_pad_pt = 6                  # セル上下マージン＋最小余白（約）
        side_pad_pt = 18                 # セル左右マージン合計（約．Pt(10)+Pt(6)＋余裕）
        safety = 1.15                    # 折り返し・最小行高ぶんの安全係数
        total_pt = 0.0
        for row in data:
            row_h = line_h + cell_pad_pt
            for ci, cw in enumerate(col_w):
                text = row[ci] if ci < len(row) else ""
                inner_pt = max(1.0, cw / 12700.0 - side_pad_pt)
                lines = max(1, math.ceil(self._text_width_pt(text, font_pt) / inner_pt))
                row_h = max(row_h, lines * line_h + cell_pad_pt)
            total_pt += row_h
        return int(total_pt * safety * 12700)

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

        col_w = self._table_col_widths(ncols, width, col_ratios)
        for ci, cw in enumerate(col_w):
            tbl.columns[ci].width = cw

        data = ([table.header] if table.header else []) + list(table.rows)
        # フォントは本文標準（lvl1）を基本に，収まらなければ下位レベルへ切り替える．
        fsize = self._fit_font(
            lambda sz: self._table_height_emu(data, col_w, sz) <= height)

        for ri, row in enumerate(data):
            is_header = bool(table.header) and ri == 0
            for ci in range(ncols):
                cell = tbl.cell(ri, ci)
                cell.vertical_anchor = MSO_ANCHOR.MIDDLE
                # 列の水平寄せ（区切り行のコロン由来）．未指定は左寄せ扱い．
                al = table.aligns[ci] if ci < len(table.aligns) else "left"
                cell.margin_left = Pt(10)
                # 右寄せ列は既定 margin_right(6pt)だと数字が右壁へ貼りつくため広げる．
                cell.margin_right = Pt(12) if al == "right" else Pt(6)
                cell.margin_top = Pt(2)
                cell.margin_bottom = Pt(2)
                pa = cell.text_frame.paragraphs[0]
                pa.text = row[ci] if ci < len(row) else ""
                if al != "left":
                    pa.alignment = _TABLE_ALIGN[al]
                for run in pa.runs:
                    run.font.size = Pt(fsize)
                    if is_header:
                        run.font.bold = True
                        run.font.color.theme_color = self.BG
                if is_header:
                    cell.fill.solid()
                    cell.fill.fore_color.theme_color = self.A2
        return gf

    def _resolve_image_path(self, src):
        """画像パスを base_dir 基準で解決し，存在しなければ fail fast する．"""
        path = src if os.path.isabs(src) else os.path.join(self.base_dir or ".", src)
        if not os.path.isfile(path):
            raise FileNotFoundError(f"image not found: {src}")
        return path

    @staticmethod
    def _crop_fractions(crop, W, H):
        """Crop（残す矩形）を PowerPoint のクロップ割合と可視画素サイズへ換算する．

        戻り値 (cl, ct, cr, cb, vis_w_px, vis_h_px)．cl 等は各辺で削る割合（0..1）．
        """
        if crop is None:
            return 0.0, 0.0, 0.0, 0.0, float(W), float(H)
        if crop.unit == "px":
            x, y, w, h = crop.x, crop.y, crop.w, crop.h
        else:  # ソース画像サイズに対する割合
            x, y = crop.x / 100.0 * W, crop.y / 100.0 * H
            w, h = crop.w / 100.0 * W, crop.h / 100.0 * H
        # 各辺で削る割合（0..1）に正規化して検証する．単位（px / %）に依らず
        # 分数で評価できるうえ，許容誤差も相対値（eps）で一貫して扱える．
        cl, ct = x / W, y / H
        cr, cb = (W - (x + w)) / W, (H - (y + h)) / H
        eps = 1e-6  # 割合換算の丸め誤差の吸収（絶対 px ではなく相対量で判定）
        if (w <= 0 or h <= 0 or cl < -eps or ct < -eps
                or cr < -eps or cb < -eps):
            raise ValueError(
                f"crop rectangle out of bounds for {W}x{H}px source: "
                f"keep x={x:g},y={y:g},w={w:g},h={h:g}")
        clamp = lambda v: min(1.0, max(0.0, v))  # 誤差ぶんを [0,1] に丸め込む
        return clamp(cl), clamp(ct), clamp(cr), clamp(cb), w, h

    @staticmethod
    def _resolve_len(length, base_emu):
        """Length を EMU（float）へ解決する．割合は base_emu 比，絶対はそのまま．None は None．"""
        if length is None:
            return None
        if length.unit == "percent":
            return length.value / 100.0 * base_emu
        return float(length.value)

    def render_image(self, slide, img, left, top, width, seg_h):
        """Image ブロックをセグメント矩形 (left, top, width, seg_h) 内に配置する．

        ソース画像のピクセル寸法を読み，crop（残す矩形）を PowerPoint のクロップ割合へ
        換算．width/height はアスペクト維持で解決し（両指定かつ fit=fill のときのみ歪ませ），
        align と縦中央でセグメント内へ収める．caption があれば画像下に描画する．
        overflow=True の場合は最終クランプを行わず，明示サイズのまま下方向への
        はみ出しを許可する（上端はセグメント上端まで．タイトル・導入文に重ねない）．
        """
        path = self._resolve_image_path(img.src)
        W, H = _read_image_size(path)                   # ソースのピクセル寸法
        cl, ct, cr, cb, vis_w, vis_h = self._crop_fractions(img.crop, W, H)
        aspect = (vis_w / vis_h) if vis_h else 1.0      # クロップ後の可視領域の比

        # キャプション用の高さを確保（1 行分）．
        cap_h = int(Pt(self._body_font_size()) * 1.4) if img.caption else 0
        avail_w = float(width)
        avail_h = float(max(1, seg_h - cap_h))

        # width/height を EMU へ解決（未指定は None）．
        w = self._resolve_len(img.width, avail_w)
        h = self._resolve_len(img.height, avail_h)
        if w is None and h is None:                     # 両省略：領域に内接
            w, h = self._fit_within(avail_w, avail_h, aspect)
        elif h is None:                                 # 幅のみ：高さは比で
            h = w / aspect
        elif w is None:                                 # 高さのみ：幅は比で
            w = h * aspect
        elif img.fit != "fill":                         # 両指定・contain：比維持で内接
            w, h = self._fit_within(w, h, aspect)
        # 極端な指定（0% 等）でも非正にならないよう下限を張る（ゼロ除算・負サイズ回避）．
        w, h = max(w, 1.0), max(h, 1.0)
        # セグメントを超えないよう最終クランプ（比維持）．overflow 指定時は
        # クランプせず，明示サイズのまま帯からのはみ出しを許可する．
        if not img.overflow:
            w, h = self._fit_within(min(w, avail_w), min(h, avail_h), w / h)

        # 水平寄せ（align）と縦中央でセグメント内へ配置．
        if img.align == "left":
            x = left
        elif img.align == "right":
            x = left + (avail_w - w)
        else:
            x = left + (avail_w - w) / 2.0
        y = top + (avail_h - h) / 2.0
        if img.overflow:
            # はみ出しは下（結論文・罫線側）のみ．上端はセグメント上端で止め，
            # タイトル・導入文には重ねない（top も y も同じ EMU 数値）．
            y = max(y, top)

        pic = slide.shapes.add_picture(path, int(x), int(y), int(w), int(h))
        if img.crop is not None:
            pic.crop_left, pic.crop_top = cl, ct
            pic.crop_right, pic.crop_bottom = cr, cb

        bottom = int(y + h)
        if img.caption:
            # 画像直下に置く．h ≤ avail_h なので通常 y+h ≤ top+avail_h だが，丸め等で
            # セグメント外へ出ないよう cap 上端を [.., top+seg_h-cap_h] にクランプする．
            # overflow 時は画像に追従してさらに下がる（スライド外に出うる）．
            cap_top = int(y + h)
            if not img.overflow:
                cap_top = min(cap_top, int(top + seg_h - cap_h))
            self._draw_caption(slide, img.caption, left, cap_top, width, cap_h)
            bottom = cap_top + cap_h
        if img.overflow and bottom > self.SH:
            sys.stderr.write(
                f"md2pptx: warning: overflowing image/caption ({img.src}) "
                "extends beyond the slide bottom edge\n"
            )
        return pic

    @staticmethod
    def _fit_within(box_w, box_h, aspect):
        """アスペクト比 aspect の矩形を (box_w, box_h) に内接させた (w, h) を返す．

        box_w / box_h / aspect が非正のときは安全側（最低 1 EMU）に倒し，ゼロ除算や
        負サイズを避ける（極端に狭いセグメントでも描画を止めない）．
        """
        if aspect <= 0 or box_w <= 0 or box_h <= 0:
            return max(box_w, 1.0), max(box_h, 1.0)
        if box_w / aspect <= box_h:     # 幅が制約：幅いっぱい
            return box_w, box_w / aspect
        return box_h * aspect, box_h    # 高さが制約：高さいっぱい

    def _draw_caption(self, slide, text, left, top, width, height):
        """図下キャプションを中央寄せの小さめ本文サイズで描く．"""
        tb = slide.shapes.add_textbox(left, top, width, max(height, Pt(12)))
        tf = tb.text_frame
        tf.word_wrap = True
        # キャプションは短文前提（1 行分の高さを確保）．枠を内容で伸ばさない
        # （長文で下方向へはみ出さないよう auto_size を無効化）．必要なら折り返す．
        tf.auto_size = MSO_AUTO_SIZE.NONE
        p = tf.paragraphs[0]
        p.alignment = PP_ALIGN.CENTER
        p.text = text
        # 本文標準より 1 段小さめ（テーマ既定サイズ体系の中で縮小）．
        levels = self._body_font_levels()
        size = levels[1] if len(levels) > 1 else levels[0]
        for r in p.runs:
            r.font.size = Pt(size)

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

        # @widths によるプレースホルダ幅の上書きは，本文描画より
        # 前に済ませる（以降の _effective_geom / _content_rect が上書き後を参照）．
        self._apply_placeholder_widths(s, directives,
                                       is_columns=bool(slide.columns))

        # スライド既定の採番色（@autonum-color）．Line.num_color が優先．
        default_num_color = directives.get("autonum_color")
        # スライド既定の相対サイズ段数（@body-size）．Line.size_delta が優先．
        default_size_delta = self._body_size_delta(directives)
        scale = self._autofit_scale(directives)
        blocks = slide.blocks or []

        if slide.columns:
            self._render_columns(s, slide.columns, default_num_color, scale,
                                 default_autofit, default_size_delta,
                                 self._col_ratios(directives))
        elif any(isinstance(b, (Table, Flow, Image)) for b in blocks):
            self._render_stacked(s, blocks, default_num_color, scale, default_autofit,
                                 self._col_ratios(directives), default_size_delta)
        else:
            line_blocks = [b for b in blocks if isinstance(b, Line)]
            body = self._body_placeholder(s)
            if body is not None and line_blocks:
                tf = body.text_frame
                self._fill_lines(tf, line_blocks, default_num_color,
                                 default_size_delta)
                self._apply_autofit(tf, scale, default_autofit)

        if slide_number:
            self.add_slide_number(s)
        return s

    def _body_size_delta(self, directives):
        """@body-size ディレクティブをスライド既定の相対サイズ段数へ解釈する．

        未指定・非整数値はいずれも None（＝スライド既定なし）を返す．None は
        「未指定」を明示する番兵で，size_delta=None の行はテーマ既定のままになる
        （0＝明示的に 0 段，とは区別する）．

        `@body-size: 0`（0 段）は「スライド既定なし」と同義として None を返す．
        スライド全体に対する 0 段は変化なし＝既定なしと区別する意味がないため．
        （行トークン `{0}` の「テーマ既定へ明示的に戻す」意味は別物で，スライド既定
        が非 0 のとき個別行を素のテーマ既定へ戻す用途に残る．Line.size_delta=0 が
        担い，こちらには波及しない．）

        parser 経由なら body_size は _INT_DIRECTIVES で int 化済みのため int()
        は素通りする．try/except は parser を介さず directives を直接組み立てる
        ケース（テスト・他コードからの呼び出し）に対する防御で，不正値で落とさず
        「スライド既定なし」に倒す．
        """
        val = directives.get("body_size")
        if val is None:
            return None
        try:
            iv = int(val)
        except (TypeError, ValueError):
            sys.stderr.write(
                f"md2pptx: warning: ignoring non-integer @body-size value "
                f"{val!r}\n"
            )
            return None
        return iv if iv != 0 else None

    def _render_columns(self, slide, columns, default_num_color, scale,
                        default_autofit, default_size_delta=None,
                        col_ratios=None):
        """多カラム（「2つのコンテンツ」）：各カラムを idx 1, 2 … へ流す．

        columns[i] を プレースホルダ idx=i+1 へ描画する（idx 0 はタイトル）．
        Line（箇条書き・採番・no_bullet）はプレースホルダへ流し込み，Table/Image/
        Flow を含むカラムはそのプレースホルダ矩形へ座標スタック配置する（地の文と
        混在する場合は _render_stacked_into が空行帯で棲み分ける）．
        """
        for ci, col_blocks in enumerate(columns):
            ph = self._find_placeholder(slide, ci + 1)
            if ph is None:
                continue  # レイアウトに該当プレースホルダが無ければスキップ
            if any(isinstance(b, (Table, Flow, Image)) for b in col_blocks):
                # カラム矩形へ表・図をスタック配置．継承ジオメトリはレイアウトで補う．
                # 通常 layout 3 は idx1/idx2 のジオメトリを持つため，解決失敗はテーマ
                # 異常時のみ．その場合は本文領域へフォールバックする（表が消えるより，
                # 見えて重なる方が原因に気づきやすい）が，複数カラムが重なりうるので警告
                # を出す．
                left, top, width, height = self._effective_geom(ph, slide)
                if None in (left, top, width, height):
                    sys.stderr.write(
                        f"md2pptx: warning: could not resolve geometry for "
                        f"column {ci}; falling back to the body area "
                        f"(columns may overlap)\n"
                    )
                    left, top, width, height = self._content_rect(slide)
                # @table-widths はスライド共通で全カラムの表に適用する．列数が比率の
                # 要素数と一致しない表は _table_col_widths が等幅へフォールバックする．
                self._render_stacked_into(slide, col_blocks, ph, left, top,
                                          width, height, default_num_color, scale,
                                          default_autofit, col_ratios,
                                          default_size_delta)
                continue
            # Line のみのカラムはプレースホルダへ直接流し込む．_render_stacked_into は
            # objects（表・図）が空だと何も描画せず return する設計なので，ここを通すと
            # 箇条書きが消える．そのため表・図を含むカラムとは意図的に経路を分ける．
            lines = [b for b in col_blocks if isinstance(b, Line)]
            if lines:
                tf = ph.text_frame
                self._fill_lines(tf, lines, default_num_color, default_size_delta)
                self._apply_autofit(tf, scale, default_autofit)

    # ----------------------------------------------------- 描画ユーティリティ
    def _append_lines(self, tf, line_blocks, first, default_num_color,
                      default_size_delta=None):
        """Line 列を text_frame に段落として追記する（採番／no_bullet を適用）．

        first=True なら最初の 1 行は既存の paragraphs[0] を使う．残りの行を
        追記しても良いよう，処理後の first 状態を返す．

        相対サイズは行の size_delta を優先し，None の行はスライド既定
        （default_size_delta，@body-size 由来）を継承する．
        """
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
            delta = blk.size_delta if blk.size_delta is not None else default_size_delta
            # 行トークン {0} は「スライド既定を無効化してテーマ既定へ戻す」意味だが，
            # そもそもスライド既定が無い（default_size_delta is None）なら戻す対象が
            # なく無意味な no-op なので適用しない（テーマの段落サイズに触れない）．
            if delta == 0 and default_size_delta is None:
                delta = None
            self._apply_size_delta(p, blk.level, delta)
        return first

    def _fill_lines(self, tf, line_blocks, default_num_color, default_size_delta=None):
        """Line 列を text_frame の段落として流し込む（先頭から）．"""
        self._append_lines(tf, line_blocks, True, default_num_color,
                           default_size_delta)

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
        """@table-widths ディレクティブ（"45,55" 等）を表の列幅比リストへ解釈する．"""
        v = directives.get("table_widths")
        if not v:
            return None
        try:
            return [float(x) for x in str(v).replace("，", ",").split(",")]
        except ValueError:
            return None

    def _split_allow_left(self, value):
        """値末尾の ``!``（左余白の使用許可）を分離して (本体文字列, フラグ) を返す．"""
        v = str(value).strip()
        if v.endswith(("!", "！")):
            return v[:-1].strip(), True
        return v, False

    def _parse_pct_list(self, value):
        """百分率リスト（例: "55,45"，"55%,45%"）を float の list へ解釈する．

        全角の区切り（，／％）も受理する．不正値は None を返す．
        """
        try:
            return [float(str(x).strip().rstrip("%％"))
                    for x in str(value).replace("，", ",").split(",")]
        except ValueError:
            return None

    def _override_geom(self, ph, left, top, width, height):
        """スライド側へ明示ジオメトリを書き，レイアウト継承を上書きする．

        4 値すべて設定するのは，一部のみ明示すると xfrm が不完全になり
        PowerPoint 側の解釈が実装依存になるため（継承値で補って全指定する）．
        """
        ph.left, ph.top = int(left), int(top)
        ph.width, ph.height = int(width), int(height)

    _PH_MARGIN = Inches(0.1)   # プレースホルダ拡幅時にスライド端へ残す余白

    def _apply_placeholder_widths(self, slide, directives, is_columns):
        """@widths をスライド種別（単カラム／多カラム）に応じて適用する．

        いずれも「標準の使用可能幅に対する百分率」で解釈する（詳細は各メソッド）．
        値 1 個は単カラム本文幅（例: "104"），複数はカラムごとの幅（例: "62,40"）．
        拡幅は**左端固定・右余白のみ**が既定（箇条書きの行頭位置がスライド間で
        揃い，遷移時の見た目が安定する）．右余白で収まらない指定はクランプして
        警告する．値の末尾に ``!`` を付けると（例: "108!" / "62,47!"），収まら
        ない分だけ左余白へ逃がすことを許可する．その場合もスライド端は余白
        _PH_MARGIN でクランプし，それでも収まらない指定は警告のうえ比例縮小する．

        スライド種別と値の個数が合わない指定は無視し，警告を出す．
        """
        val = directives.get("widths")
        if val is None:
            return
        if is_columns:
            self._apply_ph_widths(slide, val)
        else:
            self._apply_body_width(slide, val)

    def _apply_ph_widths(self, slide, val):
        """@widths: "55,45" — 多カラムのプレースホルダ幅を再指定する．

        カラム群の合計スパンからカラム間ギャップを除いた幅を 100% とし，
        各カラム幅を百分率で再指定する（ギャップは維持）．合計が 100 を
        超えると全体が右方向へ広がる（55,50 → 全体が標準の 105%）．
        ジオメトリを解決できない場合は警告して何もしない（従来描画）．
        """
        val, allow_left = self._split_allow_left(val)
        pcts = self._parse_pct_list(val)
        if not pcts or any(p <= 0 for p in pcts):
            sys.stderr.write(
                f"md2pptx: warning: ignoring invalid @widths value {val!r}\n")
            return
        if len(pcts) < 2:
            sys.stderr.write(
                "md2pptx: warning: @widths on a multi-column slide expects "
                f"one value per column, got {val!r}; ignoring\n")
            return
        # md のカラム順＝プレースホルダ idx 順（_render_columns と同じ対応）で集める．
        phs = []
        for i, _pct in enumerate(pcts):
            ph = self._find_placeholder(slide, i + 1)
            if ph is None:
                sys.stderr.write(
                    f"md2pptx: warning: @widths has {len(pcts)} values but "
                    f"column placeholder {i + 1} does not exist; ignoring\n")
                return
            geom = self._effective_geom(ph, slide)
            if None in geom:
                sys.stderr.write(
                    "md2pptx: warning: @widths skipped "
                    "(could not resolve column geometry)\n")
                return
            phs.append((ph, geom))
        lefts = [g[0] for _, g in phs]
        rights = [g[0] + g[2] for _, g in phs]
        gaps = [lefts[i + 1] - rights[i] for i in range(len(phs) - 1)]
        if any(g < 0 for g in gaps):
            # 重なったプレースホルダ（負のギャップ）は usable を過大にするため 0 扱い．
            sys.stderr.write(
                "md2pptx: warning: @widths found overlapping column "
                "placeholders; treating the negative gap as 0\n")
            gaps = [max(g, 0) for g in gaps]
        span_l, span_r = lefts[0], rights[-1]
        usable = (span_r - span_l) - sum(gaps)
        widths = [usable * p / 100.0 for p in pcts]
        new_span = sum(widths) + sum(gaps)
        # 既定は左端固定（右余白のみ使用）．"...!" で左余白の使用を許可する．
        max_span = ((self.SW - self._PH_MARGIN - span_l) if not allow_left
                    else (self.SW - 2 * self._PH_MARGIN))
        if new_span > max_span:
            sys.stderr.write(
                "md2pptx: warning: @widths total exceeds the "
                f"{'slide' if allow_left else 'right margin'}; clamping"
                f"{'' if allow_left else ' (append ! to use the left margin)'}\n")
            k = (max_span - sum(gaps)) / float(sum(widths))
            widths = [w * k for w in widths]
            new_span = max_span
        new_left = span_l
        if allow_left:
            overflow = (new_left + new_span) - (self.SW - self._PH_MARGIN)
            if overflow > 0:
                new_left -= overflow
            new_left = max(new_left, self._PH_MARGIN)
        x = new_left
        for i, ((ph, (_l, t, _w, h)), nw) in enumerate(zip(phs, widths)):
            self._override_geom(ph, x, t, nw, h)
            x += nw + (gaps[i] if i < len(gaps) else 0)

    def _apply_body_width(self, slide, val):
        """@widths: "105" — 単カラム本文プレースホルダ幅を再指定する．

        継承した本文プレースホルダ幅に対する百分率（% 付き可）．値は 1 個のみ
        （複数値は多カラム用）．ジオメトリを解決できない場合は何もしない（従来描画）．
        """
        val, allow_left = self._split_allow_left(val)
        pcts = self._parse_pct_list(val)
        if pcts is None:
            sys.stderr.write(
                f"md2pptx: warning: ignoring invalid @widths value {val!r}\n")
            return
        if len(pcts) != 1:
            sys.stderr.write(
                "md2pptx: warning: @widths on a single-column slide expects "
                f"exactly 1 value, got {val!r}; ignoring\n")
            return
        pct = pcts[0]
        if pct <= 0:
            sys.stderr.write(
                f"md2pptx: warning: ignoring non-positive @widths value {val!r}\n")
            return
        ph = self._body_placeholder(slide)
        if ph is None:
            return
        left, top, width, height = self._effective_geom(ph, slide)
        if None in (left, top, width, height):
            return
        new_w = width * pct / 100.0
        # 既定は左端固定（右余白のみ使用）．"...!" で左余白の使用を許可する．
        max_w = ((self.SW - self._PH_MARGIN - left) if not allow_left
                 else (self.SW - 2 * self._PH_MARGIN))
        if new_w > max_w:
            sys.stderr.write(
                "md2pptx: warning: @widths exceeds the "
                f"{'slide' if allow_left else 'right margin'}; clamping"
                f"{'' if allow_left else ' (append ! to use the left margin)'}\n")
            new_w = max_w
        new_l = left
        if allow_left:
            overflow = (new_l + new_w) - (self.SW - self._PH_MARGIN)
            if overflow > 0:
                new_l -= overflow
            new_l = max(new_l, self._PH_MARGIN)
        self._override_geom(ph, new_l, top, new_w, height)

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

    def _note_to_line(self, text):
        """Flow の note 文字列を本文プレースホルダ用の Line へ変換する．

        行頭マーカー（→ は no_bullet）の最小限の解釈だけ行う．
        """
        t = (text or "").strip()
        if t.startswith("→"):
            return Line(text=t, kind="plain")
        return Line(text=t, kind="bullet")

    def _obj_weight(self, obj):
        """オブジェクト（Table / Flow / Image）の縦方向の重み（高さ配分用）．"""
        if isinstance(obj, Flow):
            return max(4, len(obj.nodes) + 2)
        if isinstance(obj, Image):
            # 画像は帯を広めに確保（キャプションぶんを少し足す）．細かな大きさは
            # width/height でセグメント内に調整する．
            return 8 + (1 if obj.caption else 0)
        return max(2, len(obj.rows) + (1 if obj.header else 0))

    def _stack_objects(self, slide, objects, left, top, width, height, col_ratios):
        """Table / Flow / Image を矩形領域内に重みづけで縦に積んで座標配置する．"""
        weights = [self._obj_weight(o) for o in objects]
        total = float(sum(weights)) or 1.0
        gap = Pt(6)
        avail = height - gap * (len(objects) - 1)
        y = top
        for obj, w in zip(objects, weights):
            seg_h = int(avail * w / total)
            if isinstance(obj, Flow):
                self.render_flow(slide, obj, left, y, width, seg_h)
            elif isinstance(obj, Image):
                self.render_image(slide, obj, left, y, width, seg_h)
            else:
                self.render_table(slide, obj, left, y, width, seg_h, col_ratios)
            y += seg_h + gap

    def _render_stacked(self, slide, blocks, default_num_color, scale,
                        default_autofit, col_ratios, default_size_delta=None):
        """表／図を含むスライドを描画する．

        地の文（Line）は **標準の本文プレースホルダ**へ流し込み，表・図だけを
        座標配置する．プレースホルダには「導入文＋空行スペーサ＋結論文」を入れ，
        確保した中央帯に表・図を重ねる（``参照スクリプト`` の図スライドと同方式）．
        地の文を自由位置のテキストボックスには置かない．
        """
        left, top, width, height = self._content_rect(slide)
        body = self._body_placeholder(slide)
        self._render_stacked_into(slide, blocks, body, left, top, width, height,
                                  default_num_color, scale, default_autofit,
                                  col_ratios, default_size_delta)

    def _render_stacked_into(self, slide, blocks, body, left, top, width, height,
                             default_num_color, scale, default_autofit, col_ratios,
                             default_size_delta=None):
        """``blocks`` を矩形 (left, top, width, height) 内へスタック描画する．

        地の文（Line）は ``body`` プレースホルダへ流し込み，表・図は矩形内に
        座標配置する．描画先（プレースホルダ＋矩形）を引数で受けるため，本文領域
        （単一カラム）にも多カラムの各カラム矩形にも使える．``body`` が None の
        場合は地の文を捨て，矩形全体にオブジェクトを積む．
        """
        # 地の文（前後）とオブジェクト（表・図）に分ける．
        # Flow の note(top)/note(bottom) も地の文としてプレースホルダへ回す．
        prose_before, objects, prose_after = [], [], []
        seen_obj = False
        for b in blocks:
            if isinstance(b, (Table, Flow, Image)):
                if isinstance(b, Flow) and b.note_top:
                    bucket = prose_after if seen_obj else prose_before
                    bucket.append(self._note_to_line(b.note_top))
                objects.append(b)
                seen_obj = True
                if isinstance(b, Flow) and b.note_bottom:
                    prose_after.append(self._note_to_line(b.note_bottom))
            elif isinstance(b, Line):
                (prose_after if seen_obj else prose_before).append(b)
        if not objects:
            return

        # 表・図スライドの地の文に相対サイズが効くと，帯高計算（_body_font_size
        # 固定）と食い違い，帯が詰まって結論文が重なりうる（既知の制約．TODO(v2)）．
        # 判定は _append_lines と同じ実効デルタ（行トークン優先，無ければスライド既定
        # @body-size）で行う．行 size_delta=None でも @body-size 由来で拡縮する場合を
        # 取りこぼさないため．0／None（変化なし）は対象外．
        def _eff_delta(ln):
            return ln.size_delta if ln.size_delta is not None else default_size_delta
        # 0／None は「サイズ変化なし」＝帯高（本文標準サイズ前提）と食い違わないので
        # 対象外．{0} はテーマ既定＝標準サイズそのものなので警告不要．
        if any(_eff_delta(ln) not in (None, 0) for ln in prose_before + prose_after):
            sys.stderr.write(
                "md2pptx: warning: relative font size on body text of a "
                "table/figure slide may cause layout crowding "
                "(band height is estimated at the standard body size)\n"
            )

        # 地の文が無ければプレースホルダは使わず，領域全体にオブジェクトを置く．
        if not prose_before and not prose_after:
            if body is not None:
                body._element.getparent().remove(body._element)
            self._stack_objects(slide, objects, left, top, width, height, col_ratios)
            return

        # 地の文あり：プレースホルダに導入文＋空行＋結論文を流して中央帯を確保．
        # 帯と空行数はプレースホルダ矩形から逆算し，地の文＋空行＋結論文が
        # プレースホルダ高を超えないようにする（結論文がスライド外へ出ない）．
        # TODO(v2): prose の size_delta を行高に反映する（現在は本文標準サイズ固定）．
        # 導入文を {+2} 等で大きく拡大すると帯が詰まり結論文と重なりうる（既知の制約）．
        bsz = self._body_font_size()
        line_h = int(Pt(bsz) * 1.32)        # 行間込みの保守的な行高
        nb, na = len(prose_before), len(prose_after)
        inset = Pt(4)
        band_h = height - (nb + na) * line_h - 2 * inset
        if band_h < Inches(0.8):
            band_h = Inches(0.8)
        band_top = top + nb * line_h + inset
        blanks = max(1, int(band_h / line_h))   # 帯を埋める空行数（超過しない）

        if body is not None:
            tf = body.text_frame
            first = self._append_lines(tf, prose_before, True, default_num_color,
                                       default_size_delta)
            for _ in range(blanks):
                p = tf.paragraphs[0] if first else tf.add_paragraph()
                first = False
            self._append_lines(tf, prose_after, first, default_num_color,
                               default_size_delta)
            self._apply_autofit(tf, scale, default_autofit)

        # 結論文との重なりを避けるため帯を少しだけ詰めてオブジェクトを置く．
        obj_h = max(Inches(0.8), band_h - Pt(8))
        self._stack_objects(slide, objects, left, band_top, width, obj_h, col_ratios)

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


def build(deck, base_pptx_path, out_path, base_dir=None):
    """Deck を base pptx 上に描画して out_path に保存する（CLI 用エントリ）．

    Args:
        base_dir: 画像などの相対パスを解決する基準ディレクトリ（既定は Markdown の
            置き場）．None なら実行時のカレントを基準にする．

    Returns:
        out_path（保存先パス）．
    """
    r = Renderer(base_pptx_path, base_dir=base_dir)
    r.render(deck)
    r.save(out_path)
    return out_path
