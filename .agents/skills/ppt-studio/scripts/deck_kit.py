#!/usr/bin/env python3
"""deck_kit — 样式包（style pack）驱动的 PPT 渲染引擎。

设计原则
--------
1. **样式与内容分离**：`styles/<name>.json` 只描述设计令牌（颜色/字阶/几何/装饰），
   本文件不含任何具体风格。
2. **版式原型固定、可扩展**：每页一个 archetype（cover / statement / stats ...），
   新增版式 = 加一个 `slide_<name>()` 函数并注册到 `ARCHETYPES`。
3. **排版度量显式**：`disp_bottom()` 用字号 × 行距 × line_factor 推算真实占位，
   避免大字号文本框「溢出压字」——这是中文大标题最常见的翻车点。
4. **中英文分权重**：`rich()` 按字符类别拆 run，中文用 Semibold 家族、数字用 Bold，
   避免整段粗体导致中文字重过重。

用法见 `build_deck.py`。
"""

from __future__ import annotations

import json
import os
import re
from typing import Any, Iterable, Sequence

from lxml import etree
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.util import Inches, Pt

# ---------------------------------------------------------------- 基础工具

CJK_RE = re.compile(r"([\u3000-\u303f\u3000-\u303f\u4e00-\u9fff\uff00-\uffef\u2014\u2018\u2019\u201c\u201d]+)")

ALIGN = {"left": PP_ALIGN.LEFT, "center": PP_ALIGN.CENTER, "right": PP_ALIGN.RIGHT}


def hexrgb(h: str) -> RGBColor:
    h = h.lstrip("#")
    return RGBColor(int(h[0:2], 16), int(h[2:4], 16), int(h[4:6], 16))


def _deep_get(d: dict, path: str, default=None):
    cur: Any = d
    for k in path.split("."):
        if not isinstance(cur, dict) or k not in cur:
            return default
        cur = cur[k]
    return cur


class Style:
    """样式包：把 tokens JSON 包成好用的访问器。"""

    def __init__(self, data: dict, path: str = ""):
        self.data = data
        self.path = path
        self.name = data.get("name", "custom")
        self.label = data.get("label", self.name)
        self.fonts = data.get("fonts", {})
        self.type = data.get("type", {})
        self.track = data.get("tracking", {})
        self.geo = data.get("geometry", {})
        self.decor = data.get("decor", {})
        self._colors = data.get("colors", {})

    # --- 字体
    @property
    def latin(self) -> str:
        return self.fonts.get("latin", "Helvetica Neue")

    @property
    def cjk(self) -> str:
        return self.fonts.get("cjk", "PingFang SC")

    @property
    def cjk_bold(self) -> str:
        return self.fonts.get("cjk_bold") or self.cjk

    @property
    def weight_mode(self) -> str:
        # split: 中文用 cjk_bold 家族、数字用 bold 标志；flag: 一律 bold 标志
        return self.fonts.get("weight_mode", "split")

    # --- 颜色
    def c(self, key: str, tone: str = "light") -> RGBColor | None:
        """按语义取色，带深浅底自动切换。"""
        table = {
            "bg": ("light_bg", "dark_bg"),
            "tint": ("tint_bg", "dark_bg"),
            "ink": ("ink", "ink_dark"),
            "muted": ("muted", "muted_dark"),
            "hair": ("hair", "hair_dark"),
            "accent": ("accent", "accent_dark"),
            "card": ("card", "card_dark"),
            "card_line": ("card_line", "card_line_dark"),
            "ink_on_card": ("ink", "ink_dark"),
            "muted_on_card": ("muted", "muted_dark"),
        }
        light_key, dark_key = table.get(key, (key, key))
        # 只有 dark 页用深色对应值；light / tint 都用浅色值
        name = dark_key if tone == "dark" else light_key
        val = self._colors.get(name) or self._colors.get(light_key)
        return hexrgb(val) if val else None

    # --- 字阶
    def size(self, role: str, default: float = 14) -> float:
        return float(self.type.get(role, default))

    def spc(self, role: str) -> float:
        return float(self.track.get(role, 0))

    # --- 几何
    def g(self, key: str, default: float) -> float:
        return float(self.geo.get(key, default))

    @property
    def line_factor(self) -> float:
        return self.g("line_factor", 1.34)


def load_style(path_or_name: str, style_dir: str | None = None) -> Style:
    p = path_or_name
    if not os.path.isfile(p):
        base = style_dir or os.path.join(os.path.dirname(os.path.dirname(__file__)), "styles")
        cand = os.path.join(base, p if p.endswith(".json") else p + ".json")
        if not os.path.isfile(cand):
            avail = ", ".join(sorted(
                f[:-5] for f in os.listdir(base) if f.endswith(".json") and not f.startswith("_")))
            raise SystemExit(f"找不到样式 '{path_or_name}'。可用：{avail}")
        p = cand
    with open(p, encoding="utf-8") as f:
        return Style(json.load(f), p)


# ---------------------------------------------------------------- 文本引擎

class Deck:
    def __init__(self, style: Style, w: float = 13.3333, h: float = 7.5):
        self.st = style
        self.W, self.H = w, h
        self.M = style.g("margin_x", 0.95)
        self.CW = w - 2 * self.M
        self.prs = Presentation()
        self.prs.slide_width = Inches(w)
        self.prs.slide_height = Inches(h)
        self._tone = "light"

    # ---------- 版心 ----------
    def disp_bottom(self, y: float, size: float, lines: int = 1, ls: float = 1.0) -> float:
        return y + lines * size * ls * self.st.line_factor / 72.0

    # ---------- 形状 ----------
    def rect(self, x, y, w, h, fill=None, *, radius=None, border=None, border_w=0.0):
        shape = MSO_SHAPE.ROUNDED_RECTANGLE if radius else MSO_SHAPE.RECTANGLE
        shp = self.prs.slides[-1].shapes.add_shape(shape, Inches(x), Inches(y), Inches(w), Inches(h))
        if radius:
            shp.adjustments[0] = min(0.5, float(radius))
        shp.shadow.inherit = False
        if fill is None:
            shp.fill.background()
        else:
            shp.fill.solid()
            shp.fill.fore_color.rgb = fill
        if border is None or not border_w:
            shp.line.fill.background()
        else:
            shp.line.color.rgb = border
            shp.line.width = Pt(border_w)
        return shp

    def hairline(self, x, y, w, color=None, weight=1.0):
        color = color or self.st.c("hair", self._tone)
        return self.rect(x, y, w, weight / 72.0, fill=color)

    def dot(self, cx, cy, r, color=None):
        color = color or self.st.c("ink", self._tone)
        shp = self.prs.slides[-1].shapes.add_shape(
            MSO_SHAPE.OVAL, Inches(cx - r), Inches(cy - r), Inches(2 * r), Inches(2 * r))
        shp.shadow.inherit = False
        shp.fill.solid()
        shp.fill.fore_color.rgb = color
        shp.line.fill.background()
        return shp

    # ---------- 卡片（装饰由样式包决定） ----------
    def card(self, x, y, w, h, *, tone_card: bool = True):
        d = self.st.decor
        shape = d.get("card_shape", "rounded")
        radius = d.get("card_radius", 0.05) if shape == "rounded" else None
        fill = self.st.c("card", self._tone)
        border = self.st.c("card_line", self._tone)
        bw = float(d.get("card_border", 0) or 0)
        shadow = d.get("card_shadow", "none")

        if shadow == "hard":  # 贴纸风：先画一层错位的硬投影
            off = float(d.get("shadow_offset", 0.07))
            self.rect(x + off, y + off, w, h, fill=self.st.c("shadow", self._tone) or hexrgb("#000000"),
                      radius=radius)
        elif shadow == "soft":
            off = float(d.get("shadow_offset", 0.05))
            ghost = self.rect(x + 0.01, y + off, w, h, fill=self.st.c("shadow", self._tone),
                              radius=radius)
            ghost.fill.fore_color.rgb = self.st.c("shadow", self._tone)

        return self.rect(x, y, w, h, fill=fill, radius=radius, border=border, border_w=bw)

    # ---------- 文本 ----------
    def _run(self, p, txt, size, color, bold, spc, cjk_face):
        r = p.add_run()
        r.text = txt
        f = r.font
        f.size = Pt(size)
        f.bold = bold
        f.color.rgb = color
        f.name = self.st.latin
        rPr = r._r.get_or_add_rPr()
        latin = rPr.find(qn("a:latin"))
        ea = rPr.find(qn("a:ea"))
        if ea is None:
            ea = etree.SubElement(rPr, qn("a:ea"))
        ea.set("typeface", cjk_face)
        if latin is not None:
            latin.addnext(ea)
        if spc:
            rPr.set("spc", str(int(spc * 100)))

    def _rich(self, p, s: str, size, color, bold: bool, spc: float):
        st = self.st
        cjk_face = st.cjk_bold if (bold and st.weight_mode == "split") else st.cjk
        # split 模式下中文用家族体现字重，拉丁/数字仍靠 bold 标志
        cjk_bold_flag = bold and st.weight_mode != "split"
        for seg in CJK_RE.split(s):
            if not seg:
                continue
            is_cjk = bool(CJK_RE.fullmatch(seg))
            self._run(p, seg, size, color, cjk_bold_flag if is_cjk else bold, spc, cjk_face)

    def text(self, x, y, w, body, *, size=14, color=None, bold=False, role=None,
             align="left", spc=None, line=1.0, space=0, h=None, anchor="top"):
        st = self.st
        if role:
            size = st.size(role, size)
            if spc is None:
                spc = st.spc(role)
        if spc is None:
            spc = 0
        color = color or st.c("ink", self._tone)
        items = body if isinstance(body, (list, tuple)) else [body]
        if h is None:
            h = self.disp_bottom(0, size, len(items), line) + 0.02
        tb = self.prs.slides[-1].shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
        tf = tb.text_frame
        tf.word_wrap = True
        tf.auto_size = None
        for a in ("margin_left", "margin_right", "margin_top", "margin_bottom"):
            setattr(tf, a, 0)
        tf.vertical_anchor = MSO_ANCHOR.TOP if anchor == "top" else MSO_ANCHOR.MIDDLE
        for i, para in enumerate(items):
            p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
            p.alignment = ALIGN[align]
            p.line_spacing = line
            p.space_before = Pt(0)
            p.space_after = Pt(space)
            self._rich(p, str(para), size, color, bold, spc)
        return tb

    def eyebrow(self, text, y=None, x=None):
        y = y if y is not None else self.st.g("header_y", 0.92)
        x = x if x is not None else self.M
        self.text(x, y, self.CW, text, role="eyebrow", bold=True,
                  color=self.st.c("muted", self._tone))

    # ---------- 页面 ----------
    def new_slide(self, tone: str = "light", pattern: str | None = None):
        self._tone = tone
        s = self.prs.slides.add_slide(self.prs.slide_layouts[6])
        bg_key = {"dark": "bg", "light": "bg", "tint": "tint"}[tone]
        fill = self.st.c(bg_key, "dark" if tone == "dark" else "light")
        if tone == "light":
            fill = self.st.c("bg", "light")
        self.rect(0, 0, self.W, self.H, fill=fill)

        pat = pattern or self.st.decor.get("bg_pattern", "none")
        if pat and pat != "none":
            self._pattern(pat, tone)

        bar = self.st.decor.get("top_bar")
        if bar:
            bh = float(bar.get("h", 0.12)) if isinstance(bar, dict) else float(bar)
            bc = hexrgb(bar["color"]) if isinstance(bar, dict) and bar.get("color") else self.st.c("accent", tone)
            self.rect(0, 0, self.W, bh, fill=bc)

        if tone == "dark" and self.st.decor.get("glow"):
            self._glow()
        return s

    def _pattern(self, kind: str, tone: str):
        color = self.st.c("hair", tone)
        if kind == "dots":
            step = float(self.st.decor.get("pattern_step", 0.34))
            r = float(self.st.decor.get("pattern_dot", 0.012))
            x = self.M / 2
            while x < self.W:
                y = 0.4
                while y < self.H:
                    self.dot(x, y, r, color)
                    y += step
                x += step
        elif kind == "grid":  # 只用发丝线画网格
            step = float(self.st.decor.get("pattern_step", 0.7))
            x = self.M
            while x < self.W - self.M:
                self.rect(x, 0, 1 / 72.0, self.H, fill=color)
                x += step
            y = self.H / 2
            while y < self.H:
                self.rect(0, y, self.W, 1 / 72.0, fill=color)
                y += step

    def _glow(self):
        """深色页的环境光：本地生成一张径向渐变 PNG 再铺上去。"""
        color = self.st.decor.get("glow_color", "#5A78FF")
        alpha = float(self.st.decor.get("glow_alpha", 0.42))
        path = _glow_asset(color, alpha)
        if path:
            x = float(self.st.decor.get("glow_x", 4.2))
            y = float(self.st.decor.get("glow_y", -1.6))
            w = float(self.st.decor.get("glow_w", 10.5))
            h = float(self.st.decor.get("glow_h", 6.0))
            self.prs.slides[-1].shapes.add_picture(path, Inches(x), Inches(y), Inches(w), Inches(h))

    def save(self, path: str):
        os.makedirs(os.path.dirname(os.path.abspath(path)) or ".", exist_ok=True)
        self.prs.save(path)
        return path


def qn(tag: str) -> str:
    from pptx.oxml.ns import qn as _qn
    return _qn(tag)


# ---------------------------------------------------------------- 资源

_GLOW_CACHE = os.path.join(os.path.dirname(os.path.abspath(__file__)), ".cache")


def _glow_asset(color: str, alpha: float) -> str | None:
    """返回一张径向渐变 PNG 的路径；没有 Pillow 时静默跳过。"""
    try:
        from PIL import Image
        import numpy as np
    except Exception:
        return None
    os.makedirs(_GLOW_CACHE, exist_ok=True)
    rgb = hexrgb(color)
    name = f"glow_{color.lstrip('#')}_{int(alpha*100)}.png"
    path = os.path.join(_GLOW_CACHE, name)
    if os.path.isfile(path):
        return path
    w = h = 900
    y, x = np.mgrid[0:h, 0:w]
    d = np.sqrt(((x - w / 2) / (w * 0.42)) ** 2 + ((y - h / 2) / (h * 0.42)) ** 2)
    a = np.clip(1 - d, 0, 1) ** 2.4 * alpha
    arr = np.zeros((h, w, 4), np.uint8)
    arr[..., 0], arr[..., 1], arr[..., 2] = rgb[0], rgb[1], rgb[2]
    arr[..., 3] = (a * 255).astype(np.uint8)
    Image.fromarray(arr, "RGBA").save(path)
    return path


# ---------------------------------------------------------------- 版式原型
# 每个函数签名统一为 (deck, spec) → None，注册在 ARCHETYPES。

def slide_cover(d: Deck, s: dict):
    st = d.st
    d.new_slide(s.get("tone", "dark"))
    if s.get("eyebrow"):
        d.eyebrow(s["eyebrow"])
    y = s.get("title_y", st.g("cover_y", 1.6))
    d.text(d.M, y, d.CW, s.get("title", ""), role="hero", bold=True, line=1.0)
    y = d.disp_bottom(y, st.size("hero", 96)) + 0.30
    if s.get("subtitle"):
        d.text(d.M, y, min(9.0, d.CW), s["subtitle"], role="lead",
               color=st.c("muted", d._tone), line=st.g("lead_line", 1.45))
    if s.get("foot_left") or s.get("foot_right"):
        fy = st.g("foot_y", 5.72)
        d.hairline(d.M, fy, d.CW)
        if s.get("foot_left"):
            d.text(d.M, fy + 0.23, 6.0, s["foot_left"], role="micro", color=st.c("muted", d._tone))
        if s.get("foot_right"):
            d.text(d.W - d.M - 6.0, fy + 0.23, 6.0, s["foot_right"], role="micro",
                   color=st.c("muted", d._tone), align="right")


def slide_statement(d: Deck, s: dict):
    st = d.st
    d.new_slide(s.get("tone", "tint"))
    if s.get("eyebrow"):
        d.eyebrow(s["eyebrow"])
    y = st.g("content_y", 1.34)
    d.text(d.M, y, d.CW * 0.92, s.get("title", ""), role="title", bold=True, line=1.08)
    y = d.disp_bottom(y, st.size("title", 52), s.get("title_lines", 1), 1.08) + 0.22
    if s.get("body"):
        d.text(d.M, y, min(9.8, d.CW), s["body"], role="lead",
               color=st.c("muted", d._tone), line=st.g("lead_line", 1.55))


def slide_bigstat(d: Deck, s: dict):
    st = d.st
    d.new_slide(s.get("tone", "tint"))
    if s.get("eyebrow"):
        d.eyebrow(s["eyebrow"])
    y = st.g("content_y", 1.3)
    d.text(d.M, y, d.CW * 0.62, s.get("title", ""), role="title", bold=True)
    y = d.disp_bottom(y, st.size("title", 52)) + 0.18
    if s.get("caption"):
        d.text(d.M, y, 5.0, s["caption"], role="small", color=st.c("muted", d._tone))
        y = d.disp_bottom(y, st.size("small", 12)) + 0.28
    d.text(d.M, y, d.CW * 0.5, s.get("number", ""), role="stat_huge", bold=True, line=1.0)
    y = d.disp_bottom(y, st.size("stat_huge", 124)) - 0.05
    if s.get("notes"):
        d.text(d.M, y, d.CW * 0.52, s["notes"], role="small",
               color=st.c("muted", d._tone), line=1.55)
    items = s.get("items") or []
    if items:
        rx = s.get("right_x", 6.95)
        yy = st.g("list_y", 2.34)
        for it in items:
            d.hairline(rx, yy, d.W - d.M - rx)
            d.text(rx, yy + 0.26, d.W - d.M - rx, it.get("title", ""), role="body_xl", bold=True)
            if it.get("body"):
                d.text(rx, yy + 0.68, d.W - d.M - rx, it["body"], role="small",
                       color=st.c("muted", d._tone))
            yy += float(it.get("step", st.g("list_step", 1.28)))


def slide_stats(d: Deck, s: dict):
    st = d.st
    d.new_slide(s.get("tone", "dark"))
    if s.get("eyebrow"):
        d.eyebrow(s["eyebrow"])
    y = st.g("content_y", 1.45)
    d.text(d.M, y, d.CW * 0.94, s.get("title", ""), role="title", bold=True, line=1.05)
    y = d.disp_bottom(y, st.size("title", 52), 1, 1.05) + 0.24
    if s.get("subtitle"):
        d.text(d.M, y, min(9.6, d.CW), s["subtitle"], role="lead",
               color=st.c("muted", d._tone), line=st.g("lead_line", 1.5))

    stats = s.get("stats") or []
    n = max(1, len(stats))
    col_w = (d.CW - 0.3 * (n - 1)) / n
    sy = float(s.get("stats_y", st.g("stats_y", 3.86)))
    for i, stt in enumerate(stats):
        x = d.M + i * (col_w + 0.3)
        d.text(x, sy, col_w, stt.get("value", ""), role="stat_big", bold=True, line=1.0)
        ny = d.disp_bottom(sy, st.size("stat_big", 62)) + 0.04
        d.text(x, ny, col_w, stt.get("label", ""), role="micro", bold=True)
        if stt.get("note"):
            d.text(x, ny + 0.32, col_w, stt["note"], role="micro", color=st.c("muted", d._tone))
        if i and s.get("dividers", True):
            d.rect(x - 0.16, sy + 0.05, 0.007, 1.0, fill=st.c("hair", d._tone))

    if s.get("footnote"):
        d.hairline(d.M, d.H - 1.16, d.CW)
        d.text(d.M, d.H - 0.95, d.CW, s["footnote"], role="small", color=st.c("muted", d._tone))


def slide_cards(d: Deck, s: dict):
    st = d.st
    d.new_slide(s.get("tone", "light"))
    if s.get("eyebrow"):
        d.eyebrow(s["eyebrow"])
    y = st.g("content_y", 1.34)
    d.text(d.M, y, d.CW * 0.94, s.get("title", ""), role="title", bold=True)
    y = d.disp_bottom(y, st.size("title", 52)) + 0.20
    if s.get("subtitle"):
        body = s["subtitle"]
        d.text(d.M, y, min(9.8, d.CW), body, role="lead", color=st.c("muted", d._tone),
               line=st.g("lead_line", 1.6))

    cards = s.get("cards") or []
    n = max(1, len(cards))
    gap = float(s.get("gap", 0.26))
    cw = (d.CW - gap * (n - 1)) / n
    cy = float(s.get("cards_y", st.g("cards_y", 3.9)))
    ch = float(s.get("card_h", 1.95))
    for i, c in enumerate(cards):
        x = d.M + i * (cw + gap)
        d.card(x, cy, cw, ch)
        pad = float(s.get("pad", 0.30))
        d.text(x + pad, cy + 0.28, cw - 2 * pad, c.get("title", ""), role="card_title", bold=True)
        body_y = cy + 0.28 + d.disp_bottom(0, st.size("card_title", 20))
        if c.get("body"):
            d.text(x + pad, body_y + 0.10, cw - 2 * pad, c["body"], role="small",
                   color=st.c("muted", d._tone), line=st.g("card_line_sp", 1.5))
        if c.get("tag"):
            d.text(x + pad, cy + ch - 0.46, cw - 2 * pad, c["tag"], role="micro",
                   color=st.c("accent", d._tone), bold=True)


def slide_lanes(d: Deck, s: dict):
    st = d.st
    d.new_slide(s.get("tone", "tint"))
    if s.get("eyebrow"):
        d.eyebrow(s["eyebrow"])
    y = st.g("content_y", 1.3)
    d.text(d.M, y, d.CW * 0.8, s.get("title", ""), role="title_lg", bold=True)

    items = s.get("items") or []
    cols = int(s.get("cols", 2))
    gap = float(s.get("gap", 0.30))
    tw = (d.CW - gap * (cols - 1)) / cols
    th = float(s.get("tile_h", 1.82))
    top = float(s.get("items_y", 2.42))
    for i, it in enumerate(items):
        col, row = i % cols, i // cols
        x = d.M + col * (tw + gap)
        yy = top + row * (th + gap)
        d.card(x, yy, tw, th)
        pad = 0.34
        d.text(x + pad, yy + 0.28, tw - 2 * pad, it.get("title", ""), role="card_title", bold=True)
        if it.get("detail"):
            d.text(x + pad, yy + 0.74, tw - 2 * pad + 0.05, it["detail"], role="small",
                   color=st.c("muted", d._tone), line=1.5)
        if it.get("tag"):
            d.text(x + pad, yy + th - 0.42, tw - 2 * pad, it["tag"], role="micro",
                   color=st.c("accent", d._tone), bold=True)


def slide_timeline(d: Deck, s: dict):
    st = d.st
    d.new_slide(s.get("tone", "light"))
    if s.get("eyebrow"):
        d.eyebrow(s["eyebrow"])
    y = st.g("content_y", 1.34)
    d.text(d.M, y, d.CW * 0.9, s.get("title", ""), role="title", bold=True)
    y = d.disp_bottom(y, st.size("title", 52)) + 0.22
    if s.get("subtitle"):
        d.text(d.M, y, d.CW * 0.9, s["subtitle"], role="lead", color=st.c("muted", d._tone))

    nodes = s.get("nodes") or []
    tl = float(s.get("line_y", 4.32))
    d.hairline(d.M, tl, d.CW)
    n = len(nodes)
    for i, nd in enumerate(nodes):
        frac = nd.get("pos", i / max(1, n - 1))
        cx = d.M + float(frac) * d.CW
        accent = i in (0, n - 1)
        d.dot(cx, tl + 0.007, 0.043,
              st.c("ink", d._tone) if accent else st.c("muted", d._tone))
        tw = float(s.get("label_w", 1.6))
        tx = min(max(cx - tw / 2, 0.3), d.W - tw - 0.3)
        if nd.get("name"):
            d.text(tx, tl - 0.52, tw, nd["name"], role="body_sm", bold=True, align="center")
        if nd.get("label"):
            d.text(tx, tl + 0.22, tw, nd["label"], role="micro",
                   color=st.c("muted", d._tone), align="center")
        if nd.get("tag"):
            d.text(tx, tl + 0.50, tw, nd["tag"], role="micro_faint",
                   color=st.c("muted", d._tone), align="center")

    if s.get("footnote"):
        d.text(d.M, d.H - 0.95, d.CW, s["footnote"], role="small", color=st.c("muted", d._tone))


def slide_bullets(d: Deck, s: dict):
    """文字较多的版式：左侧标题，右侧条目（公文体常用）。"""
    st = d.st
    d.new_slide(s.get("tone", "light"))
    if s.get("eyebrow"):
        d.eyebrow(s["eyebrow"])
    y = st.g("content_y", 1.34)
    d.text(d.M, y, d.CW * 0.38, s.get("title", ""), role="title", bold=True, line=1.08)
    rx = d.M + d.CW * 0.42
    rw = d.W - d.M - rx
    yy = y
    for it in (s.get("bullets") or []):
        head = it.get("title", "")
        d.text(rx, yy, rw, head, role="body_xl", bold=True)
        yy = d.disp_bottom(yy, st.size("body_xl", 17), 1, 1.2) + 0.10
        if it.get("body"):
            d.text(rx, yy, rw, it["body"], role="body", color=st.c("muted", d._tone),
                   line=st.g("body_line", 1.6))
            yy += d.disp_bottom(0, st.size("body", 14), len(str(it["body"]).split("\n")), 1.6)
        yy += 0.34


def slide_table(d: Deck, s: dict):
    st = d.st
    d.new_slide(s.get("tone", "light"))
    if s.get("eyebrow"):
        d.eyebrow(s["eyebrow"])
    y = st.g("content_y", 1.26)
    d.text(d.M, y, d.CW * 0.7, s.get("title", ""), role="title_lg", bold=True)

    cols = s.get("columns") or []
    rows = s.get("rows") or []
    widths = s.get("widths") or [d.CW / len(cols)] * len(cols)
    scale = d.CW / sum(widths)
    widths = [w * scale for w in widths]
    top = float(s.get("table_y", 2.16))
    hdr_h = float(s.get("header_h", 0.36))
    row_h = float(s.get("row_h", 0.50))

    shp = d.prs.slides[-1].shapes.add_table(
        len(rows) + 1, len(cols), Inches(d.M), Inches(top),
        Inches(sum(widths)), Inches(hdr_h + row_h * len(rows)))
    tbl = shp.table
    tbl.first_row = True
    tbl.horz_banding = False
    for i, wd in enumerate(widths):
        tbl.columns[i].width = Inches(wd)
    tbl.rows[0].height = Inches(hdr_h)
    for i in range(1, len(rows) + 1):
        tbl.rows[i].height = Inches(row_h)

    decorate_table(d, tbl, s, cols, rows)
    if s.get("footnote"):
        d.text(d.M, d.H - 0.58, d.CW, s["footnote"], role="micro_faint",
               color=st.c("muted", d._tone))


def decorate_table(d: Deck, tbl, s: dict, cols, rows):
    st = d.st
    tone = d._tone
    bg = st.c("bg", tone) or st.c("bg", "light")
    hair = st.c("hair", tone)
    head_ink = st.decor.get("table_head_ink")
    head_color = hexrgb(head_ink) if head_ink else st.c("muted", tone)
    ink = st.c("ink", tone)
    muted = st.c("muted", tone)
    head_bg = st.decor.get("table_head_bg")
    zebra = st.decor.get("table_zebra")
    bold_cols = set(s.get("bold_cols", []))
    muted_cols = set(s.get("muted_cols", []))

    def style_cell(cell, val, size, color, bold, fill_rgb):
        cell.fill.solid()
        cell.fill.fore_color.rgb = fill_rgb
        cell.margin_left = Inches(0.02)
        cell.margin_right = Inches(0.10)
        cell.margin_top = Inches(0)
        cell.margin_bottom = Inches(0)
        cell.vertical_anchor = MSO_ANCHOR.MIDDLE
        tf = cell.text_frame
        tf.word_wrap = True
        p = tf.paragraphs[0]
        p.alignment = PP_ALIGN.LEFT
        d._rich(p, str(val), size, color, bold, 0)
        tcPr = cell._tc.get_or_add_tcPr()
        for tag in ("a:lnL", "a:lnR", "a:lnT", "a:lnB"):
            old = tcPr.find(qn(tag))
            if old is not None:
                tcPr.remove(old)
        for tag in ("a:lnL", "a:lnR", "a:lnT", "a:lnB"):
            ln = etree.SubElement(tcPr, qn(tag))
            if tag == "a:lnB":
                ln.set("w", "6350")
                sf = etree.SubElement(ln, qn("a:solidFill"))
                etree.SubElement(sf, qn("a:srgbClr")).set(
                    "val", "%02X%02X%02X" % (hair[0], hair[1], hair[2]))
            else:
                etree.SubElement(ln, qn("a:noFill"))

    head_fill = hexrgb(head_bg) if head_bg else bg
    for i, h in enumerate(cols):
        style_cell(tbl.cell(0, i), h, st.size("micro", 10.5), head_color, True, head_fill)
    z = hexrgb(zebra) if zebra else None
    for r, row in enumerate(rows, start=1):
        for i, val in enumerate(row):
            fill = z if (z and r % 2 == 0) else bg
            if i in bold_cols:
                style_cell(tbl.cell(r, i), val, st.size("body_sm", 12.5), ink, True, fill)
            elif i in muted_cols:
                style_cell(tbl.cell(r, i), val, st.size("body_sm", 11), muted, False, fill)
            else:
                style_cell(tbl.cell(r, i), val, st.size("body_sm", 11), muted, False, fill)


def slide_closing(d: Deck, s: dict):
    st = d.st
    d.new_slide(s.get("tone", "dark"))
    y = float(s.get("title_y", 2.62))
    align = s.get("align", "center")
    x = d.M if align == "left" else d.M
    w = d.CW
    d.text(x, y, w, s.get("title", ""), role="title", bold=True, align=align)
    y = d.disp_bottom(y, st.size("title", 52)) + 0.30
    if s.get("subtitle"):
        d.text(x, y, w, s["subtitle"], role="lead", color=st.c("muted", d._tone), align=align)
        y += d.disp_bottom(0, st.size("lead", 19)) + 0.24
    if s.get("link"):
        d.text(x, y + 0.08, w, s["link"], role="body_sm", bold=True,
               color=st.c("accent", d._tone), align=align)


ARCHETYPES = {
    "cover": slide_cover,
    "statement": slide_statement,
    "bigstat": slide_bigstat,
    "stats": slide_stats,
    "cards": slide_cards,
    "lanes": slide_lanes,
    "timeline": slide_timeline,
    "bullets": slide_bullets,
    "table": slide_table,
    "closing": slide_closing,
}


def build(spec: dict, style: Style, out_path: str) -> Deck:
    deck = Deck(style, float(spec.get("width", 13.3333)), float(spec.get("height", 7.5)))
    for i, s in enumerate(spec.get("slides", [])):
        kind = s.get("type")
        fn = ARCHETYPES.get(kind)
        if not fn:
            raise SystemExit(f"第 {i+1} 页：未知版式 '{kind}'。可用：{', '.join(ARCHETYPES)}")
        fn(deck, s)
    deck.save(out_path)
    return deck
