# -*- coding: utf-8 -*-
"""deck.json + 主题包 -> 可编辑的 .pptx。

双轨渲染
--------
* **克隆轨**：每页先从 stencil.pptx 克隆一张「外壳页」，背景、装饰、校徽、
  底纹全部是模板原件，保真。
* **生成轨**：导航条和全部正文用 python-pptx 原生形状画，读 theme.json 的
  token 取色取字。于是同一份 deck.json 换个 --theme 就能整套换皮，
  而模板里没有的版式（时间轴、KPI、对比）也一样画得出来。

用法：
    python build_deck.py deck.json --out 汇报.pptx [--theme jnu-teal]
"""
from __future__ import annotations

import argparse
import json
import sys
from datetime import datetime, timezone
from pathlib import Path

from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR
from pptx.chart.data import ChartData
from pptx.enum.chart import (XL_CHART_TYPE, XL_LABEL_POSITION, XL_LEGEND_POSITION,
                             XL_TICK_LABEL_POSITION)
from pptx.oxml.ns import qn
from pptx.util import Pt

sys.path.insert(0, str(Path(__file__).parent))
from ppt_kit import (  # noqa: E402
    Theme, autofit_size, clone_slide, drop_slides, estimate_lines, inches,
    block_height, line, line_height_in, picture, rect, rgb, sanitize_filename,
    set_text, text_width_units, textbox,
)
from pptx import Presentation  # noqa: E402

ROOT = Path(__file__).resolve().parent.parent
CJK_NUM = "一二三四五六七八九十"


# ==========================================================================
class Builder:
    def __init__(self, deck: dict, theme: Theme):
        self.deck = deck
        self.t = theme
        self.prs = Presentation(str(theme.stencil))
        self.shell_idx = json.loads((theme.dir / "shells.json").read_text(encoding="utf-8"))["shells"]
        self.n_shells = len(self.prs.slides)
        self.chapters = deck.get("chapters") or self._derive_chapters()
        self.cur_chapter = 0  # 1-based；0 = 封面/目录这类无章节页
        self.warnings: list[str] = []
        self.missing_assets: list[Path] = []
        self.assets = Path(deck.get("asset_dir") or ".")
        # 画布缩放：模板 2/4/1 是 13.33×7.5 英寸，模板 3/5 是 20×11.25。
        # 版式代码里的内边距/间距都按 13.33 写，乘上 k 才能在大画布上保持同样的视觉比例。
        self.k = self.t.geometry["canvas"]["w"] / 13.333

    def p(self, v: float) -> float:
        """把「13.33 英寸画布下的尺寸」换算到当前画布。"""
        return v * self.k

    def _breathe(self, items, w, box_h, size, ls, base_gap, target=0.80, cap=2.4):
        """内容撑不满正文框时，把段间距按比例拉开（最多 cap 倍）。
        与其让四条要点挤在两倍高的框顶端，不如让它们均匀呼吸。"""
        if len(items) < 2:
            return base_gap
        need = block_height(items, w, size, ls, base_gap)
        if need >= box_h * target:
            return base_gap
        extra_in = box_h * target - need
        add_pt = extra_in * 72.0 / (len(items) - 1)
        return min(base_gap * cap, base_gap + add_pt)

    def balance(self, t, box_h, need_h):
        """正文框往往比内容高得多，纯顶对齐会让整页重心浮在上半部分
        （20 英寸画布上尤其明显，实测下方能空掉 40%）。这里把多余空间
        按主题声明的比例分一部分到上方，让版心落回视觉中心附近。
        0 = 顶对齐，0.5 = 垂直居中。

        返回 (新上边, 新高度)——高度必须同步收，否则框底会穿出画布。"""
        f = self.t.data.get("body_balance", 0.0)
        if f <= 0 or need_h <= 0:
            return t, box_h
        off = max(0.0, box_h - need_h) * f
        return t + off, box_h - off

    # ---------------- 基础 ----------------
    def _derive_chapters(self):
        return [{"title": s.get("title", "")} for s in self.deck["slides"]
                if s.get("layout") == "section"]

    def ts(self, key, default=18):
        return self.t.data.get("type_scale", {}).get(key, default)

    def g(self, key):
        return self.t.box(key)

    def shell(self, role):
        sid = self.t.data.get("shell_use", {}).get(role, "base")
        idx = self.shell_idx.get(sid, self.shell_idx.get("base", 0))
        return clone_slide(self.prs, self.prs.slides[idx])

    def notes(self, slide, text):
        if text:
            slide.notes_slide.notes_text_frame.text = str(text)

    # ---------------- 导航 ----------------
    def draw_nav(self, slide, active: int):
        self._header_label_used = False
        nav = self.t.nav
        kind = nav.get("kind", "none")
        if kind == "none" or not self.chapters:
            return
        titles = [c.get("title", "") for c in self.chapters]
        if kind == "rail-left":
            self._nav_rail(slide, titles, active, nav)
        elif kind == "tabs-top":
            self._nav_tabs(slide, titles, active, nav)
        elif kind == "bar-top":
            self._nav_bar(slide, titles, active, nav)
        elif kind == "header-label":
            self._nav_header(slide, titles, active, nav)

    def _nav_rail(self, slide, titles, active, nav):
        n = len(titles)
        tile, span = nav["tile"], nav["span"]
        avail = span["b"] - span["t"]
        h = tile["h"]
        if n > nav.get("max_items", 6):
            h = max(0.42, (avail - 0.12 * (n - 1)) / n)
        # 保持原模板的节奏：牌子之间留固定间距，整块居中；章节少时不要拉得太散
        gap = min(nav.get("gap", 0.39), (avail - n * h) / (n - 1)) if n > 1 else 0
        total = n * h + max(0, n - 1) * gap
        top = span["t"] + (avail - total) / 2

        for i in range(n):
            on = (i + 1) == active
            st = nav["active"] if on else nav["inactive"]
            y = top + i * (h + gap)
            box = rect(slide, tile["l"], y, tile["w"], h,
                       fill=self.t.c(st["fill"]) if st["fill"] else None)
            box.name = f"nav_{i+1}"
            label = CJK_NUM[i] if i < len(CJK_NUM) and nav.get("numerals") == "cjk" else str(i + 1)
            # 导航牌上也是数字，跟章节号/KPI 用同一套字面；当前章靠颜色和字号区分，
            # 不靠字重（numeral 字面已把字重钉死）
            set_text(box.text_frame, label, self.t, role="numeral",
                     size=min(st["size"], h * 40), color=st["text"], bold=st["bold"],
                     align="center", anchor="middle", space_after=0, line_spacing=1.0)

    def _nav_tabs(self, slide, titles, active, nav):
        n = min(len(titles), nav.get("max_items", 5))
        bar, ul, span = nav["bar"], nav["underline"], nav["span"]
        cw = self.t.geometry["canvas"]["w"]
        rect(slide, 0, bar["t"], cw, bar["h"], fill=self.t.c(bar["fill"])).name = "nav_bar"
        rect(slide, 0, ul["t"], cw, ul["h"], fill=self.t.c(ul["fill"])).name = "nav_rule"

        avail = span["r"] - span["l"]
        w = (avail - nav["tab_gap"] * (n - 1)) / n
        for i in range(n):
            on = (i + 1) == active
            st = nav["active"] if on else nav["inactive"]
            x = span["l"] + i * (w + nav["tab_gap"])
            if st["fill"]:
                sh = rect(slide, x, bar["t"] - 0.06, w, nav["tab_h"],
                          fill=self.t.c(st["fill"]), shape=MSO_SHAPE.ROUNDED_RECTANGLE,
                          adj=nav.get("radius", 0.14))
            else:
                sh = rect(slide, x, bar["t"] - 0.06, w, nav["tab_h"], fill=None)
            sh.name = f"nav_{i+1}"
            set_text(sh.text_frame, self._short(titles[i], self._fit_units(w, st["size"])),
                     self.t, role="heading",
                     size=st["size"], color=st["text"], bold=st["bold"],
                     align="center", anchor="middle", space_after=0, line_spacing=1.0)

    def _nav_bar(self, slide, titles, active, nav):
        n = min(len(titles), nav.get("max_items", 5))
        span = nav["span"]
        w = (span["r"] - span["l"]) / n
        for i in range(n):
            on = (i + 1) == active
            st = nav["active"] if on else nav["inactive"]
            # 主题给的 span 高度可能比一行字的行盒还矮（jnu-rigor：20pt 字配 0.32in），
            # 以 span 中线为准把框撑到放得下一行，视觉位置不变
            bh = max(span["h"], line_height_in(st["size"], 1.0) * 1.05)
            by = span["t"] + (span["h"] - bh) / 2
            box = textbox(slide, span["l"] + i * w, by, w - 0.2, bh)
            box.name = f"nav_{i+1}"
            set_text(box.text_frame,
                     self._short(titles[i], self._fit_units(w - 0.2, st["size"])),
                     self.t, role="heading",
                     size=st["size"], color=st["text"], bold=st["bold"],
                     align="left", anchor="middle", space_after=0, line_spacing=1.0)

    def _nav_header(self, slide, titles, active, nav):
        if not (1 <= active <= len(titles)):
            return
        # 这类主题的章节标签就占在页眉的小字位上，页眉不能再画 kicker，
        # 否则两段字会精确地叠在一起。
        self._header_label_used = True
        lb = nav["label"]
        st = nav["active"]
        # 「第N部分」本身要占掉约 1.1in，剩下的宽度才归章节名
        room = self._fit_units(lb["w"] - self.p(1.1), st["size"])
        label = "第%s部分   %s" % (CJK_NUM[active - 1], self._short(titles[active - 1], room))
        box = textbox(slide, lb["l"], lb["t"], lb["w"], lb["h"])
        box.name = "nav_label"
        set_text(box.text_frame, label, self.t, role="heading", size=st["size"],
                 color=st["text"], bold=st["bold"], anchor="middle",
                 space_after=0, line_spacing=1.0)

    @staticmethod
    def _short(s, cap_units):
        """按**字宽**截断，不是按字符数。

        「Transformer 怎么搭」有 16 个字符，但西文窄，实际只占 11 个全角字的
        宽度；按字符数截会白白砍掉一半还空着的导航条。"""
        s = (s or "").strip()
        if text_width_units(s) <= cap_units:
            return s
        used, out = 0.0, []
        for ch in s:
            w = 1.0 if ord(ch) > 0x2E7F else 0.55
            if used + w > cap_units - 1.0:      # 给省略号留位
                break
            used += w
            out.append(ch)
        return "".join(out).rstrip() + "…"

    @staticmethod
    def _fit_units(box_w_in, size_pt, margin=0.9):
        """这个宽度、这个字号大约能放几个全角字宽。"""
        return max(3.0, box_w_in * margin / (size_pt / 72.0))

    # ---------------- 页眉 ----------------
    def page_head(self, slide, title, kicker=None, rule=True, color="ink",
                  kicker_color="muted", rule_color="line"):
        if getattr(self, "_header_label_used", False):
            kicker = None
        if kicker:
            l, t, w, h = self.g("kicker")
            set_text(textbox(slide, l, t, w, h).text_frame, kicker, self.t,
                     role="body", size=self.ts("kicker", 12), color=kicker_color,
                     space_after=0, line_spacing=1.0)
        if title:
            l, t, w, h = self.g("title")
            size = autofit_size([title], w, h, self.ts("title", 24),
                                self.ts("title", 24) * 0.58, line_spacing=1.1,
                                max_lines=2)
            set_text(textbox(slide, l, t, w, h).text_frame, title, self.t,
                     role="heading", size=size, color=color, bold=True,
                     anchor="middle", space_after=0, line_spacing=1.1)
        if rule:
            l, t, w, _ = self.g("rule")
            line(slide, l, t, w, 0, self.t.c(rule_color), 1.25)

    # ---------------- 各版式 ----------------
    def s_cover(self, slide, s):
        d = self.deck.get("meta", {})
        l, t, w, h = self.g("cover_title")
        col = s.get("color", "white")
        title = s.get("title") or d.get("title", "")
        size = autofit_size([title], w, h, self.ts("cover_title", 40),
                            self.ts("cover_title", 40) * 0.42, line_spacing=1.15,
                            max_lines=s.get("max_lines", 2))
        role = "display" if s.get("display_font", True) else "heading"
        set_text(textbox(slide, l, t, w, h).text_frame, title, self.t, role=role,
                 size=size, color=col, bold=True, anchor="middle", line_spacing=1.15,
                 align=s.get("align", "left"))

        sub = s.get("subtitle") or d.get("subtitle")
        if sub:
            l, t, w, h = self.g("cover_sub")
            set_text(textbox(slide, l, t, w, h).text_frame, sub, self.t, role="body",
                     size=self.ts("cover_sub", 18), color=s.get("sub_color", col),
                     anchor="middle", align=s.get("align", "left"))

        meta = s.get("meta") or self._meta_lines(d)
        if meta:
            l, t, w, h = self.g("cover_meta")
            # 学院+专业写全时署名行会折成两行，框高撑不住，按框缩字号
            lines = [meta] if isinstance(meta, str) else [str(x) for x in meta]
            base = self.ts("cover_meta", 14)
            size = autofit_size(lines, w, h, base, base * 0.75,
                                line_spacing=1.5, space_after_pt=2)
            set_text(textbox(slide, l, t, w, h).text_frame, meta, self.t, role="body",
                     size=size, color=s.get("meta_color", "ink"),
                     align=s.get("meta_align", "center"), anchor="middle",
                     line_spacing=1.5, space_after=2)

    def _meta_lines(self, d):
        """按主题声明的统一格式生成封面署名；meta_lines 可完全覆盖。"""
        if d.get("meta_lines"):
            return [str(x) for x in d["meta_lines"]]
        fmt = self.t.data.get("meta_format", {})
        sep = fmt.get("affiliation_separator", "｜")
        out = []
        presenter = d.get("presenter")
        affiliation = d.get("affiliation")
        if presenter:
            label = fmt.get("presenter_label", "汇报人")
            who = f"{label}：{presenter}"
            if affiliation:
                who += f" {sep} {affiliation}"
            out.append(who)
        elif affiliation:
            out.append(str(affiliation))
        if d.get("advisor"):
            out.append(f"{fmt.get('advisor_label', '指导教师')}：{d['advisor']}")
        if d.get("date"):
            prefix = f"{fmt.get('date_label', '日期')}：" if fmt.get("show_date_label") else ""
            out.append(prefix + str(d["date"]))
        return out

    def s_toc(self, slide, s):
        k = self.k
        if self.t.data.get("nav_on_toc", True):
            self.draw_nav(slide, 0)
        col = s.get("color", "ink")
        on_dark = col in ("white", "#FFFFFF", "#ffffff")
        sub_col = "secondary" if on_dark else "muted"
        rule_col = "secondary" if on_dark else "line"
        self.page_head(slide, s.get("title", "目录"), kicker=s.get("kicker", "CONTENTS"),
                       color=col, kicker_color=sub_col, rule_color=rule_col)
        # 有些模板的目录外壳上有校徽/水印会压到正文区，主题可以单独声明
        # toc_body 把目录挪开；没声明就还用通用正文区。
        l, t, w, h = self.g("toc_body") if "toc_body" in self.t.geometry else self.g("body")
        items = self.chapters
        n = max(1, len(items))
        cols = 2 if n > 4 else 1
        rows = (n + cols - 1) // cols
        gapx = self.p(0.5)
        cw = (w - gapx * (cols - 1)) / cols
        rh = min(h / rows, self.p(1.28))
        num_w, num_gap = self.p(0.95), self.p(1.05)
        num_col = self.t.c("secondary" if on_dark else "primary")
        for i, ch in enumerate(items):
            r, c = i % rows, i // rows
            x, y = l + c * (cw + gapx), t + r * rh
            nb = textbox(slide, x, y, num_w, rh * 0.8)
            set_text(nb.text_frame, "%02d" % (i + 1), self.t, role="numeral",
                     size=self.ts("toc_num", 30), color=num_col, bold=True,
                     anchor="middle", space_after=0, line_spacing=1.0)
            tb = textbox(slide, x + num_gap, y, cw - num_gap, rh * 0.8)
            lines = [(ch.get("title", ""), self.ts("lead", 18), True, col)]
            if ch.get("title_en"):
                lines.append((ch["title_en"], self.ts("caption", 11), False, sub_col))
            set_text(tb.text_frame, lines, self.t, role="body", anchor="middle",
                     space_after=1, line_spacing=1.2)
            line(slide, x + num_gap, y + rh * 0.82, cw - num_gap, 0, self.t.c(rule_col), 0.75)

    def s_section(self, slide, s, index):
        self.draw_nav(slide, index)
        l, t, w, h = self.g("section_no")
        set_text(textbox(slide, l, t, w, h).text_frame, "%02d" % index, self.t,
                 role="numeral", size=self.ts("section_no", 54),
                 color=s.get("no_color", "secondary"), bold=True, anchor="middle",
                 align=s.get("no_align", "left"), space_after=0, line_spacing=1.0)
        l, t, w, h = self.g("section_title")
        title = s.get("title", "")
        size = autofit_size([title], w, h, self.ts("section_title", 32),
                            self.ts("section_title", 32) * 0.45, line_spacing=1.1,
                            max_lines=s.get("max_lines", 1))
        set_text(textbox(slide, l, t, w, h).text_frame, title, self.t, role="heading",
                 size=size, color=s.get("color", "ink"), bold=True, anchor="middle",
                 line_spacing=1.1, align=s.get("align", "left"))
        sub = s.get("subtitle") or s.get("title_en")
        if sub:
            l, t, w, h = self.g("section_sub")
            on_dark = s.get("color") in ("white", "#FFFFFF")
            set_text(textbox(slide, l, t, w, h).text_frame, sub, self.t, role="body",
                     size=self.ts("small", 14),
                     color=s.get("sub_color", "secondary" if on_dark else "muted"),
                     anchor="middle", align=s.get("align", "left"))

    def s_bullets(self, slide, s):
        l, t, w, h = self.g("body")
        items = [str(x) for x in (s.get("bullets") or [])]
        lead = s.get("lead")
        lead_h = self.p(0.85) + self.p(0.12) if lead else 0.0
        if not items:
            if lead:
                set_text(textbox(slide, l, t, w, self.p(0.85)).text_frame, lead, self.t,
                         role="body", size=self.ts("lead", 18), color="primary",
                         bold=True, anchor="middle", line_spacing=1.25)
            return

        avail = h - lead_h
        size = autofit_size(items, w - self.p(0.45), avail, self.ts("body", 17),
                            self.ts("min_body", 12), line_spacing=1.45, space_after_pt=10)
        gap_pt = self._breathe(items, w - self.p(0.45), avail, size, 1.45, 10)
        need = block_height(items, w - self.p(0.45), size, 1.45, gap_pt)
        # 结论句和要点是一整块，配重要一起挪——否则结论贴在标题下面，
        # 要点掉到下半页，中间裂开一道空白。
        t, h = self.balance(t, h, lead_h + need)
        if lead:
            set_text(textbox(slide, l, t, w, self.p(0.85)).text_frame, lead, self.t,
                     role="body", size=self.ts("lead", 18), color="primary", bold=True,
                     anchor="middle", line_spacing=1.25)
            t, h = t + lead_h, h - lead_h
        set_text(textbox(slide, l, t, w, h).text_frame, items, self.t, role="body",
                 size=size, color="ink", bullet=s.get("bullet_char", "▪"),
                 line_spacing=1.45, space_after=gap_pt)

    def s_two_col(self, slide, s):
        l, t, w, h = self.g("body")
        gap = self.p(0.5)
        cw = (w - gap) / 2
        cols = s.get("columns") or [{}, {}]
        head_h = self.p(0.5) + self.p(0.22)
        has_head = any(c.get("title") for c in cols[:2])
        avail = h - (head_h if has_head else 0)

        # 两栏必须共用一个字号：各自独立缩放会出现左栏 17pt、右栏 15pt 的参差
        shared = self.ts("body", 17)
        for col in cols[:2]:
            items = [str(v) for v in (col.get("bullets") or [])]
            if items:
                shared = min(shared, autofit_size(
                    items, cw - self.p(0.4), avail, self.ts("body", 17),
                    self.ts("min_body", 12), line_spacing=1.4, space_after_pt=8))
        need = 0.0
        for col in cols[:2]:
            items = [str(v) for v in (col.get("bullets") or [])]
            if items:
                need = max(need, block_height(items, cw - self.p(0.4), shared, 1.4, 8))
        if need:
            t, h = self.balance(t, h, need + (head_h if has_head else 0))
        for i, col in enumerate(cols[:2]):
            x = l + i * (cw + gap)
            y, ch = t, h
            if col.get("title"):
                hh = self.p(0.5)
                set_text(textbox(slide, x, y, cw, hh).text_frame, col["title"], self.t,
                         role="heading", size=self.ts("card_title", 17), color="primary",
                         bold=True, anchor="middle", space_after=0)
                line(slide, x, y + hh + self.p(0.02), cw, 0, self.t.c("primary"), 1.75)
                y, ch = y + hh + self.p(0.22), ch - hh - self.p(0.22)
            items = [str(v) for v in (col.get("bullets") or [])]
            if items:
                set_text(textbox(slide, x, y, cw, ch).text_frame, items, self.t,
                         role="body", size=shared, color="ink",
                         bullet=s.get("bullet_char", "▪"), line_spacing=1.4, space_after=8)
            elif col.get("text"):
                set_text(textbox(slide, x, y, cw, ch).text_frame, col["text"], self.t,
                         role="body", size=self.ts("body", 17), color="ink", line_spacing=1.45)

    def s_cards(self, slide, s):
        l, t, w, h = self.g("body")
        cards = s.get("cards") or []
        n = max(1, len(cards))
        gap = self.p(0.32)
        pad = self.p(0.26)
        cw = (w - gap * (n - 1)) / n
        head_h = self.p(1.55)
        need = 0.0
        sz = self.ts("card_body", 13.5)
        for c in cards:
            items = c.get("bullets") or ([c["text"]] if c.get("text") else [])
            lines = sum(estimate_lines(str(b), cw - 2 * pad, sz) for b in items)
            need = max(need, lines * sz * 1.4 / 72.0 + self.p(0.09) * max(0, len(items) - 1))
        ch = max(self.p(2.1), min(h, head_h + need + self.p(0.30)))
        top = t + (h - ch) / 2
        for i, c in enumerate(cards):
            x = l + i * (cw + gap)
            card = rect(slide, x, top, cw, ch, fill=self.t.c(c.get("fill", "surface")),
                        line=self.t.c("line"), shape=MSO_SHAPE.ROUNDED_RECTANGLE, adj=0.05)
            card.name = f"card_{i+1}"
            tag = c.get("tag") or "%02d" % (i + 1)
            set_text(textbox(slide, x + pad, top + self.p(0.20), cw - 2 * pad,
                             self.p(0.42)).text_frame, tag, self.t, role="heading",
                     size=self.ts("card_tag", 18),
                     color=c.get("tag_color", "primary"), bold=True, space_after=0)
            set_text(textbox(slide, x + pad, top + self.p(0.68), cw - 2 * pad,
                             self.p(0.62)).text_frame, c.get("title", ""), self.t,
                     role="heading", size=self.ts("card_title", 17), color="ink",
                     bold=True, line_spacing=1.2, space_after=0)
            body_items = c.get("bullets") or ([c["text"]] if c.get("text") else [])
            if body_items:
                bh = ch - self.p(1.50)
                size = autofit_size([str(b) for b in body_items], cw - 2 * pad, bh, sz,
                                    self.ts("min_body", 12) - 1, line_spacing=1.4,
                                    space_after_pt=6)
                set_text(textbox(slide, x + pad, top + self.p(1.36), cw - 2 * pad,
                                 bh).text_frame, [str(b) for b in body_items], self.t,
                         role="body", size=size, color="muted", line_spacing=1.4,
                         space_after=6, bullet="·" if len(body_items) > 1 else None)

    def _place_image(self, slide, s, key, x, y, w, h):
        """放图；找不到就画占位框并记一条警告，绝不静默跳过。"""
        img = s.get(key)
        if not img:
            return
        p = Path(img) if Path(img).is_absolute() else (self.assets / img)
        if p.exists():
            picture(slide, p, x, y, w, h, mode=s.get("image_mode", "fit"))
        else:
            ph = rect(slide, x, y, w, h, fill=self.t.c("surface"), line=self.t.c("line"))
            set_text(ph.text_frame, f"[缺图] {img}", self.t, role="body",
                     size=self.ts("small", 14), color="muted", align="center",
                     anchor="middle", space_after=0)
            self.warnings.append(f"找不到图片：{p}")
            self.missing_assets.append(p)

    def s_image_text(self, slide, s):
        l, t, w, h = self.g("body")
        gap = self.p(0.45)
        ratio = float(s.get("image_ratio", 0.5))
        img_w = w * ratio
        side = s.get("image_side", "right")
        ix = l + (w - img_w) if side == "right" else l
        tx = l if side == "right" else l + img_w + gap
        tw = w - img_w - gap

        cap = s.get("caption")
        cap_h = self.p(0.42) if cap else 0
        self._place_image(slide, s, "image", ix, t, img_w, h - cap_h)
        if cap:
            set_text(textbox(slide, ix, t + h - cap_h + self.p(0.06), img_w,
                             self.p(0.34)).text_frame, cap, self.t, role="body",
                     size=self.ts("caption", 11), color="muted", align="center",
                     space_after=0)

        lead = s.get("lead")
        lead_h = self.p(0.80) + self.p(0.12) if lead else 0.0
        items = [str(x) for x in (s.get("bullets") or [])]
        t2, h2 = t, h
        if items:
            size = autofit_size(items, tw - self.p(0.4), h2 - lead_h, self.ts("body", 17),
                                self.ts("min_body", 12), line_spacing=1.45, space_after_pt=9)
            gap2 = self._breathe(items, tw - self.p(0.4), h2 - lead_h, size, 1.45, 9)
            need = block_height(items, tw - self.p(0.4), size, 1.45, gap2)
            t2, h2 = self.balance(t2, h2, lead_h + need)
        if lead:
            set_text(textbox(slide, tx, t2, tw, self.p(0.80)).text_frame, lead, self.t,
                     role="body", size=self.ts("lead", 18), color="primary", bold=True,
                     anchor="middle", line_spacing=1.25)
            t2, h2 = t2 + lead_h, h2 - lead_h
        if items:
            set_text(textbox(slide, tx, t2, tw, h2).text_frame, items, self.t,
                     role="body", size=size, color="ink", bullet=s.get("bullet_char", "▪"),
                     line_spacing=1.45, space_after=gap2)
        elif s.get("text"):
            set_text(textbox(slide, tx, t2, tw, h2).text_frame, s["text"], self.t,
                     role="body", size=self.ts("body", 17), color="ink", line_spacing=1.5)

    def s_image_full(self, slide, s):
        l, t, w, h = self.g("body")
        cap = s.get("caption")
        cap_h = self.p(0.45) if cap else 0
        self._place_image(slide, s, "image", l, t, w, h - cap_h)
        if cap:
            set_text(textbox(slide, l, t + h - cap_h + self.p(0.08), w,
                             self.p(0.36)).text_frame, cap, self.t, role="body",
                     size=self.ts("caption", 11), color="muted", align="center",
                     space_after=0)

    def s_compare(self, slide, s):
        l, t, w, h = self.g("body")
        gap, pad = self.p(0.4), self.p(0.28)
        cw = (w - gap) / 2
        pair = s.get("sides") or s.get("columns") or [{}, {}]
        accents = [s.get("left_color", "risk"), s.get("right_color", "ok")]
        head_h = self.p(0.92)
        body_w = cw - 2 * pad
        side_items = [[str(v) for v in (col.get("bullets") or [])] for col in pair[:2]]
        # 两侧共用一个字号，卡片高度跟着内容走：以前卡片固定撑满正文区，
        # 三条短要点只占上方三成，下面大片空白
        size = self.ts("body", 17) - 1
        for items in side_items:
            if items:
                size = min(size, autofit_size(
                    items, body_w, h - head_h - self.p(0.13), self.ts("body", 17) - 1,
                    self.ts("min_body", 12), line_spacing=1.4, space_after_pt=8))
        need = max((block_height(items, body_w, size, 1.4, 8) for items in side_items
                    if items), default=0.0)
        ch = min(h, max(self.p(2.4), head_h + need * 1.08 + self.p(0.35)))
        top = t + (h - ch) / 2
        for i, col in enumerate(pair[:2]):
            x = l + i * (cw + gap)
            acc = self.t.c(col.get("color", accents[i]))
            card = rect(slide, x, top, cw, ch, fill=self.t.c("surface_alt"),
                        line=self.t.c("line"), shape=MSO_SHAPE.ROUNDED_RECTANGLE, adj=0.04)
            card.name = f"cmp_{i+1}"
            rect(slide, x, top, cw, self.p(0.09), fill=acc)
            set_text(textbox(slide, x + pad, top + pad, body_w,
                             self.p(0.5)).text_frame, col.get("title", ""), self.t,
                     role="heading", size=self.ts("card_title", 17), color=acc,
                     bold=True, anchor="middle", space_after=0)
            items = side_items[i]
            if items:
                set_text(textbox(slide, x + pad, top + head_h, body_w,
                                 ch - head_h - self.p(0.13)).text_frame,
                         items, self.t, role="body", size=size, color="ink",
                         bullet="·", line_spacing=1.4, space_after=8)

    def s_timeline(self, slide, s):
        l, t, w, h = self.g("body")
        steps = s.get("steps") or s.get("items") or []
        n = max(1, len(steps))
        axis_y = t + h * 0.42
        line(slide, l + self.p(0.2), axis_y, w - self.p(0.4), 0, self.t.c("line"), 2.0)
        cw = w / n
        dot_r = self.p(0.16)
        for i, st in enumerate(steps):
            cx = l + cw * (i + 0.5)
            dot = rect(slide, cx - dot_r, axis_y - dot_r, dot_r * 2, dot_r * 2,
                       fill=self.t.c("primary"), shape=MSO_SHAPE.OVAL)
            dot.name = f"tl_dot_{i+1}"
            above = i % 2 == 0
            bh = h * 0.36
            by = axis_y - self.p(0.30) - bh if above else axis_y + self.p(0.30)
            bx = cx - cw / 2 + self.p(0.12)
            bw = cw - self.p(0.24)
            lines = []
            if st.get("label") or st.get("time"):
                lines.append((st.get("label") or st.get("time"),
                              self.ts("caption", 11), True, "primary"))
            if st.get("title"):
                lines.append((st["title"], self.ts("card_title", 16), True, "ink"))
            if st.get("text"):
                lines.append((st["text"], self.ts("card_body", 13), False, "muted"))
            set_text(textbox(slide, bx, by, bw, bh).text_frame, lines, self.t,
                     role="body", align="center", anchor="bottom" if above else "top",
                     line_spacing=1.3, space_after=3)

    def s_steps(self, slide, s):
        l, t, w, h = self.g("body")
        steps = s.get("steps") or s.get("items") or []
        n = max(1, len(steps))
        gap, arrow, pad = self.p(0.26), self.p(0.30), self.p(0.22)
        cw = (w - (gap + arrow) * (n - 1)) / n
        ch = min(h * 0.72, self.p(2.6))
        top = t + (h - ch) / 2
        for i, st in enumerate(steps):
            x = l + i * (cw + gap + arrow)
            box = rect(slide, x, top, cw, ch, fill=self.t.c("surface"),
                       line=self.t.c("line"), shape=MSO_SHAPE.ROUNDED_RECTANGLE, adj=0.06)
            box.name = f"step_{i+1}"
            lines = [("STEP %d" % (i + 1), self.ts("caption", 11), True, "primary"),
                     (st.get("title", ""), self.ts("card_title", 16), True, "ink")]
            if st.get("text"):
                lines.append((st["text"], self.ts("card_body", 13), False, "muted"))
            set_text(textbox(slide, x + pad, top + pad, cw - 2 * pad,
                             ch - 2 * pad).text_frame, lines, self.t, role="body",
                     anchor="middle", line_spacing=1.3, space_after=4)
            if i < n - 1:
                ah = self.p(0.26)
                ar = rect(slide, x + cw + gap / 2, top + ch / 2 - ah / 2, arrow, ah,
                          fill=self.t.c("secondary"), shape=MSO_SHAPE.RIGHT_ARROW)
                ar.name = f"step_arrow_{i+1}"

    def s_kpi(self, slide, s):
        l, t, w, h = self.g("body")
        items = s.get("items") or s.get("kpis") or []
        n = max(1, len(items))
        gap, pad = self.p(0.4), self.p(0.25)
        cw = (w - gap * (n - 1)) / n
        ch = min(h * 0.66, self.p(2.8))
        top = t + (h - ch) / 2
        for i, kk in enumerate(items):
            x = l + i * (cw + gap)
            card = rect(slide, x, top, cw, ch, fill=self.t.c("surface"),
                        shape=MSO_SHAPE.ROUNDED_RECTANGLE, adj=0.05)
            card.name = f"kpi_{i+1}"
            val = str(kk.get("value", ""))
            vsize = autofit_size([val], cw - 2 * pad, ch * 0.5,
                                 self.ts("kpi_value", 40),
                                 self.ts("title", 24), line_spacing=1.0)
            set_text(textbox(slide, x + pad, top + self.p(0.22), cw - 2 * pad,
                             ch * 0.46).text_frame, val, self.t, role="numeral",
                     size=vsize, color=kk.get("color", "primary"), bold=True,
                     align="center", anchor="middle", space_after=0, line_spacing=1.0)
            lines = [(kk.get("label", ""), self.ts("card_title", 16), True, "ink")]
            if kk.get("note"):
                lines.append((kk["note"], self.ts("caption", 11), False, "muted"))
            set_text(textbox(slide, x + pad, top + ch * 0.56, cw - 2 * pad,
                             ch * 0.40).text_frame, lines, self.t, role="body",
                     align="center", anchor="top", line_spacing=1.25, space_after=2)

    def s_table(self, slide, s):
        l, t, w, h = self.g("body")
        head = s.get("header") or []
        rows_in = s.get("rows") or []
        ncol = max(len(head), max((len(r) for r in rows_in), default=1))
        nrow = len(rows_in) + (1 if head else 0)
        if nrow == 0:
            return
        rh = min(self.p(0.58), h / nrow)
        th = rh * nrow
        gfx = slide.shapes.add_table(nrow, ncol, inches(l), inches(t + (h - th) / 2),
                                     inches(w), inches(th))
        tbl = gfx.table
        tbl.first_row = bool(head)
        body_pt = self.ts("small", 14)
        data = ([head] if head else []) + rows_in
        for r, row in enumerate(data):
            tbl.rows[r].height = inches(rh)
            is_head = bool(head) and r == 0
            for c in range(ncol):
                cell = tbl.cell(r, c)
                cell.margin_left = cell.margin_right = inches(self.p(0.12))
                cell.margin_top = cell.margin_bottom = inches(self.p(0.04))
                cell.fill.solid()
                cell.fill.fore_color.rgb = rgb(
                    self.t.c("primary") if is_head
                    else (self.t.c("surface_alt") if r % 2 else self.t.c("white")))
                txt = str(row[c]) if c < len(row) else ""
                set_text(cell.text_frame, txt, self.t, role="body", size=body_pt,
                         color="white" if is_head else "ink", bold=is_head,
                         anchor="middle", space_after=0, line_spacing=1.15,
                         align="left" if c == 0 else s.get("align", "left"))
                # 表格单元格不认 text_frame 的 anchor，垂直居中要设在单元格上
                cell.vertical_anchor = MSO_ANCHOR.MIDDLE

    def s_chart(self, slide, s):
        """可编辑的原生图表；只接受用户提供或可追溯的数据。"""
        l, t, w, h = self.g("body")
        cfg = s.get("chart") or s
        kinds = {
            "column": XL_CHART_TYPE.COLUMN_CLUSTERED,
            "bar": XL_CHART_TYPE.BAR_CLUSTERED,
            "line": XL_CHART_TYPE.LINE_MARKERS,
        }
        kind = kinds.get(cfg.get("type", cfg.get("chart_type", "column")),
                         XL_CHART_TYPE.COLUMN_CLUSTERED)
        # 横向条形图在 PowerPoint 里第一个类别画在最下面，和阅读顺序相反；
        # 这里倒序喂数据，让 deck.json 里写的第一项出现在最上面
        flip = (lambda xs: list(xs)[::-1]) if kind == XL_CHART_TYPE.BAR_CLUSTERED \
            else (lambda xs: list(xs))
        data = ChartData()
        data.categories = flip(str(x) for x in (cfg.get("categories") or []))
        normalized_series = []
        for series in cfg.get("series") or []:
            values = flip(int(v) if isinstance(v, float) and v.is_integer() else v
                          for v in (series.get("values") or []))
            normalized_series.append(values)
            data.add_series(str(series.get("name", "系列")), values)
        if not data.categories or not (cfg.get("series") or []):
            self.warnings.append(f"图表《{s.get('title', '')}》缺少 categories 或 series")
            return
        lead = s.get("lead")
        if lead:
            lead_h = self.p(0.85)
            set_text(textbox(slide, l, t, w, lead_h).text_frame, lead, self.t,
                     role="body", size=self.ts("lead", 18), color="primary", bold=True,
                     anchor="middle", line_spacing=1.25)
            t, h = t + lead_h + self.p(0.12), h - lead_h - self.p(0.12)
        cap = s.get("caption")
        cap_h = self.p(0.42) if cap else 0
        chart = slide.shapes.add_chart(
            kind, inches(l), inches(t), inches(w), inches(h - cap_h), data
        ).chart
        # python-pptx 可能生成负数 axId/crossAx。PowerPoint 会容忍，但严格的
        # OOXML 渲染器按 unsignedInt 解析会直接失败，因此统一映射为正整数。
        axis_id_map = {}
        for el in chart._chartSpace.iter():
            if el.tag.endswith(("}axId", "}crossAx")) and el.get("val"):
                value = int(el.get("val"))
                if value < 0:
                    axis_id_map[value] = abs(value)
        for el in chart._chartSpace.iter():
            if el.tag.endswith(("}axId", "}crossAx")) and el.get("val"):
                value = int(el.get("val"))
                if value in axis_id_map:
                    el.set("val", str(axis_id_map[value]))
        # python-pptx 的内嵌 XLSX 用 16 位有效数字序列化浮点数，而图表缓存
        # 默认用 str(float)。例如 78.4 会分别写成 78.40000000000001 与 78.4，
        # 严格查看器会认为缓存过期。这里让缓存遵循同一序列化规则。
        c_ns = "http://schemas.openxmlformats.org/drawingml/2006/chart"
        series_nodes = chart._chartSpace.findall(f".//{{{c_ns}}}ser")
        for series_node, values in zip(series_nodes, normalized_series):
            cache = series_node.find(
                f"{{{c_ns}}}val/{{{c_ns}}}numRef/{{{c_ns}}}numCache")
            if cache is None:
                cache = series_node.find(
                    f"{{{c_ns}}}yVal/{{{c_ns}}}numRef/{{{c_ns}}}numCache")
            if cache is None:
                continue
            points = sorted(cache.findall(f"{{{c_ns}}}pt"),
                            key=lambda el: int(el.get("idx", "0")))
            for point, value in zip(points, values):
                node = point.find(f"{{{c_ns}}}v")
                if node is not None and isinstance(value, (int, float)):
                    node.text = str(value) if isinstance(value, int) else format(value, ".16g")
        chart.has_legend = bool(cfg.get("show_legend",
                                        cfg.get("legend", len(cfg.get("series") or []) > 1)))
        if chart.has_legend:
            chart.legend.position = XL_LEGEND_POSITION.BOTTOM
            chart.legend.include_in_layout = False
        chart.has_title = False
        chart.value_axis.has_major_gridlines = True
        self._style_chart(chart, kind)
        self._chart_labels(chart, kind, cfg, normalized_series)
        if cfg.get("value_min") is not None:
            chart.value_axis.minimum_scale = float(cfg["value_min"])
        if cfg.get("value_max") is not None:
            chart.value_axis.maximum_scale = float(cfg["value_max"])
        for axis, key in ((chart.category_axis, "x_title"), (chart.value_axis, "y_title")):
            if cfg.get(key):
                axis.has_title = True
                tf = axis.axis_title.text_frame
                tf.text = str(cfg[key])
                for run in tf.paragraphs[0].runs:
                    run.font.size = Pt(self.ts("caption", 11))
                    run.font.bold = False
                    run.font.color.rgb = rgb(self.t.c("muted"))
        if cap:
            set_text(textbox(slide, l, t + h - cap_h + self.p(0.06), w,
                             self.p(0.34)).text_frame, cap, self.t, role="body",
                     size=self.ts("caption", 11), color="muted", align="center",
                     space_after=0)

    CHART_SERIES_COLORS = ["primary", "secondary", "accent", "primary_dark", "muted"]

    def _style_chart(self, chart, kind):
        """让原生图表跟随主题。python-pptx 的默认样式是 Office 彩虹色：
        单系列柱状图每根柱子一个颜色、黑色网格线、宋体/Calibri 轴文字，
        放进暨大主题里非常突兀。"""
        plot = chart.plots[0]
        plot.vary_by_categories = False     # 单系列不按类别变色
        for i, ser in enumerate(plot.series):
            col = rgb(self.t.c(self.CHART_SERIES_COLORS[i % len(self.CHART_SERIES_COLORS)]))
            if kind == XL_CHART_TYPE.LINE_MARKERS:
                ser.format.line.color.rgb = col
                ser.format.line.width = Pt(2.25)
                ser.marker.format.fill.solid()
                ser.marker.format.fill.fore_color.rgb = col
                ser.marker.format.line.color.rgb = col
            else:
                ser.format.fill.solid()
                ser.format.fill.fore_color.rgb = col
                # PowerPoint 默认把负值柱子反色成白底黑框，看起来像空柱
                ser.invert_if_negative = False
        grid = chart.value_axis.major_gridlines.format.line
        grid.color.rgb = rgb(self.t.c("line"))
        grid.width = Pt(0.75)
        for axis in (chart.value_axis, chart.category_axis):
            axis.format.line.color.rgb = rgb(self.t.c("line"))
            axis.tick_labels.font.size = Pt(self.ts("caption", 11))
            axis.tick_labels.font.color.rgb = rgb(self.t.c("muted"))
        # 全图默认字体：latin + 东亚字形都指定，否则中文轴标题会回退成宋体
        spec = self.t.font("body")
        latin = spec.get("latin") or spec.get("ea")
        ea = spec.get("ea") or latin
        cf = chart.font
        cf.size = Pt(self.ts("caption", 11))
        cf.color.rgb = rgb(self.t.c("muted"))
        if latin:
            cf.name = latin
        if ea:
            rpr = cf._rPr
            el = rpr.find(qn("a:ea"))
            if el is None:
                el = rpr.makeelement(qn("a:ea"), {})
                latin_el = rpr.find(qn("a:latin"))
                if latin_el is not None:
                    latin_el.addnext(el)    # schema 要求 ea 紧跟在 latin 后面
                else:
                    rpr.append(el)
            el.set("typeface", ea)
        if chart.has_legend:
            chart.legend.font.size = Pt(self.ts("caption", 11))

    def _chart_labels(self, chart, kind, cfg, series_values):
        """数值标签与负值处理。

        有负值（比如“相对基线的变化量”）时，类别标签默认贴着零线画，
        会压在朝负方向的柱子上，所以挪到坐标轴最低端。"""
        if any(isinstance(v, (int, float)) and v < 0
               for vals in series_values for v in vals):
            chart.category_axis.tick_label_position = XL_TICK_LABEL_POSITION.LOW
        # 「相对基线的变化量」这类图：下降用 risk、提升用 ok（语义色，见 slide-rules.md）
        if cfg.get("sign_colors") and kind != XL_CHART_TYPE.LINE_MARKERS:
            for ser, vals in zip(chart.plots[0].series, series_values):
                for i, v in enumerate(vals):
                    if isinstance(v, (int, float)) and v != 0:
                        fill = ser.points[i].format.fill
                        fill.solid()
                        fill.fore_color.rgb = rgb(self.t.c("risk" if v < 0 else "ok"))
                # 单点格式 c:dPt 不继承系列上的 invertIfNegative=0，PowerPoint 会把
                # 负值点又反成白底黑框，必须在每个 dPt 里（紧跟 c:idx）再写一次
                for dpt in ser._element.findall(qn("c:dPt")):
                    if dpt.find(qn("c:invertIfNegative")) is None:
                        inv = dpt.makeelement(qn("c:invertIfNegative"), {"val": "0"})
                        dpt.find(qn("c:idx")).addnext(inv)
        fmt = cfg.get("number_format")
        if fmt:
            chart.value_axis.tick_labels.number_format = fmt
            chart.value_axis.tick_labels.number_format_is_linked = False
        if cfg.get("data_labels"):
            plot = chart.plots[0]
            plot.has_data_labels = True
            labels = plot.data_labels
            labels.number_format = fmt or "General"
            labels.number_format_is_linked = False
            labels.position = (XL_LABEL_POSITION.ABOVE if kind == XL_CHART_TYPE.LINE_MARKERS
                               else XL_LABEL_POSITION.OUTSIDE_END)
            labels.font.size = Pt(self.ts("caption", 11))
            labels.font.color.rgb = rgb(self.t.c("ink"))

    def s_quote(self, slide, s):
        l, t, w, h = self.g("body")
        bar_w, inset = self.p(0.10), self.p(0.30)
        rect(slide, l + inset, t + h * 0.12, bar_w, h * 0.76, fill=self.t.c("primary"))
        tx = l + inset + bar_w + self.p(0.35)
        tw = w - (tx - l) - self.p(0.3)
        text = s.get("text") or s.get("quote") or ""
        size = autofit_size([text], tw, h * 0.6, self.ts("quote", 28),
                            self.ts("body", 17), line_spacing=1.4)
        set_text(textbox(slide, tx, t + h * 0.12, tw, h * 0.62).text_frame, text,
                 self.t, role="heading", size=size, color="ink", bold=True,
                 anchor="middle", line_spacing=1.4)
        if s.get("source"):
            set_text(textbox(slide, tx, t + h * 0.78, tw, self.p(0.4)).text_frame,
                     "—— " + s["source"], self.t, role="body",
                     size=self.ts("small", 14), color="muted")

    def s_closing(self, slide, s):
        title = s.get("title")
        # 有的结尾外壳自带「谢谢观看」艺术字（jnu-crisp），标题位已经被占了，
        # 再画一行就会正好压在上面。这是外壳的物理约束，不是风格偏好，
        # 所以由主题覆盖 deck；确实要叠字就写 force_title。
        if s.get("shell_has_title") and not s.get("force_title"):
            title = None
        if title:
            l, t, w, h = self.g("closing_title")
            size = autofit_size([title], w, h, self.ts("cover_title", 40),
                                self.ts("cover_title", 40) * 0.42, line_spacing=1.15,
                                max_lines=s.get("max_lines", 2))
            set_text(textbox(slide, l, t, w, h).text_frame, title, self.t,
                     role="display" if s.get("display_font", True) else "heading",
                     size=size, color=s.get("color", "white"), bold=True,
                     align=s.get("align", "center"), anchor="middle", line_spacing=1.15)
        sub = s.get("subtitle")
        if sub:
            l, t, w, h = self.g("closing_sub")
            set_text(textbox(slide, l, t, w, h).text_frame, sub, self.t, role="body",
                     size=self.ts("cover_sub", 18), color=s.get("sub_color", "white"),
                     align=s.get("align", "center"), anchor="middle")


    # ---------------- 主循环 ----------------
    LAYOUTS = {
        "bullets": "s_bullets", "two-col": "s_two_col", "cards": "s_cards",
        "image-text": "s_image_text", "image-full": "s_image_full",
        "compare": "s_compare", "timeline": "s_timeline", "steps": "s_steps",
        "kpi": "s_kpi", "table": "s_table", "chart": "s_chart",
        "quote": "s_quote", "blank": None,
    }

    QA_NOTES_BUDGET = 600   # 结尾页问答备份的字数上限（不含用户自己写的备注）
    QA_ANSWER_CAP = 120     # 备份里每条回答最多保留的字数

    def _qa_backup(self):
        """把顶层 qa 压成结尾页备注里的精简备份。

        只有问答备份受 600 字预算约束；按整题取舍，不在句子中间截断。
        放不下的题数写进备注并记警告——完整版本来就该在对话里交付。"""
        head = "模拟导师问答（精简备份，完整版以对话交付为准）："
        lines, used, dropped = [head], len(head), 0
        for item in self.deck.get("qa") or []:
            q = str(item.get("question", "")).strip()
            if not q:
                continue
            a = str(item.get("answer", "")).strip()
            if len(a) > self.QA_ANSWER_CAP:
                a = a[:self.QA_ANSWER_CAP].rstrip() + "…"
            block = [f"Q：{q}"] + ([f"A：{a}"] if a else [])
            cost = sum(len(x) + 1 for x in block)
            if dropped or used + cost > self.QA_NOTES_BUDGET:
                dropped += 1
                continue
            lines += block
            used += cost
        if dropped:
            lines.append(f"（另有 {dropped} 题超出备注篇幅，见对话中的完整版）")
            self.warnings.append(f"问答备份超出 {self.QA_NOTES_BUDGET} 字，"
                                 f"{dropped} 题未写入结尾页备注")
        return "\n".join(lines)

    def _notes_for(self, slide_spec):
        # 用户自己写的讲稿备注原样保留，不截断
        parts = []
        if slide_spec.get("notes"):
            parts.append(str(slide_spec["notes"]).strip())
        if slide_spec.get("layout") == "closing" and self.deck.get("qa") \
                and not self._qa_written:
            parts.append(self._qa_backup())
            self._qa_written = True
        return "\n\n".join(x for x in parts if x)

    def _add_source(self, slide, source):
        if not source:
            return
        l, t, w, h = self.g("footer")
        set_text(textbox(slide, l, t, w, h).text_frame,
                 "来源：" + str(source), self.t, role="body",
                 size=self.ts("caption", 11), color="muted",
                 align="right", anchor="middle", space_after=0, line_spacing=1.0)

    def _is_compact(self):
        """显式 presentation_mode 优先；没写时 10 分钟以内的汇报默认紧凑。"""
        meta = self.deck.get("meta", {})
        mode = meta.get("presentation_mode")
        if mode:
            return mode == "compact"
        try:
            return 0 < float(meta.get("duration_minutes") or 0) <= 10
        except (TypeError, ValueError):
            return False

    def run(self):
        sec_i = 0
        defaults = self.t.data.get("defaults", {})
        compact = self._is_compact()
        self._qa_written = False
        for raw in self.deck["slides"]:
            lay = raw.get("layout", "bullets")
            # 主题可以给某类页定默认配色（比如「结尾页外壳是浅色照片，字必须深」），
            # deck.json 里显式写的永远优先。
            s = {**defaults.get(lay, {}), **raw}

            if compact and lay == "toc":
                continue
            if compact and lay == "section":
                sec_i += 1
                self.cur_chapter = sec_i
                continue

            if lay == "cover":
                sl = self.shell("cover")
                self.s_cover(sl, s)
            elif lay == "toc":
                sl = self.shell("toc")
                self.s_toc(sl, s)
            elif lay == "section":
                sec_i += 1
                self.cur_chapter = sec_i
                sl = self.shell("section")
                self.s_section(sl, s, sec_i)
            elif lay == "closing":
                sl = self.shell("closing")
                self.s_closing(sl, s)
            else:
                sl = self.shell("content")
                self.draw_nav(sl, s.get("chapter", self.cur_chapter))
                self.page_head(sl, s.get("title"), s.get("kicker"),
                               rule=s.get("rule", True))
                fn = self.LAYOUTS.get(lay)
                if fn is None and lay not in self.LAYOUTS:
                    self.warnings.append(f"未知版式 '{lay}'，已按 bullets 处理")
                    fn = "s_bullets"
                if fn:
                    getattr(self, fn)(sl, s)

            self._add_source(sl, s.get("source"))
            self.notes(sl, self._notes_for(s))

        if self.deck.get("qa") and not self._qa_written:
            self.warnings.append("deck 里有 qa 但没有 closing 页，问答备份没有写进任何备注")
        drop_slides(self.prs, list(range(self.n_shells)))
        return self.prs


# ==========================================================================
def build(deck_path, out_path=None, theme_name=None, asset_dir=None,
          allow_missing_assets=False):
    deck = json.loads(Path(deck_path).read_text(encoding="utf-8"))
    tname = theme_name or deck.get("theme") or "jnu-teal"
    tdir = ROOT / "themes" / tname
    if not tdir.exists():
        raise SystemExit(f"没有这套主题：{tname}（可选：{[p.name for p in (ROOT/'themes').iterdir() if p.is_dir()]}）")
    # 图片根目录：命令行 > deck.json > deck.json 所在目录。
    # deck.json 里写的相对路径按**deck.json 自己的位置**解析，不是当前工作目录——
    # 否则换个目录跑同一份 deck 就全变成缺图。
    deck_dir = Path(deck_path).resolve().parent
    if asset_dir:
        deck["asset_dir"] = str(Path(asset_dir).resolve())
    elif deck.get("asset_dir"):
        deck["asset_dir"] = str((deck_dir / deck["asset_dir"]).resolve())
    else:
        deck["asset_dir"] = str(deck_dir)

    b = Builder(deck, Theme(tdir))
    prs = b.run()

    if b.missing_assets and not allow_missing_assets:
        missing = "\n".join(f"  - {p}" for p in b.missing_assets)
        raise SystemExit(
            "发现缺失图片，已停止导出，避免把占位框当成成品：\n" + missing +
            "\n若只是在调试版式，可显式加 --allow-missing-assets。")

    if not out_path:
        meta = deck.get("meta", {})
        stem = sanitize_filename(meta.get("file_name") or meta.get("title") or "汇报")
        out_path = Path(deck_path).parent / f"{stem}.pptx"
    out_path = Path(out_path)
    out_path.parent.mkdir(parents=True, exist_ok=True)
    meta = deck.get("meta", {})
    # 文件属性一律按本次 deck 重写，不继承模板原作者/创建时间
    cp = prs.core_properties
    cp.title = str(meta.get("title") or out_path.stem)
    cp.subject = str(meta.get("subtitle") or "学术汇报")
    cp.author = str(meta.get("presenter") or "")
    cp.last_modified_by = str(meta.get("presenter") or "")
    cp.keywords = "暨南大学, 学术汇报, 组会, 答辩"
    cp.comments = ""
    cp.category = ""
    cp.revision = 1
    cp.created = cp.modified = datetime.now(timezone.utc).replace(tzinfo=None)
    prs.save(str(out_path))
    return out_path, b.warnings


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("deck")
    ap.add_argument("--out")
    ap.add_argument("--theme")
    ap.add_argument("--assets", help="图片相对路径的根目录，默认是 deck.json 所在目录")
    ap.add_argument("--allow-missing-assets", action="store_true",
                    help="仅调试用：缺图时仍导出带占位框的 PPT")
    a = ap.parse_args()
    out, warns = build(a.deck, a.out, a.theme, a.assets, a.allow_missing_assets)
    print(f"已生成：{out}  ({out.stat().st_size/1048576:.2f} MB)")
    for w in warns:
        print("  ! " + w)


if __name__ == "__main__":
    main()
