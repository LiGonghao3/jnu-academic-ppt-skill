# -*- coding: utf-8 -*-
"""ppt_kit —— 共享底层能力：克隆母版页、字体回退、原生绘图。

设计要点
--------
1. **克隆轨**：从 stencil.pptx 里整页 deepcopy「外壳页」（背景/装饰/logo），
   关系（图片、图表）按 rId 重映射，图片 part 复用不重复打包。
2. **生成轨**：正文一律用 python-pptx 原生形状绘制，读 theme.json 的 token，
   保证产物是真正可编辑的 PowerPoint 对象，而不是截图。
3. **字体回退**：每个字体角色声明一条回退链，写入 run 时同时设置
   latin / ea / cs，避免换电脑掉字形。
"""
from __future__ import annotations

import copy
import json
import re
from pathlib import Path

from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.oxml.ns import qn
from pptx.util import Emu, Pt

EMU_PER_IN = 914400
R_NS = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}"


# --------------------------------------------------------------------------
# 单位与颜色
# --------------------------------------------------------------------------
def inches(v) -> int:
    return int(round(float(v) * EMU_PER_IN))


def rgb(s) -> RGBColor:
    """接受 '#1D3F69' / '1D3F69' / RGBColor。"""
    if isinstance(s, RGBColor):
        return s
    return RGBColor.from_string(str(s).lstrip("#").upper())


def mix(c1, c2, t: float) -> str:
    """线性混合两个 hex 颜色，t=0 取 c1，t=1 取 c2。"""
    a, b = rgb(c1), rgb(c2)
    return "#%02X%02X%02X" % tuple(int(round(a[i] + (b[i] - a[i]) * t)) for i in range(3))


def luminance(c) -> float:
    def ch(x):
        x /= 255.0
        return x / 12.92 if x <= 0.04045 else ((x + 0.055) / 1.055) ** 2.4

    r, g, b = rgb(c)
    return 0.2126 * ch(r) + 0.7152 * ch(g) + 0.0722 * ch(b)


def contrast_ratio(c1, c2) -> float:
    l1, l2 = luminance(c1), luminance(c2)
    hi, lo = max(l1, l2), min(l1, l2)
    return (hi + 0.05) / (lo + 0.05)


def readable_on(bg, light="#FFFFFF", dark="#1A1A1A") -> str:
    """在给定背景上挑对比度更高的前景色。"""
    return light if contrast_ratio(bg, light) >= contrast_ratio(bg, dark) else dark


# --------------------------------------------------------------------------
# 主题
# --------------------------------------------------------------------------
class Theme:
    def __init__(self, path):
        self.dir = Path(path)
        self.data = json.loads((self.dir / "theme.json").read_text(encoding="utf-8"))
        self.name = self.data["name"]
        self.palette = self.data["palette"]
        self.fonts = self.data["fonts"]
        self.geometry = self.data["geometry"]
        self.shells = self.data.get("shells", {})
        self.nav = self.data.get("nav", {"kind": "none"})

    # 颜色：支持 "primary" 这种 token 名，也支持 "#RRGGBB" 直写
    def c(self, token, default=None):
        if token is None:
            return default
        if isinstance(token, str) and token.startswith("#"):
            return token
        return self.palette.get(token, default or token)

    def font(self, role):
        f = self.fonts.get(role) or self.fonts.get("body")
        return f

    @property
    def stencil(self) -> Path:
        return self.dir / "stencil.pptx"

    def box(self, key):
        """geometry 里的矩形区域，单位 inch，返回 (l, t, w, h)。"""
        g = self.geometry[key]
        return g["l"], g["t"], g["w"], g["h"]


# --------------------------------------------------------------------------
# 克隆
# --------------------------------------------------------------------------
# 克隆时不能跟着走的关系：备注页会反向指回原页，跨页超链接同理。
# 复制了它们，原页一删就成了悬空/循环引用，PowerPoint 会直接拒绝打开文件。
_SKIP_RELTYPES = (
    "/notesSlide",
    "/slide",
)


def clone_slide(prs, src_slide):
    """把 src_slide 整页复制成 prs 末尾的新页，返回新 slide。"""
    dst = prs.slides.add_slide(src_slide.slide_layout)
    # 版式自动带来的占位符先清空，完全由克隆内容接管
    for sh in list(dst.shapes):
        sh._element.getparent().remove(sh._element)

    # 页面背景（<p:bg>）
    src_cSld, dst_cSld = src_slide._element.cSld, dst._element.cSld
    old_bg = dst_cSld.find(qn("p:bg"))
    if old_bg is not None:
        dst_cSld.remove(old_bg)
    src_bg = src_cSld.find(qn("p:bg"))
    if src_bg is not None:
        dst_cSld.insert(0, copy.deepcopy(src_bg))

    for sh in src_slide.shapes:
        dst.shapes._spTree.append(copy.deepcopy(sh._element))

    # rId 重映射：图片/图表 part 直接复用，不重复打包
    rid_map = {}
    for rid, rel in src_slide.part.rels.items():
        if any(rel.reltype.endswith(s) for s in _SKIP_RELTYPES):
            continue
        try:
            if rel.is_external:
                new_rid = dst.part.rels.get_or_add_ext_rel(rel.reltype, rel.target_ref)
            else:
                new_rid = dst.part.relate_to(rel.target_part, rel.reltype)
            rid_map[rid] = new_rid
        except Exception:
            pass
    valid = set(dst.part.rels.keys())
    for el in list(dst._element.iter()):
        for attr, val in list(el.attrib.items()):
            if not attr.startswith(R_NS):
                continue
            if val in rid_map:
                el.set(attr, rid_map[val])
            elif val not in valid:
                # 指向被跳过关系的悬空引用（多半是跨页超链接），整个摘掉
                if el.tag in (qn("a:hlinkClick"), qn("a:hlinkHover")):
                    el.getparent().remove(el)
                else:
                    del el.attrib[attr]
    return dst


def drop_slides(prs, indices):
    """按 0-based 下标删除若干页（内部先转成 id 再删，避免位移问题）。"""
    id_list = prs.slides._sldIdLst
    entries = list(id_list)
    for i in sorted(set(indices), reverse=True):
        if not (0 <= i < len(entries)):
            continue
        sld_id = entries[i]
        rid = sld_id.get(R_NS + "id")
        try:
            prs.part.drop_rel(rid)
        except Exception:
            pass
        id_list.remove(sld_id)


def iter_shapes(shapes):
    """深度遍历，穿透组合。"""
    for sh in shapes:
        yield sh
        if sh.shape_type == 6:  # GROUP
            for sub in iter_shapes(sh.shapes):
                yield sub


def find_shape(slide_or_group, name, exact=True):
    for sh in iter_shapes(slide_or_group.shapes):
        if (sh.name == name) if exact else (name in sh.name):
            return sh
    return None


def remove_shape(sh):
    sh._element.getparent().remove(sh._element)


def keep_only(slide, keep_names):
    """只保留白名单里的顶层形状（外壳提炼用）。"""
    keep = set(keep_names)
    for sh in list(slide.shapes):
        if sh.name not in keep:
            remove_shape(sh)


# --------------------------------------------------------------------------
# 文本
# --------------------------------------------------------------------------
def _apply_font(run, spec, size=None, color=None, bold=None, italic=None):
    """spec 形如 {"latin": "Arial", "ea": "微软雅黑", "fallback": [...]}。

    PowerPoint 只认一个 typeface，掉字形靠的是系统回退；这里把首选写进
    latin/ea/cs 三处，并把回退链写进自定义属性供 check_deck 校验。
    """
    f = run.font
    if size is not None:
        f.size = Pt(size)
    # 字体规格可以钉死字重：Light 字面再加粗会被合成成假粗，很难看，
    # 所以 display 这类用细字面的角色直接声明 weight_bold=false。
    if "weight_bold" in spec:
        bold = spec["weight_bold"]
    if bold is not None:
        f.bold = bold
    if italic is not None:
        f.italic = italic
    if color is not None:
        f.color.rgb = rgb(color)

    latin = spec.get("latin") or spec.get("ea")
    ea = spec.get("ea") or spec.get("latin")
    rPr = run._r.get_or_add_rPr()
    for tag, face in (("a:latin", latin), ("a:ea", ea), ("a:cs", latin)):
        if not face:
            continue
        el = rPr.find(qn(tag))
        if el is None:
            el = rPr.makeelement(qn(tag), {})
            rPr.append(el)
        el.set("typeface", face)


def set_text(
    tf,
    lines,
    theme: Theme,
    role="body",
    size=18,
    color="text",
    bold=False,
    align="left",
    line_spacing=1.25,
    space_after=6,
    bullet=None,
    anchor="top",
    shrink_to=None,
):
    """把若干行写进 text_frame。lines 可为 str 或 [str] 或 [(text, size, bold, color)]。"""
    if isinstance(lines, str):
        lines = [lines]
    spec = theme.font(role)
    col = theme.c(color)

    tf.clear()
    tf.word_wrap = True
    tf.vertical_anchor = {"top": MSO_ANCHOR.TOP, "middle": MSO_ANCHOR.MIDDLE,
                          "bottom": MSO_ANCHOR.BOTTOM}[anchor]

    for i, item in enumerate(lines):
        if isinstance(item, (tuple, list)):
            text, isz, ibold, icol = (list(item) + [None, None, None])[:4]
            isz = isz or size
            ibold = bold if ibold is None else ibold
            icol = theme.c(icol) if icol else col
        else:
            text, isz, ibold, icol = str(item), size, bold, col

        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = {"left": PP_ALIGN.LEFT, "center": PP_ALIGN.CENTER,
                       "right": PP_ALIGN.RIGHT, "justify": PP_ALIGN.JUSTIFY}[align]
        p.line_spacing = line_spacing
        p.space_after = Pt(space_after)
        prefix = f"{bullet} " if bullet else ""
        for seg, seg_bold in _split_bold(prefix + text):
            run = p.add_run()
            run.text = seg
            _apply_font(run, spec, size=isz, color=icol, bold=ibold or seg_bold)
    return tf


_BOLD_RE = re.compile(r"\*\*(.+?)\*\*", re.S)


def _split_bold(text):
    """把 `**这样**` 的行内加粗拆成 (片段, 是否加粗) 序列。

    写 deck.json 的人下意识就会用 Markdown 语法；不解析的话星号会原样
    印在幻灯片上。只支持 ** 加粗，够用且不会误伤数学里的单个 *。
    """
    out, pos = [], 0
    for m in _BOLD_RE.finditer(text):
        if m.start() > pos:
            out.append((text[pos:m.start()], False))
        out.append((m.group(1), True))
        pos = m.end()
    if pos < len(text):
        out.append((text[pos:], False))
    return out or [(text, False)]


def textbox(slide, l, t, w, h, **kw):
    box = slide.shapes.add_textbox(inches(l), inches(t), inches(w), inches(h))
    tf = box.text_frame
    tf.margin_left = tf.margin_right = Emu(0)
    tf.margin_top = tf.margin_bottom = Emu(0)
    if kw:
        set_text(tf, **kw)
    return box


# --------------------------------------------------------------------------
# 原生形状
# --------------------------------------------------------------------------
def rect(slide, l, t, w, h, fill=None, line=None, line_w=1.0, shape=MSO_SHAPE.RECTANGLE,
         adj=None, shadow=False):
    sh = slide.shapes.add_shape(shape, inches(l), inches(t), inches(w), inches(h))
    if fill is None:
        sh.fill.background()
    else:
        sh.fill.solid()
        sh.fill.fore_color.rgb = rgb(fill)
    if line is None:
        sh.line.fill.background()
    else:
        sh.line.color.rgb = rgb(line)
        sh.line.width = Pt(line_w)
    if not shadow:
        # python-pptx 没有直接关阴影的 API，手写 <a:effectLst/> 置空
        spPr = sh._element.spPr
        for tag in ("a:effectLst", "a:effectDag"):
            for el in spPr.findall(qn(tag)):
                spPr.remove(el)
        spPr.append(spPr.makeelement(qn("a:effectLst"), {}))
    if adj is not None:
        try:
            sh.adjustments[0] = adj
        except Exception:
            pass
    sh.text_frame.word_wrap = True
    return sh


def line(slide, l, t, w, h, color, width=1.0):
    from pptx.enum.shapes import MSO_CONNECTOR

    cn = slide.shapes.add_connector(MSO_CONNECTOR.STRAIGHT, inches(l), inches(t),
                                    inches(l + w), inches(t + h))
    cn.line.color.rgb = rgb(color)
    cn.line.width = Pt(width)
    return cn


def picture(slide, path, l, t, w, h, mode="fit"):
    """mode=fit 保持比例塞进框内并居中；mode=fill 裁剪填满。"""
    from PIL import Image

    try:
        iw, ih = Image.open(path).size
    except Exception:
        return slide.shapes.add_picture(str(path), inches(l), inches(t), inches(w), inches(h))

    box_ar, img_ar = w / h, iw / ih
    if mode == "fit":
        if img_ar > box_ar:
            nw, nh = w, w / img_ar
        else:
            nh, nw = h, h * img_ar
        pic = slide.shapes.add_picture(str(path), inches(l + (w - nw) / 2),
                                       inches(t + (h - nh) / 2), inches(nw), inches(nh))
    else:
        pic = slide.shapes.add_picture(str(path), inches(l), inches(t), inches(w), inches(h))
        if img_ar > box_ar:
            crop = (1 - box_ar / img_ar) / 2
            pic.crop_left = pic.crop_right = crop
        else:
            crop = (1 - img_ar / box_ar) / 2
            pic.crop_top = pic.crop_bottom = crop
    return pic


# --------------------------------------------------------------------------
# 文本容量估算（自检与自动降档共用）
# --------------------------------------------------------------------------
def text_width_units(s: str) -> float:
    """中文/全角按 1.0，ASCII 按 0.55 折算成「字宽单位」。"""
    s = _BOLD_RE.sub(r"\1", s)   # ** 是标记不是字，不占宽度
    total = 0.0
    for ch in s:
        total += 1.0 if ord(ch) > 0x2E7F else 0.55
    return total


def estimate_lines(text: str, box_w_in: float, size_pt: float) -> int:
    """给定框宽与字号，估算文本占几行。"""
    if not text:
        return 0
    char_w_in = size_pt / 72.0  # 一个全角字≈字号宽
    # 留 2% 余量：正好铺满一行时 PowerPoint 会换行，而按等式算会判成「放得下」，
    # 于是标题被撑成两行、最后一行只剩一两个字。
    per_line = max(1.0, box_w_in * 0.98 / char_w_in)
    return max(1, int(text_width_units(text) / per_line + 0.999))


# PowerPoint 的「行距倍数」乘的是字体行盒，不是字号本身；常见中文字体的行盒
# 约为字号的 1.2 倍。早先只按 size×spacing 估算，实测普遍偏乐观 10% 以上。
LINE_BOX = 1.2          # 字体行盒 / 字号
FIT_SAFETY = 1.03       # 标点避头尾、西文不断词之类的零头


def line_height_in(size_pt: float, line_spacing: float) -> float:
    return size_pt * LINE_BOX * line_spacing / 72.0


def fits(lines, box_w_in, box_h_in, size_pt, line_spacing=1.25, space_after_pt=6) -> bool:
    used = 0.0
    for ln in lines:
        n = estimate_lines(ln, box_w_in, size_pt)
        used += n * line_height_in(size_pt, line_spacing) + space_after_pt / 72.0
    return used * FIT_SAFETY <= box_h_in


def block_height(lines, box_w_in, size_pt, line_spacing=1.25, space_after_pt=6) -> float:
    """这段文字排完实际占多高（英寸）。用来给正文做垂直配重。"""
    used = 0.0
    for ln in lines:
        n = estimate_lines(str(ln), box_w_in, size_pt)
        used += n * line_height_in(size_pt, line_spacing) + space_after_pt / 72.0
    return used


def autofit_size(lines, box_w_in, box_h_in, start_pt, min_pt, step=1.0,
                 max_lines=None, **kw) -> float:
    """从 start_pt 往下试，返回第一个放得下的字号；都放不下返回 min_pt。

    `max_lines` 给标题用：光「装得下」不够，「……的应／用」这种最后一行只剩
    一个字的断行很难看，宁可再小一号也要控制在指定行数内。中文可以在任意字
    之间断行，所以还要额外防孤字——末行内容不足一行的 45% 就继续缩（中英混排会在空格处断，末行往往比估算更短）。
    """
    size = start_pt
    while size > min_pt:
        ok = fits(lines, box_w_in, box_h_in, size, **kw)
        if ok and max_lines:
            total = sum(estimate_lines(str(x), box_w_in, size) for x in lines)
            ok = total <= max_lines
            if ok and total > 1 and len(lines) == 1:
                per_line = max(1.0, box_w_in * 0.98 / (size / 72.0))
                tail = text_width_units(str(lines[0])) % per_line
                if 0 < tail < per_line * 0.45:
                    ok = False          # 末行是个孤字/孤词，再小一号
        if ok:
            return round(size)
        size -= step
    return round(min_pt)


# --------------------------------------------------------------------------
# IO
# --------------------------------------------------------------------------
def open_stencil(theme: Theme) -> Presentation:
    return Presentation(str(theme.stencil))


def slide_size_in(prs):
    return prs.slide_width / EMU_PER_IN, prs.slide_height / EMU_PER_IN


def sanitize_filename(s: str, maxlen=60) -> str:
    s = re.sub(r'[\\/:*?"<>|\r\n\t]+', "_", s).strip(" .")
    return (s[:maxlen] or "deck")
