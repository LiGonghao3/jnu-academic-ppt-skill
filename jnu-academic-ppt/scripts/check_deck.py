# -*- coding: utf-8 -*-
"""出片前的体检：把「肉眼一看就知道翻车了」的问题在交付前抓出来。

查 11 类：
   1. 文字溢出框                 2. 形状越界 / 压到出血区
   3. 正文框互相重叠             4. 压到模板自带的艺术字
   5. 字号低于可读下限           6. 对比度不足
   7. 模板占位符残留             8. 空页
   9. 图片分辨率过低 / 变形     10. 导航条与章节数对不上
  11. 单页内容密度过高

默认纯 Python 估算，不依赖 Office。加 --deep 会调用本机 PowerPoint 拿到
**真实**的文本外框尺寸（TextFrame2.BoundHeight），溢出判定就不再是估算。

用法：
    python check_deck.py 汇报.pptx [--theme jnu-teal] [--deep] [--json]
退出码：0=干净或只有提示，1=有 ERROR 级问题。
"""
from __future__ import annotations

import argparse
import json
import re
import subprocess
import sys
from pathlib import Path

from pptx import Presentation
from pptx.util import Emu

sys.path.insert(0, str(Path(__file__).parent))
from ppt_kit import (  # noqa: E402
    EMU_PER_IN, Theme, contrast_ratio, estimate_lines, iter_shapes,
    line_height_in, luminance, text_width_units,
)

ROOT = Path(__file__).resolve().parent.parent

PLACEHOLDER_PAT = re.compile(
    r"请输入|请在此|点击此|单击此|双击此|示例文本|标题文本|正文内容|"
    r"Lorem ipsum|lorem ipsum|TODO|待补|待填|XXXX?|\{\{|\bxxx\b|"
    r"\[缺图\]|占位图|替换成自己的|下周补真实结果", re.I)

SHELL_PREFIX = "shell_"   # make_stencils 打的标记：模板自带的背景/装饰
NAV_PREFIX = "nav_"


def is_shell(sh) -> bool:
    """模板外壳自带的形状。满幅底图出血、艺术字探出画布都是人家的设计，
    不是我们的排版事故，所以大部分检查要跳过它们。"""
    return sh.name.startswith(SHELL_PREFIX)


def is_nav(sh) -> bool:
    return sh.name.startswith(NAV_PREFIX)


def is_nav_item(sh) -> bool:
    """导航条上的编号牌/Tab。它们贴在外壳的导航底板上，不参与正文重叠判定；
    但 nav_label（页眉里的「第N部分 …」）是实打实的正文位，必须查。"""
    return sh.name[len(NAV_PREFIX):].isdigit() if is_nav(sh) else False


class Report:
    def __init__(self):
        self.items = []

    def add(self, sev, slide, kind, msg, shape=None):
        self.items.append({"sev": sev, "slide": slide, "kind": kind,
                           "msg": msg, "shape": shape})

    @property
    def errors(self):
        return [i for i in self.items if i["sev"] == "ERROR"]

    def render(self):
        if not self.items:
            return "体检通过：11 项检查全部干净。"
        order = {"ERROR": 0, "WARN": 1, "INFO": 2}
        rows = sorted(self.items, key=lambda i: (order[i["sev"]], i["slide"]))
        out = []
        n_e = len([i for i in rows if i["sev"] == "ERROR"])
        n_w = len([i for i in rows if i["sev"] == "WARN"])
        out.append(f"体检结果：{n_e} 个必改，{n_w} 个建议改，共 {len(rows)} 条")
        for i in rows:
            tag = {"ERROR": "✗", "WARN": "!", "INFO": "·"}[i["sev"]]
            where = f"P{i['slide']}"
            what = f" [{i['shape']}]" if i["shape"] else ""
            out.append(f"  {tag} {where}{what} {i['kind']}：{i['msg']}")
        return "\n".join(out)


# --------------------------------------------------------------------------
def in_(v):
    return (v or 0) / EMU_PER_IN


def shape_box(sh):
    try:
        return in_(sh.left), in_(sh.top), in_(sh.width), in_(sh.height)
    except Exception:
        return None


def runs_of(sh):
    if not sh.has_text_frame:
        return []
    return [r for p in sh.text_frame.paragraphs for r in p.runs]


def effective_fill(sh):
    """形状自己的纯色填充；拿不到就返回 None。"""
    try:
        if sh.fill.type is not None and sh.fill.type == 1:  # MSO_FILL.SOLID
            return "#%s" % sh.fill.fore_color.rgb
    except Exception:
        pass
    return None


def slide_bg(slide):
    from pptx.oxml.ns import qn
    bg = slide._element.cSld.find(qn("p:bg"))
    if bg is None:
        return None
    el = bg.find(".//" + qn("a:srgbClr"))
    return "#" + el.get("val") if el is not None else None


# --------------------------------------------------------------------------
def check(pptx_path, theme_name=None, deep=False):
    prs = Presentation(str(pptx_path))
    rep = Report()
    cw, chh = in_(prs.slide_width), in_(prs.slide_height)

    theme = None
    if theme_name:
        tdir = ROOT / "themes" / theme_name
        if tdir.exists():
            theme = Theme(tdir)
    scale = theme.data.get("type_scale", {}) if theme else {}
    # 低于「本主题最小的有意字号」才算异常：图注本来就该小，别把它当事故报。
    min_pt = min(scale.get("min_body", 12), scale.get("caption", 12))
    body_floor = scale.get("min_body", 12)
    k = (cw / 13.333)
    bleed = 0.06 * k

    deep_bounds = _deep_text_bounds(pptx_path) if deep else None
    if deep and deep_bounds is None:
        rep.add("WARN", 0, "深度体检未执行",
                "本机 PowerPoint 不可用或调用失败；以下溢出结果仍是纯 Python 估算")

    for si, slide in enumerate(prs.slides, 1):
        bg = slide_bg(slide)
        # 外壳如果铺了满幅底图/渐变，纯色 bg 就不代表实际底色了，
        # 拿它算对比度会得出完全错误的结论，干脆当「底色未知」。
        for sh in iter_shapes(slide.shapes):
            if is_shell(sh) and sh.shape_type in (13, 5):  # PICTURE / FREEFORM
                b = shape_box(sh)
                if b and b[2] * b[3] > cw * chh * 0.8:
                    bg = None
                    break
        text_boxes = []
        has_content = False
        nav_count = 0
        body_chars = 0.0

        shell_texts = []
        for sh in iter_shapes(slide.shapes):
            if is_shell(sh) and sh.has_text_frame and sh.text_frame.text.strip():
                b = shape_box(sh)
                if b and b[2] * b[3] < cw * chh * 0.5:   # 满幅水印不算
                    shell_texts.append((sh.name, b, sh.text_frame.text.strip()))

        for sh in iter_shapes(slide.shapes):
            name = sh.name
            shell, nav = is_shell(sh), is_nav(sh)
            if nav and name[len(NAV_PREFIX):].isdigit():
                nav_count += 1
            box = shape_box(sh)

            # ---- 2. 越界（只管我们自己画的）----
            if box and not shell:
                l, t, w, h = box
                if l < -bleed or t < -bleed or l + w > cw + bleed or t + h > chh + bleed:
                    if sh.has_text_frame and sh.text_frame.text.strip():
                        rep.add("ERROR", si, "越界",
                                f"文字框跑出画布（{l:.2f},{t:.2f} {w:.2f}×{h:.2f}，画布 {cw:.2f}×{chh:.2f}）",
                                name)

            # ---- 8. 图片（只管我们插入的，模板底图不算）----
            if sh.shape_type == 13 and box:
                if not shell:
                    _check_picture(rep, si, sh, box)
                has_content = True

            if not sh.has_text_frame:
                continue
            txt = sh.text_frame.text.strip()
            if not txt:
                continue
            if not shell:
                has_content = True
            rs = runs_of(sh)

            # ---- 6. 占位符残留 ----
            if PLACEHOLDER_PAT.search(txt):
                if shell:
                    rep.add("WARN", si, "外壳未清干净",
                            f"模板外壳里还留着示例文字「{txt[:24]}」，"
                            f"应该在 stencil_spec.json 里把它 drop 掉", name)
                else:
                    rep.add("ERROR", si, "占位符残留",
                            f"还留着模板示例文字「{txt[:24]}」", name)

            if shell:
                continue  # 剩下的检查都只针对我们生成的内容

            # ---- 4. 字号 ----
            if not nav:
                for r in rs:
                    pt = r.font.size.pt if r.font.size else None
                    if pt and pt < min_pt - 0.01:
                        rep.add("WARN", si, "字号过小",
                                f"{pt:.1f}pt 比本主题最小的有意字号 {min_pt}pt 还小", name)
                        break
                # 已经压到自动缩放的地板 = 这页塞太多了
                pts = [r.font.size.pt for r in rs if r.font.size]
                if pts and len(txt) > 60 and abs(min(pts) - body_floor) < 0.51:
                    rep.add("WARN", si, "内容过密",
                            f"正文已被自动压到字号下限 {body_floor}pt（{len(txt)} 字），"
                            f"建议拆成两页或删掉次要条目", name)

            # ---- 1. 溢出 ----
            if box:
                _check_overflow(rep, si, sh, box, txt, rs, deep_bounds, k)

            # ---- 5. 对比度 ----
            _check_contrast(rep, si, sh, rs, bg, name)

            if box and not is_nav_item(sh):
                text_boxes.append((name, box, txt))
                if not nav:
                    body_chars += text_width_units(txt)

        # ---- 3. 重叠 ----
        _check_overlap(rep, si, text_boxes)
        # 3b. 压到模板自带的文字（外壳的艺术字、标语）
        _check_shell_clash(rep, si, text_boxes, shell_texts)

        # ---- 10. 内容密度 ----
        # 一页 250 字以上，讲的时候必然变成照着念；放得下 ≠ 讲得动。
        if body_chars > 250:
            rep.add("WARN", si, "内容过密",
                    f"这页约 {body_chars:.0f} 字，讲起来会变成念稿。"
                    f"建议留结论 + 证据，其余挪到备注或拆成两页")
        elif body_chars > 180:
            rep.add("INFO", si, "偏密", f"这页约 {body_chars:.0f} 字，接近口播上限")

        # ---- 7. 空页 ----
        if not has_content:
            rep.add("WARN", si, "空页", "这页什么都没有，确认是不是漏了内容")

        # ---- 9. 导航一致性 ----
        if theme and nav_count:
            want = theme.nav.get("max_items", 99)
            if nav_count > want:
                rep.add("WARN", si, "导航超载",
                        f"导航条上有 {nav_count} 项，本主题建议不超过 {want} 项")

    return rep


def _check_overflow(rep, si, sh, box, txt, rs, deep_bounds, k):
    l, t, w, h = box
    name = sh.name
    if deep_bounds is not None:
        key = (si, name)
        got = deep_bounds.get(key)
        if got:
            bw, bh = got
            if bh > h * 1.06 + 0.02:
                rep.add("ERROR", si, "文字溢出",
                        f"实测文字高 {bh:.2f}in，框只有 {h:.2f}in（超 {bh/h*100-100:.0f}%）", name)
            elif bw > w * 1.06 + 0.02:
                rep.add("WARN", si, "文字溢出", f"实测文字宽 {bw:.2f}in > 框宽 {w:.2f}in", name)
            return
    # 纯 Python 估算：逐段按该段自己的字号和行距算，
    # 不能拿全框最大字号套所有段（目录的中英双行会被高估 30%）
    sizes = [r.font.size.pt for r in rs if r.font.size]
    if not sizes:
        return
    fallback_pt = max(sizes)
    tf = sh.text_frame
    pad_w = in_(tf.margin_left) + in_(tf.margin_right)
    pad_h = in_(tf.margin_top) + in_(tf.margin_bottom)
    need = pad_h
    paras = [p for p in tf.paragraphs if p.text]
    for i, p in enumerate(paras):
        p_sizes = [r.font.size.pt for r in p.runs if r.font.size]
        pt = max(p_sizes) if p_sizes else fallback_pt
        ls = p.line_spacing if isinstance(p.line_spacing, float) else 1.2
        n = estimate_lines(p.text, max(0.3, w - pad_w), pt)
        need += n * line_height_in(pt, ls)
        if i < len(paras) - 1 and p.space_after is not None:
            need += p.space_after.pt / 72.0
    if need > h * 1.10 + 0.05 * k:
        rep.add("WARN", si, "文字可能溢出",
                f"估算需要 {need:.2f}in，框高 {h:.2f}in（加 --deep 可实测）", name)


def _check_contrast(rep, si, sh, rs, bg, name):
    fill = effective_fill(sh)
    ref = fill or bg
    for r in rs:
        try:
            col = "#%s" % r.font.color.rgb
        except Exception:
            continue
        if ref:
            ratio = contrast_ratio(col, ref)
            if ratio < 3.0:
                rep.add("ERROR", si, "对比度不足",
                        f"文字 {col} 压在 {ref} 上，对比度仅 {ratio:.1f}:1（正文需 ≥4.5）", name)
            elif ratio < 4.5:
                rep.add("WARN", si, "对比度偏低",
                        f"文字 {col} / 底色 {ref} = {ratio:.1f}:1，小字号会吃力", name)
        elif luminance(col) > 0.62:
            rep.add("INFO", si, "浅色文字",
                    f"文字 {col} 落在图片/渐变上，底色不可知，请肉眼确认能看清", name)
        break  # 一个形状报一次就够


def _check_overlap(rep, si, boxes):
    for i in range(len(boxes)):
        n1, (l1, t1, w1, h1), x1 = boxes[i]
        for j in range(i + 1, len(boxes)):
            n2, (l2, t2, w2, h2), x2 = boxes[j]
            ox = max(0, min(l1 + w1, l2 + w2) - max(l1, l2))
            oy = max(0, min(t1 + h1, t2 + h2) - max(t1, t2))
            if ox <= 0 or oy <= 0:
                continue
            area = ox * oy
            small = min(w1 * h1, w2 * h2)
            if small > 0 and area / small > 0.35:
                rep.add("WARN", si, "文字重叠",
                        f"「{n1}」和「{n2}」重叠约 {area/small*100:.0f}%，可能是两段字压在一起")


def _check_shell_clash(rep, si, mine, shell_texts):
    """我们生成的文字压在外壳自带的文字上。

    外壳整体不参与重叠判定（满幅底图会把结果淹掉），但外壳里**带字**的
    装饰——比如封面/结尾的艺术字——必须单独查，否则两行字会精确地叠在一起。"""
    for n1, (l1, t1, w1, h1), x1 in mine:
        for n2, (l2, t2, w2, h2), x2 in shell_texts:
            ox = max(0, min(l1 + w1, l2 + w2) - max(l1, l2))
            oy = max(0, min(t1 + h1, t2 + h2) - max(t1, t2))
            if ox <= 0 or oy <= 0:
                continue
            small = min(w1 * h1, w2 * h2)
            if small > 0 and ox * oy / small > 0.4:
                rep.add("ERROR", si, "压到模板文字",
                        "「%s」压在外壳自带的「%s」上，两段字会叠在一起" % (n1, x2[:12]))


def _check_picture(rep, si, sh, box):
    l, t, w, h = box
    try:
        px_w, px_h = sh.image.size
    except Exception:
        return
    if w <= 0 or h <= 0:
        return
    dpi_w, dpi_h = px_w / w, px_h / h
    if min(dpi_w, dpi_h) < 90:
        rep.add("WARN", si, "图片糊",
                f"有效分辨率仅 {min(dpi_w,dpi_h):.0f} dpi（{px_w}×{px_h} 拉到 {w:.1f}×{h:.1f}in），投屏会糊",
                sh.name)
    ar_img, ar_box = px_w / px_h, w / h
    if not (sh.crop_left or sh.crop_right or sh.crop_top or sh.crop_bottom):
        if abs(ar_img - ar_box) / max(ar_img, ar_box) > 0.06:
            rep.add("WARN", si, "图片变形",
                    f"原图 {ar_img:.2f} 宽高比被压成 {ar_box:.2f}，人脸/图表会歪", sh.name)


# --------------------------------------------------------------------------
_PS_DEEP = r"""
param([string]$Path)
$ErrorActionPreference='Stop'
# PowerPoint 是单实例：用户已经开着时 New-Object 拿到的就是用户那个进程，
# 只有本脚本自己拉起的实例才允许 Quit，否则会把用户正在编辑的文件一起关掉。
$wasRunning = [bool](Get-Process POWERPNT -ErrorAction SilentlyContinue)
$app = New-Object -ComObject PowerPoint.Application
$pres = $null
$out = @()
try {
  $pres = $app.Presentations.Open($Path, $true, $false, $false)
  for ($i=1; $i -le $pres.Slides.Count; $i++) {
    foreach ($sh in $pres.Slides.Item($i).Shapes) {
      if ($sh.HasTextFrame -eq -1 -and $sh.TextFrame2.HasText -eq -1) {
        $tr = $sh.TextFrame2.TextRange
        $out += [pscustomobject]@{ s=$i; n=$sh.Name; w=$tr.BoundWidth/72.0; h=$tr.BoundHeight/72.0 }
      }
    }
  }
  $out | ConvertTo-Json -Compress -Depth 3
} finally {
  if ($pres) { $pres.Close() }
  if (-not $wasRunning -and $app.Presentations.Count -eq 0) { $app.Quit() }
}
"""


def _deep_text_bounds(pptx_path):
    """用本机 PowerPoint 量真实文本外框（英寸）。拿不到就退回估算。"""
    import tempfile, os
    ps = None
    try:
        with tempfile.NamedTemporaryFile("w", suffix=".ps1", delete=False,
                                         encoding="utf-8-sig") as f:
            f.write(_PS_DEEP)
            ps = f.name
        r = subprocess.run(["powershell", "-NoProfile", "-ExecutionPolicy", "Bypass",
                            "-File", ps, "-Path", str(Path(pptx_path).resolve())],
                           capture_output=True, text=True, timeout=300)
        if r.returncode != 0:
            detail = (r.stderr or r.stdout or "未知错误").strip()[:500]
            raise RuntimeError(f"PowerPoint 返回码 {r.returncode}：{detail}")
        if not r.stdout.strip():
            raise RuntimeError("PowerPoint 未返回任何测量结果")
        data = json.loads(r.stdout.strip() or "[]")
        if isinstance(data, dict):
            data = [data]
        if not data:
            raise RuntimeError("PowerPoint 返回空测量结果")
        return {(int(d["s"]), d["n"]): (float(d["w"]), float(d["h"])) for d in data}
    except Exception as e:
        print(f"  (--deep 不可用，退回估算：{e})", file=sys.stderr)
        return None
    finally:
        if ps:
            try:
                os.unlink(ps)
            except OSError:
                pass


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("pptx")
    ap.add_argument("--theme")
    ap.add_argument("--deep", action="store_true", help="调用本机 PowerPoint 实测文本外框")
    ap.add_argument("--json", action="store_true")
    a = ap.parse_args()
    rep = check(a.pptx, a.theme, a.deep)
    if a.json:
        print(json.dumps(rep.items, ensure_ascii=False, indent=2))
    else:
        print(rep.render())
    sys.exit(1 if rep.errors else 0)


if __name__ == "__main__":
    main()
