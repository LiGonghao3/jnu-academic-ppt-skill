# -*- coding: utf-8 -*-
"""给 .pptx 瘦身：删掉不可达的孤儿 part，并把超大位图重采样。

删掉原始页之后，媒体文件仍然躺在包里（python-pptx 不做垃圾回收），
一套主题动辄 70MB，每周一份组会 PPT 会很难受。这里按 rels 图做可达性
分析，只保留真正引用到的 part，再把大图压到够用的分辨率。

用法：
    python slim_pptx.py <file.pptx> [--max-px 2400] [--quality 82] [--inplace]
"""
from __future__ import annotations

import argparse
import io
import posixpath
import re
import shutil
import zipfile
from pathlib import Path

from lxml import etree

RELS_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
RASTER = {".png", ".jpg", ".jpeg", ".bmp", ".tiff", ".tif"}


def rels_path_for(part: str) -> str:
    d, n = posixpath.split(part)
    return posixpath.join(d, "_rels", n + ".rels")


def part_for_rels(rels: str) -> str:
    d, n = posixpath.split(rels)
    return posixpath.join(posixpath.dirname(d), n[:-5])


def reachable_parts(z: zipfile.ZipFile) -> set[str]:
    names = set(z.namelist())
    seen: set[str] = set()
    queue = ["/_rels/.rels".lstrip("/")]

    while queue:
        rels = queue.pop()
        if rels in seen or rels not in names:
            continue
        seen.add(rels)
        base = posixpath.dirname(posixpath.dirname(rels)) or ""
        try:
            root = etree.fromstring(z.read(rels))
        except Exception:
            continue
        for rel in root.findall("{%s}Relationship" % RELS_NS):
            if rel.get("TargetMode") == "External":
                continue
            tgt = rel.get("Target")
            target = tgt[1:] if tgt.startswith("/") else posixpath.normpath(posixpath.join(base, tgt))
            if target in names:
                seen.add(target)
                queue.append(rels_path_for(target))
    # [Content_Types].xml 必留
    seen.add("[Content_Types].xml")
    return seen


def shrink_image(data: bytes, ext: str, max_px: int, quality: int):
    from PIL import Image

    try:
        im = Image.open(io.BytesIO(data))
        im.load()
    except Exception:
        return None
    w, h = im.size
    scale = min(1.0, max_px / max(w, h))
    if scale >= 1.0 and len(data) < 400_000:
        return None

    if scale < 1.0:
        im = im.resize((max(1, int(w * scale)), max(1, int(h * scale))), Image.LANCZOS)

    buf = io.BytesIO()
    has_alpha = im.mode in ("RGBA", "LA") or (im.mode == "P" and "transparency" in im.info)
    if ext == ".png" and has_alpha:
        im.convert("RGBA").save(buf, "PNG", optimize=True)
        out_ext = ".png"
    else:
        im.convert("RGB").save(buf, "JPEG", quality=quality, optimize=True, progressive=True)
        out_ext = ".jpeg"
    out = buf.getvalue()
    if len(out) >= len(data):
        return None
    return out, out_ext


FONT_RELTYPE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/font"


def strip_embedded_fonts(z: zipfile.ZipFile):
    """返回 (要跳过的 part 集合, {被改写的 part: 新字节})。

    这些模板把整套思源黑体/可画寒风体塞进了包里（单套 70MB）。可移植性确实好，
    但每周一份组会 PPT 背 70MB 不现实，默认剥掉，靠 theme.json 的回退链兜底。
    """
    skip, patched = set(), {}
    rels_name = "ppt/_rels/presentation.xml.rels"
    if rels_name not in z.namelist():
        return skip, patched

    root = etree.fromstring(z.read(rels_name))
    changed = False
    for rel in list(root.findall("{%s}Relationship" % RELS_NS)):
        if rel.get("Type") == FONT_RELTYPE:
            tgt = rel.get("Target")
            skip.add(tgt[1:] if tgt.startswith("/") else posixpath.normpath(posixpath.join("ppt", tgt)))
            root.remove(rel)
            changed = True
    if not changed:
        return skip, patched
    patched[rels_name] = etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)

    pres = "ppt/presentation.xml"
    proot = etree.fromstring(z.read(pres))
    for el in proot.findall("{http://schemas.openxmlformats.org/presentationml/2006/main}embeddedFontLst"):
        proot.remove(el)
    patched[pres] = etree.tostring(proot, xml_declaration=True, encoding="UTF-8", standalone=True)
    return skip, patched


def slim(path: Path, max_px=2400, quality=82, inplace=False, verbose=True,
         strip_fonts=True) -> Path:
    src = Path(path)
    tmp = src.with_suffix(".slim.pptx")

    with zipfile.ZipFile(src) as z:
        keep = reachable_parts(z)
        font_skip, patched = strip_embedded_fonts(z) if strip_fonts else (set(), {})
        keep -= font_skip
        dropped = [n for n in z.namelist() if n not in keep and not n.endswith("/")]
        saved_fonts = sum(i.file_size for i in z.infolist() if i.filename in font_skip)
        saved_media = 0

        with zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED, compresslevel=9) as out:
            for item in z.infolist():
                if item.filename.endswith("/") or item.filename not in keep:
                    continue
                if item.filename in patched:
                    out.writestr(item, patched[item.filename])
                    continue
                data = z.read(item.filename)
                ext = posixpath.splitext(item.filename)[1].lower()
                if item.filename.startswith("ppt/media/") and ext in RASTER:
                    res = shrink_image(data, ext, max_px, quality)
                    if res:
                        new_data, new_ext = res
                        # 只有同扩展名才能原地替换，否则 rels 的 Target 会失效
                        if new_ext == ext or (ext in (".jpg", ".jpeg") and new_ext == ".jpeg"):
                            saved_media += len(data) - len(new_data)
                            data = new_data
                out.writestr(item, data)

    before, after = src.stat().st_size, tmp.stat().st_size
    if verbose:
        bits = [f"删孤儿 {len(dropped)} 个", f"图片省 {saved_media/1048576:.1f} MB"]
        if saved_fonts:
            bits.append(f"内嵌字体省 {saved_fonts/1048576:.1f} MB")
        print(f"  {src.name}: {before/1048576:.1f} MB -> {after/1048576:.1f} MB ({'，'.join(bits)})")
    if inplace:
        shutil.move(str(tmp), str(src))
        return src
    return tmp


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("path")
    ap.add_argument("--max-px", type=int, default=2400)
    ap.add_argument("--quality", type=int, default=82)
    ap.add_argument("--inplace", action="store_true")
    ap.add_argument("--keep-fonts", action="store_true", help="保留内嵌字体（体积会暴涨，但换机零风险）")
    a = ap.parse_args()
    slim(Path(a.path), a.max_px, a.quality, a.inplace, strip_fonts=not a.keep_fonts)


if __name__ == "__main__":
    main()
