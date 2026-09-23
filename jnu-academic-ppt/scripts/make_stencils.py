# -*- coding: utf-8 -*-
"""把原始 .pptx 模板提炼成主题包里的 stencil.pptx（只留「外壳页」）。

外壳页 = 背景 + 装饰 + logo + 导航底板，**不含任何正文和章节文字**。
正文由 build_deck.py 用原生形状画上去，所以同一套外壳能承载任意版式，
这就是「泛化」而不是「硬套模板」的地方。

用法：
    python make_stencils.py --templates "<原始模板文件夹>" [--theme jnu-teal] [--all]

只需在收编新模板 / 模板更新时跑一次；日常出片不需要。
"""
from __future__ import annotations

import argparse
import json
import os
import re
import sys
from pathlib import Path

from pptx import Presentation

sys.path.insert(0, str(Path(__file__).parent))
from ppt_kit import clone_slide, drop_slides, iter_shapes, remove_shape  # noqa: E402

ROOT = Path(__file__).resolve().parent.parent
SPEC = json.loads((Path(__file__).parent / "stencil_spec.json").read_text(encoding="utf-8"))
SHELL_PREFIX = "shell_"


def clear_text_deep(shape):
    for sh in iter_shapes(shape.shapes) if shape.shape_type == 6 else [shape]:
        if sh.has_text_frame:
            for p in sh.text_frame.paragraphs:
                for r in list(p.runs):
                    r.text = ""


def apply_shell(slide, rule):
    keep = set(rule.get("keep") or [])
    drop = set(rule.get("drop") or [])
    drop_nested = set(rule.get("drop_nested") or [])

    for sh in list(slide.shapes):
        if keep and sh.name not in keep:
            remove_shape(sh)
        elif sh.name in drop:
            remove_shape(sh)

    if drop_nested:
        for sh in list(iter_shapes(slide.shapes)):
            if sh.name in drop_nested:
                remove_shape(sh)

    for name in rule.get("clear_text") or []:
        for sh in iter_shapes(slide.shapes):
            if sh.name == name:
                clear_text_deep(sh)

    # 给外壳里的每个形状打上前缀。出片后 check_deck 靠它区分
    # 「模板自带的装饰」和「我们生成的内容」——否则模板的满幅底图、
    # 出血的艺术字会被当成排版事故一路报错。
    for sh in iter_shapes(slide.shapes):
        if not sh.name.startswith(SHELL_PREFIX):
            sh.name = SHELL_PREFIX + sh.name


PERSONAL_PROPS = re.compile(
    r"<(dc:creator|cp:lastModifiedBy)>[^<]*</\1>|<(dc:creator|cp:lastModifiedBy)/>")


def scrub_personal_props(path: Path) -> bool:
    """清掉 docProps/core.xml 里的原作者姓名，其余包内容逐字节保留。

    只动作者/最后修改者两个个人字段。custom.xml 里的 AIGC 隐式标识
    （部分模板由 AI 平台生成时写入）按规定不得删除，这里刻意不碰。"""
    import shutil, tempfile, zipfile
    with zipfile.ZipFile(path) as zin:
        core = zin.read("docProps/core.xml").decode("utf-8")
        new = PERSONAL_PROPS.sub(
            lambda m: "<{0}></{0}>".format(m.group(1) or m.group(2)), core)
        if new == core:
            return False
        fd, tmp = tempfile.mkstemp(suffix=".pptx", dir=path.parent)
        os.close(fd)
        with zipfile.ZipFile(tmp, "w") as zout:
            for info in zin.infolist():
                data = new.encode("utf-8") if info.filename == "docProps/core.xml" \
                    else zin.read(info.filename)
                zout.writestr(info, data)
    shutil.move(tmp, path)
    return True


def build_theme(theme_id, cfg, templates_dir: Path):
    src = templates_dir / cfg["source"]
    if not src.exists():
        raise SystemExit(f"找不到原始模板：{src}")

    prs = Presentation(str(src))
    original_count = len(prs.slides)

    manifest = {}
    for order, rule in enumerate(cfg["shells"]):
        src_slide = prs.slides[rule["slide"] - 1]
        new = clone_slide(prs, src_slide)
        apply_shell(new, rule)
        manifest[rule["id"]] = original_count + order

    # 删掉全部原始页，只留刚追加的外壳页
    drop_slides(prs, list(range(original_count)))

    out_dir = ROOT / "themes" / theme_id
    out_dir.mkdir(parents=True, exist_ok=True)
    out = out_dir / "stencil.pptx"
    prs.save(str(out))
    scrub_personal_props(out)

    # 重新索引：外壳页现在从 0 开始，顺序与 spec 一致
    index = {rule["id"]: i for i, rule in enumerate(cfg["shells"])}
    (out_dir / "shells.json").write_text(
        json.dumps({"shells": index, "slide_count": len(cfg["shells"])},
                   ensure_ascii=False, indent=2), encoding="utf-8")

    size_mb = out.stat().st_size / 1024 / 1024
    print(f"  {theme_id:14s} -> {out.name}  {len(cfg['shells'])} 页外壳  {size_mb:.1f} MB  {index}")
    return out


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--templates", required=True, help="原始 .pptx 模板所在文件夹")
    ap.add_argument("--theme", help="只处理某一套主题")
    ap.add_argument("--scrub-only", action="store_true",
                    help="不重新提炼，只清掉现有 stencil.pptx 里的原作者姓名")
    args = ap.parse_args()

    if args.scrub_only:
        ids = [args.theme] if args.theme else list(SPEC["themes"])
        for tid in ids:
            p = ROOT / "themes" / tid / "stencil.pptx"
            print(f"  {tid:14s} {'已清理' if scrub_personal_props(p) else '无需处理'}")
        return

    tdir = Path(args.templates)
    todo = {args.theme: SPEC["themes"][args.theme]} if args.theme else SPEC["themes"]
    print(f"从 {tdir} 提炼外壳：")
    for tid, cfg in todo.items():
        build_theme(tid, cfg, tdir)


if __name__ == "__main__":
    main()
