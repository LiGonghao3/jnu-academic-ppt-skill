# -*- coding: utf-8 -*-
"""从本技能自带的暨大本科毕业论文 LaTeX 模板新建一个论文项目，并填好封面信息。

不会覆盖已有文件：目标目录里只要有同名文件就整体放弃，列出冲突后退出。

用法：
    python new_thesis.py 我的论文 --title "基于……的研究" --name 张三 --student-id 2022000000 \\
        --college 信息科学技术学院 --department 计算机科学系 --major 计算机科学与技术 \\
        --advisor 李四 --date 2027-05-20 --fontset windows
"""
from __future__ import annotations

import argparse
import re
import shutil
import sys
from pathlib import Path

TEMPLATE = Path(__file__).resolve().parent.parent / "assets" / "template"
FILES = ["jnuthesis.cls", "jnuthesis.tex", "refs.bib", "latexmkrc", "LICENSE", "NOTICE.md",
         "figs/jnu-logo.png", "figs/jnu-name.png", "fonts/README.md"]
FONTSETS = ["auto", "windows", "mac", "noto", "fandol", "bundled"]

# 元数据参数 -> 模板里的宏
FIELDS = {"college": "xueyuan", "department": "xuexi", "major": "zhuanye",
          "name": "xingming", "student_id": "xuehao", "advisor": "daoshi"}

_ESCAPE = {"\\": r"\textbackslash{}", "&": r"\&", "%": r"\%", "$": r"\$", "#": r"\#",
           "_": r"\_", "{": r"\{", "}": r"\}", "~": r"\textasciitilde{}",
           "^": r"\textasciicircum{}"}


def tex_escape(s: str) -> str:
    return "".join(_ESCAPE.get(ch, ch) for ch in s)


def fill(tex: str, a) -> str:
    def macro(name, value):
        nonlocal tex
        # 形如 "\xueyuan{}        % 学院"：只替换花括号里的内容，保留行尾注释
        tex, n = re.subn(r"^(\\%s)\{[^}\n]*\}" % name,
                         lambda m: "%s{%s}" % (m.group(1), tex_escape(value)),
                         tex, count=1, flags=re.M)
        if n == 0:
            sys.exit(f"模板里找不到 \\{name}{{}}，模板可能已被改动")

    if a.title:     # 题目默认是注释掉的占位，取消注释再填
        tex = re.sub(r"^%\s*\\biaoti\{[^}\n]*\}", lambda m: r"\biaoti{%s}" % tex_escape(a.title),
                     tex, count=1, flags=re.M)
    if a.en_title:
        tex = re.sub(r"^%\s*\\entitle\{[^}\n]*\}", lambda m: r"\entitle{%s}" % tex_escape(a.en_title),
                     tex, count=1, flags=re.M)
    for arg, name in FIELDS.items():
        value = getattr(a, arg)
        if value:
            macro(name, value)
    if a.date:
        m = re.fullmatch(r"(\d{4})-(\d{1,2})-(\d{1,2})", a.date)
        if not m:
            sys.exit("--date 格式应为 YYYY-MM-DD")
        y, mo, d = (int(x) for x in m.groups())
        tex = re.sub(r"^%?\s*\\thesisdate\{[^}\n]*\}\{[^}\n]*\}\{[^}\n]*\}",
                     lambda _: r"\thesisdate{%d}{%d}{%d}" % (y, mo, d), tex, count=1, flags=re.M)
    if a.name:      # 填了姓名就启用正式封面；只改命令行本身，文件头注释里也出现了这个词
        tex = re.sub(r"^\\makeblankcover\b", lambda _: r"\makecover", tex, count=1, flags=re.M)
    if a.fontset:
        tex = re.sub(r"\\documentclass\[fontset=\w+\]\{jnuthesis\}",
                     lambda _: r"\documentclass[fontset=%s]{jnuthesis}" % a.fontset, tex, count=1)
    return tex


def main():
    ap = argparse.ArgumentParser(description="新建暨大本科毕业论文 LaTeX 项目")
    ap.add_argument("target", help="新项目目录（不存在会自动创建）")
    ap.add_argument("--title", help="中文题目（同时用于摘要页标题与页眉）")
    ap.add_argument("--en-title", help="英文题目")
    ap.add_argument("--college", help="学院")
    ap.add_argument("--department", help="学系")
    ap.add_argument("--major", help="专业")
    ap.add_argument("--name", help="姓名；填写后封面改用 \\makecover")
    ap.add_argument("--student-id", help="学号")
    ap.add_argument("--advisor", help="指导教师")
    ap.add_argument("--date", help="封面日期 YYYY-MM-DD")
    ap.add_argument("--fontset", choices=FONTSETS,
                    help="字体方案；Windows 定稿推荐 windows，Overleaf 推荐 fandol")
    a = ap.parse_args()

    target = Path(a.target)
    conflicts = [f for f in FILES if (target / f).exists()]
    if conflicts:
        sys.exit("目标目录已有同名文件，为避免覆盖已停止：\n  " + "\n  ".join(conflicts))
    for f in FILES:
        (target / f).parent.mkdir(parents=True, exist_ok=True)
        shutil.copy2(TEMPLATE / f, target / f)

    main_tex = target / "jnuthesis.tex"
    main_tex.write_text(fill(main_tex.read_text(encoding="utf-8"), a),
                        encoding="utf-8", newline="\n")
    print(f"已创建论文项目：{target.resolve()}")
    print("下一步：在 jnuthesis.tex 里写正文，然后运行 build_thesis.py 编译。")


if __name__ == "__main__":
    main()
