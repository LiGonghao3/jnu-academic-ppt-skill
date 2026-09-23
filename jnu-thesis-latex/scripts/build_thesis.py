# -*- coding: utf-8 -*-
"""编译暨大毕业论文，并把日志里真正要处理的问题挑出来。

先试 latexmk；latexmk 跑不起来时（Windows 上的 MiKTeX 常因缺 Perl 失败）
自动退回 xelatex → biber → xelatex → xelatex 手动流程。

用法：
    python build_thesis.py [jnuthesis.tex]
退出码：0=生成了 PDF 且干净；1=编译错误、没有 PDF、未定义的引用/文献或缺字。
"""
from __future__ import annotations

import argparse
import re
import shutil
import subprocess
import sys
from pathlib import Path


def run(cmd, cwd):
    r = subprocess.run(cmd, cwd=cwd, capture_output=True, text=True,
                       encoding="utf-8", errors="replace")
    return r.returncode, (r.stdout or "") + (r.stderr or "")


def latexmk_usable(cwd) -> bool:
    if not shutil.which("latexmk"):
        return False
    code, out = run(["latexmk", "-v"], cwd)
    return code == 0 and "Latexmk" in out


def build(tex: Path) -> bool:
    cwd, stem = tex.parent, tex.stem
    if not shutil.which("xelatex"):
        sys.exit("找不到 xelatex。请安装 TeX Live / MiKTeX / MacTeX，或把项目上传到 Overleaf，"
                 "编译器选 XeLaTeX。")
    xelatex = ["xelatex", "-interaction=nonstopmode", "-halt-on-error", "-file-line-error",
               tex.name]
    if latexmk_usable(cwd):
        print("使用 latexmk 编译……")
        code, _ = run(["latexmk", "-xelatex", "-interaction=nonstopmode", "-halt-on-error",
                       "-file-line-error", tex.name], cwd)
        return code == 0
    print("latexmk 不可用（常见原因：MiKTeX 缺 Perl），改用 xelatex → biber → xelatex ×2")
    code, _ = run(xelatex, cwd)
    if code != 0:
        return False
    if (cwd / f"{stem}.bcf").exists():
        if not shutil.which("biber"):
            print("  ! 找不到 biber，参考文献不会生成")
        else:
            code, out = run(["biber", stem], cwd)
            if code != 0:
                print("  ! biber 失败：\n" + "\n".join(out.strip().splitlines()[-5:]))
    for _ in range(2):
        code, _ = run(xelatex, cwd)
        if code != 0:
            return False
    return True


def report(tex: Path) -> bool:
    log_path = tex.with_suffix(".log")
    pdf = tex.with_suffix(".pdf")
    log = log_path.read_text(encoding="utf-8", errors="replace") if log_path.exists() else ""
    errors = [l for l in log.splitlines() if l.startswith("!") or re.search(r":\d+: ", l)]
    undef_ref = sorted(set(re.findall(r"Reference `([^']+)' on page \d+ undefined", log)))
    undef_cite = sorted(set(re.findall(r"Citation '([^']+)' on page \d+ undefined", log)))
    missing = sorted(set(re.findall(r"Missing character: There is no (\S+)", log)))
    overfull = len(re.findall(r"^Overfull \\hbox", log, flags=re.M))
    pages = re.search(r"Output written on .*?\((\d+) pages?", log)

    # 未定义的引用会在 PDF 里印成“??”，缺字会印成空白，都不能当成品交付
    ok = pdf.exists() and not (errors or undef_ref or undef_cite or missing)
    print(f"\nPDF：{pdf.name}（{pages.group(1)} 页）" if pages and pdf.exists()
          else "\n没有生成 PDF")
    for title, items, hint in [
        ("编译错误", errors[:8], "按行号修正后重新编译"),
        ("未定义的交叉引用", undef_ref, "检查 \\label 与 \\ref 的名字是否一致"),
        ("未定义的参考文献", undef_cite, "检查 refs.bib 里是否有这些 key，以及是否运行了 biber"),
        ("字体缺字", missing[:10], "当前字体方案缺这些字形，换 fontset 或安装对应字体"),
    ]:
        if items:
            print(f"✗ {title}（{hint}）：" if title == "编译错误" else f"! {title}（{hint}）：")
            for x in items:
                print("   " + x)
    if overfull:
        print(f"· 有 {overfull} 处行溢出（Overfull \\hbox），多为长网址或长公式，打开 PDF 核对是否出界")
    if ok:
        print("✓ 编译干净")
    return ok


def main():
    ap = argparse.ArgumentParser(description="编译暨大毕业论文并汇总问题")
    ap.add_argument("tex", nargs="?", default="jnuthesis.tex")
    a = ap.parse_args()
    tex = Path(a.tex).resolve()
    if not tex.exists():
        sys.exit(f"找不到 {tex}")
    built = build(tex)
    ok = report(tex) and built
    sys.exit(0 if ok else 1)


if __name__ == "__main__":
    main()
