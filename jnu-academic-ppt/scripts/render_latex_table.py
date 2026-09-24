# -*- coding: utf-8 -*-
"""把结构化表格渲染成 LaTeX booktabs 三线表 PNG。

只接受普通文本单元格并逐字符转义，不执行用户提供的 LaTeX。需要 XeLaTeX/
LuaLaTeX（或 pdfLaTeX）和 pdftocairo；任一工具不可用时由调用方回退到
PowerPoint 原生表格。
"""
from __future__ import annotations

import argparse
import json
import os
import shutil
import subprocess
import tempfile
from pathlib import Path


class LatexTableUnavailable(RuntimeError):
    """本机没有完整工具链，或 LaTeX 表格未能安全编译。"""


def _disabled() -> bool:
    return os.environ.get("JNU_PPT_DISABLE_LATEX", "").strip().lower() in {
        "1", "true", "yes", "on"
    }


def _resolve_tool(explicit, env_name, candidates):
    raw = explicit or os.environ.get(env_name)
    if raw:
        path = Path(raw)
        if path.is_file():
            return str(path.resolve())
        found = shutil.which(str(raw))
        if found:
            return found
        raise LatexTableUnavailable(f"{env_name} 指定的工具不存在：{raw}")
    for name in candidates:
        found = shutil.which(name)
        if found:
            return found
    return None


def detect_toolchain(engine=None, converter=None):
    if _disabled():
        raise LatexTableUnavailable("环境变量 JNU_PPT_DISABLE_LATEX 已禁用 LaTeX")
    tex = _resolve_tool(engine, "JNU_PPT_LATEX_ENGINE",
                        ("xelatex", "lualatex", "pdflatex"))
    pdf = _resolve_tool(converter, "JNU_PPT_PDFTOCAIRO", ("pdftocairo",))
    if not tex:
        raise LatexTableUnavailable("找不到 XeLaTeX、LuaLaTeX 或 pdfLaTeX")
    if not pdf:
        raise LatexTableUnavailable("找不到 pdftocairo，无法把 PDF 表格转成 PNG")
    return tex, pdf


def escape_tex(value) -> str:
    """把单元格当普通文本处理，防止命令注入和意外公式解释。"""
    mapping = {
        "\\": r"\textbackslash{}", "&": r"\&", "%": r"\%",
        "$": r"\$", "#": r"\#", "_": r"\_", "{": r"\{",
        "}": r"\}", "~": r"\textasciitilde{}", "^": r"\textasciicircum{}",
    }
    text = str(value if value is not None else "").replace("\r", " ").replace("\n", " ")
    return "".join(mapping.get(ch, ch) for ch in text)


def _align_spec(ncol, align=None):
    choices = {"left": "l", "center": "c", "right": "r",
               "l": "l", "c": "c", "r": "r"}
    if isinstance(align, list):
        cols = [choices.get(str(v).lower(), "c") for v in align[:ncol]]
        cols.extend(["c"] * (ncol - len(cols)))
    else:
        cols = ["l"] + [choices.get(str(align or "center").lower(), "c")] * (ncol - 1)
    return "@{}" + "".join(cols) + "@{}"


def make_tex(header, rows, *, align=None, font_size=None,
             bold_rows=None, bold_cells=None,
             text_color="#1E2A2C", rule_color="#1E2A2C"):
    ncol = max(len(header or []), max((len(r) for r in rows or []), default=1))
    if ncol < 1:
        raise LatexTableUnavailable("表格没有列")
    if not rows and not header:
        raise LatexTableUnavailable("表格没有内容")
    allowed_sizes = {"normalsize", "small", "footnotesize", "scriptsize"}
    if font_size not in allowed_sizes:
        cells = ncol * (len(rows or []) + (1 if header else 0))
        font_size = "normalsize" if cells <= 20 else "small" if cells <= 32 else "footnotesize"

    bold_rows = {int(v) for v in (bold_rows or [])}
    bold_cells = {tuple(int(n) for n in pair) for pair in (bold_cells or [])
                  if isinstance(pair, (list, tuple)) and len(pair) == 2}

    def row(values, bold=False, row_index=None):
        padded = list(values or []) + [""] * (ncol - len(values or []))
        items = [escape_tex(v) for v in padded[:ncol]]
        items = [r"\textbf{" + v + "}" if bold or row_index in bold_rows
                 or (row_index, col + 1) in bold_cells else v
                 for col, v in enumerate(items)]
        return " & ".join(items) + r" \\"

    text_hex = str(text_color).lstrip("#")
    rule_hex = str(rule_color).lstrip("#")
    body = []
    if header:
        body.extend([row(header, True), r"\midrule"])
    body.extend(row(values, row_index=i) for i, values in enumerate(rows or [], 1))
    return "\n".join([
        r"\documentclass[border=8pt]{standalone}",
        r"\usepackage[UTF8,fontset=none]{ctex}",
        r"\IfFontExistsTF{Microsoft YaHei}{\setCJKmainfont{Microsoft YaHei}}{",
        r"  \IfFontExistsTF{Noto Sans CJK SC}{\setCJKmainfont{Noto Sans CJK SC}}{",
        r"    \IfFontExistsTF{Source Han Sans CN}{\setCJKmainfont{Source Han Sans CN}}{",
        r"      \IfFontExistsTF{SimSun}{\setCJKmainfont{SimSun}}{",
        r"        \IfFontExistsTF{PingFang SC}{\setCJKmainfont{PingFang SC}}{",
        r"          \PackageError{jnu-table}{No supported CJK font found}{Install a CJK font or use the native PowerPoint fallback.}",
        r"        }",
        r"      }",
        r"    }",
        r"  }",
        r"}",
        r"\usepackage{booktabs,array,xcolor}",
        rf"\definecolor{{TableText}}{{HTML}}{{{text_hex}}}",
        rf"\definecolor{{TableRule}}{{HTML}}{{{rule_hex}}}",
        r"\begin{document}",
        rf"\{font_size}",
        r"\color{TableText}",
        r"\renewcommand{\arraystretch}{1.35}",
        r"\setlength{\tabcolsep}{10pt}",
        rf"\begin{{tabular}}{{{_align_spec(ncol, align)}}}",
        r"\toprule",
        *body,
        r"\bottomrule",
        r"\end{tabular}",
        r"\end{document}",
        "",
    ])


def render_table(header, rows, output, *, align=None, font_size=None, dpi=600,
                 bold_rows=None, bold_cells=None,
                 text_color="#1E2A2C", rule_color="#1E2A2C",
                 engine=None, converter=None, timeout=90):
    tex_engine, pdf_converter = detect_toolchain(engine, converter)
    output = Path(output)
    output.parent.mkdir(parents=True, exist_ok=True)
    source = make_tex(header, rows, align=align, font_size=font_size,
                      bold_rows=bold_rows, bold_cells=bold_cells,
                      text_color=text_color, rule_color=rule_color)
    with tempfile.TemporaryDirectory(prefix="jnu-latex-table-") as tmp:
        tmpdir = Path(tmp)
        tex_path = tmpdir / "table.tex"
        tex_path.write_text(source, encoding="utf-8")
        cmd = [tex_engine, "-no-shell-escape", "-interaction=nonstopmode",
               "-halt-on-error", "-file-line-error", "-output-directory", str(tmpdir),
               str(tex_path)]
        try:
            run = subprocess.run(cmd, cwd=tmpdir, capture_output=True, text=True,
                                 encoding="utf-8", errors="replace", timeout=timeout)
        except subprocess.TimeoutExpired as exc:
            raise LatexTableUnavailable(f"LaTeX 编译超过 {timeout} 秒") from exc
        pdf_path = tmpdir / "table.pdf"
        if run.returncode != 0 or not pdf_path.exists():
            detail = (run.stdout or run.stderr or "LaTeX 未生成 PDF").strip()[-1200:]
            raise LatexTableUnavailable("LaTeX 编译失败：" + detail)
        prefix = tmpdir / "table-rendered"
        try:
            conv = subprocess.run(
                [pdf_converter, "-png", "-singlefile", "-transp", "-r", str(int(dpi)),
                 str(pdf_path), str(prefix)],
                cwd=tmpdir, capture_output=True, text=True, encoding="utf-8",
                errors="replace", timeout=timeout)
        except subprocess.TimeoutExpired as exc:
            raise LatexTableUnavailable(f"PDF 转 PNG 超过 {timeout} 秒") from exc
        png_path = prefix.with_suffix(".png")
        if conv.returncode != 0 or not png_path.exists():
            detail = (conv.stderr or conv.stdout or "pdftocairo 未生成 PNG").strip()[-800:]
            raise LatexTableUnavailable("PDF 转 PNG 失败：" + detail)
        shutil.copy2(png_path, output)
    return {"output": str(output), "engine": tex_engine, "converter": pdf_converter,
            "dpi": int(dpi)}


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("input", nargs="?", help="包含 header/rows 的 JSON 文件")
    ap.add_argument("--out", help="输出 PNG")
    ap.add_argument("--engine")
    ap.add_argument("--converter")
    ap.add_argument("--dpi", type=int, default=600)
    ap.add_argument("--check", action="store_true", help="只检测 LaTeX 工具链")
    args = ap.parse_args()
    if args.check:
        try:
            engine, converter = detect_toolchain(args.engine, args.converter)
            print(json.dumps({"available": True, "engine": engine,
                              "converter": converter}, ensure_ascii=False))
        except LatexTableUnavailable as exc:
            print(json.dumps({"available": False, "reason": str(exc)}, ensure_ascii=False))
            raise SystemExit(1)
        return
    if not args.input or not args.out:
        ap.error("需要 input 和 --out，或使用 --check")
    spec = json.loads(Path(args.input).read_text(encoding="utf-8"))
    result = render_table(spec.get("header") or [], spec.get("rows") or [], args.out,
                          align=spec.get("align"), font_size=spec.get("latex_font_size"),
                          bold_rows=spec.get("bold_rows"), bold_cells=spec.get("bold_cells"),
                          dpi=args.dpi, engine=args.engine, converter=args.converter)
    print(json.dumps(result, ensure_ascii=False))


if __name__ == "__main__":
    main()
