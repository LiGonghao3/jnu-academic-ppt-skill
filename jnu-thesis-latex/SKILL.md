---
name: jnu-thesis-latex
description: 用 LaTeX 按暨南大学本科毕业论文（设计）官方格式排版论文：新建项目、填写封面、把草稿或 Word 内容整理成章节、处理图表公式与 GB/T 7714 参考文献、编译并排查问题。仅在用户明确要写或排版暨大本科毕业论文、且选择 LaTeX 时触发；Word 排版、其他学校或研究生学位论文不使用。
---

# 暨大本科毕业论文 LaTeX 排版

本技能附带一套逐页对照学校官方 Word 模板（2026-04-13 版）校准的 LaTeX 模板，位于 `assets/template/`（GPL-3.0，来源见其中的 `NOTICE.md`）。你的任务是帮用户把内容放进这套模板，而不是改模板的版式。

## 先确认三件事

1. **编译环境**：本地有 TeX Live / MiKTeX / MacTeX，还是用 Overleaf。没有任何环境时，推荐 Overleaf，编译器选 XeLaTeX。
2. **封面信息**：题目（中英文）、学院、学系、专业、姓名、学号、指导教师、日期。缺的先留空，不要编造。
3. **内容来源**：已有 Word 稿、Markdown 笔记，还是从零写。已有内容时只做格式迁移，不擅自改写观点和数据。

## 工作流

### 1. 新建项目

脚本在本技能目录下，下文记作 `<SKILL_DIR>`；论文项目放在用户的工作目录：

```bash
python <SKILL_DIR>/scripts/new_thesis.py 我的论文 --title "中文题目" --en-title "English Title" \
    --college 学院 --department 学系 --major 专业 --name 姓名 --student-id 学号 \
    --advisor 指导教师 --date 2027-05-20 --fontset windows
```

脚本不会覆盖已有文件。Windows 定稿用 `--fontset windows`，Overleaf 用 `fandol`，macOS 保持默认 `auto`。

### 2. 填内容

按 [写作指引](references/writing-guide.md) 替换 `jnuthesis.tex` 里的示例内容。要点：

- 删掉模板里的格式说明文字（【注意事项】【插图规范】等），并按指引清空目录页说明。
- 章节用 `\chapter` / `\section` / `\subsection`，最多四级；参考文献、附录、致谢用专门命令。
- 引用写 `\supercite{key}`，文献放进 `refs.bib`，并启用 `\printbibliography[heading=jnubib]`。
- 从 Word 迁移时，图片另存为独立文件放进 `figs/`，表格重建为 `tabular`，公式改写为 LaTeX。

### 3. 编译与检查

```bash
python <SKILL_DIR>/scripts/build_thesis.py jnuthesis.tex
```

脚本会先试 `latexmk`，不可用时（Windows 上 MiKTeX 常缺 Perl）自动改用 `xelatex → biber → xelatex ×2`。退出码为 1 时必须处理：编译错误、未定义的引用或文献、缺字。行溢出只是提示，需打开 PDF 确认是否出界。其他问题见 [常见问题](references/troubleshooting.md)。

### 4. 交付前核对

对照 [格式规范要点](references/format-spec.md) 检查：题目不超过 20 字，中文摘要不超过 300 字，关键词 3–5 个，图表按章编号且连续，参考文献为顺序编码制。

最后必须提醒用户：模板是社区整理的非官方版本，学院可能另有要求，提交前要把 PDF 与学院当年下发的官方模板逐页比对。`assets/official/` 里有官方 Word 原件可供对照。

## 不要做的事

- 不改 `jnuthesis.cls` 的版式参数来“修”显示问题，除非用户明确要求并了解后果。
- 不编造参考文献条目、实验数据或致谢对象。
- 不删除 `NOTICE.md` 和 `LICENSE`：模板以 GPL-3.0 发布，分发时必须保留。
