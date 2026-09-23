# 来源与许可

本目录中的 LaTeX 模板（`jnuthesis.cls`、`jnuthesis.tex`、`refs.bib`、`latexmkrc`、`figs/`、`fonts/README.md`）来自：

- 项目：[SolarAscent/JNU-Thesis-LaTeX-Template](https://github.com/SolarAscent/JNU-Thesis-LaTeX-Template)（v2.1，2026-06-14）
- 许可：GNU General Public License v3.0，全文见同目录 [`LICENSE`](LICENSE)

本目录下的文件继续以 GPL-3.0 发布，**不适用**仓库根目录的 MIT 许可。

## 相对上游的修改（2026-06）

依据学校《本科毕业论文（设计）装订要求及撰写规范》和官方 `.dot` 模板实测，对 `jnuthesis.cls` 做了两处修改：

1. 目录中章一级条目的字号由 15bp 改为 14bp（四号），与官方 `.dot` 实测一致。
2. 重定义 `\supercite`，让上标引用带方括号并压缩连续编号（如 [5-6]），对应规范三（六）1“标注的符号为‘[ ]’，作为上标”。

`jnuthesis.tex`、`refs.bib`、`latexmkrc` 与上游一致。

## 其他素材

- `figs/` 中的校徽和校名书法图取自学校官方封面文件，版权归暨南大学所有，仅供本校学生撰写毕业论文使用。
- 上一级 `official/` 目录是学校官方下发的 Word 模板、封面和撰写规范原件，同样归暨南大学所有，放在这里供核对格式。
