# 常见问题

## 编译环境

| 现象 | 原因与处理 |
|---|---|
| `latexmk` 报 “could not find the script engine 'perl'” | Windows 上的 MiKTeX 自带 latexmk，但需要 Perl。直接用 `build_thesis.py`，它会自动改用 xelatex → biber → xelatex ×2；或者安装 Strawberry Perl |
| 中文不显示、满屏报错 | 用了 pdfLaTeX。本模板必须用 **XeLaTeX** |
| 找不到 `xelatex` | 没装 TeX 发行版。Windows 装 MiKTeX 或 TeX Live，macOS 装 MacTeX；不想装就用 Overleaf |
| MiKTeX 首次编译卡住或弹窗 | 正在自动下载缺失的宏包（如 biblatex-gb7714-2015），允许安装后重新编译 |
| Overleaf 编译失败 | Menu → Compiler 选 XeLaTeX，文档类加 `fontset=fandol` |

## 字体

| 现象 | 原因与处理 |
|---|---|
| 日志出现 “Missing character: There is no …” | 当前字体方案缺这个字形。Windows 用 `fontset=windows`；Overleaf 用 `fandol`，或上传字体后用 `bundled` |
| 字形和 Word 不像 | Windows 定稿用 `fontset=windows`（SimSun、SimHei，与官方 Word 一致）；macOS 保持默认 `auto` |
| 需要跨平台完全一致 | 从 Windows 的 `C:\Windows\Fonts\` 复制字体到 `fonts/`，用 `fontset=bundled`，见 `fonts/README.md` |

## 参考文献与引用

| 现象 | 原因与处理 |
|---|---|
| 文献表是空的 | 没有运行 biber，或 `\printbibliography[heading=jnubib]` 还是注释状态 |
| 正文出现粗体的 key（如 **vaswani2017**） | `refs.bib` 里没有这个 key，或 key 拼写不一致 |
| 正文出现 “??” | `\ref` 的标签不存在，或只编译了一遍。先修正标签，再完整编译 |
| 公式引用印成“式（式 3-1）” | 公式编号已带“式”字，正文直接写 `\eqref{…}` |

## 版面

| 现象 | 原因与处理 |
|---|---|
| Overfull \hbox 提示 | 长网址、长英文单词或宽表格超出版心。网址用 `\url{}`，表格缩小列宽或换行 |
| 图表漂到很远的位置 | 浮动体的正常行为。用 `[htbp]`，并在正文里用 `\ref` 引用，不要写“下图”“上表” |
| 目录页有一大段说明文字 | 模板默认保留了官方说明，在导言区写 `\renewcommand{\jnutocpreamble}{}` 清空 |
