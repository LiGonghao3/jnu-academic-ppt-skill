# 写作指引

`jnuthesis.tex` 是示例主文件，按官方模板的顺序排好了全部部件。写论文时保留结构，替换内容。

## 文档结构

```latex
\documentclass[fontset=windows]{jnuthesis}
\biaoti{中文题目}  \entitle{English Title}      % 题目同时用于摘要页标题和正文页眉
\xueyuan{…}\xuexi{…}\zhuanye{…}\xingming{…}\xuehao{…}\daoshi{…}
\thesisdate{2027}{5}{20}
\addbibresource{refs.bib}
\renewcommand{\jnutocpreamble}{}                 % 清空目录页的官方说明文字

\begin{document}
\makecover          % 封面（\makeblankcover 是空白占位版）
\makestatement      % 诚信声明
\begin{zhabstract} 中文摘要…… \zhaiyao{关键词1；关键词2；关键词3} \end{zhabstract}
\begin{enabstract} English abstract… \enkeywords{Keyword1; Keyword2; Keyword3} \end{enabstract}
\tableofcontents
\jnumainmatter      % 正文从这里开始用阿拉伯数字页码

\chapter{绪论} …
\printbibliography[heading=jnubib]
\jnuleftchapter{附录}   % 或用 \jnuappendix 后接 \chapter{…}，得到“附录A”
\jnucenterchapter{致谢}
\end{document}
```

## 需要删除的示例内容

模板正文里有官方说明文字，写论文时要删掉：

- 摘要、英文摘要里“编写摘要应注意……”“关键词是供检索用的……”等段落
- 【名词术语规范】【计量单位和数字表示】【注意事项】【插图规范】【公式规范】【参考文献规范】各段
- 示例表格（合金钢）、示例插图（驱油效率）和示例公式
- 附录、致谢里的说明文字

## 标题

| 命令 | 级别 | 编号 |
|---|---|---|
| `\chapter{}` | 一级，小三号宋体加粗 | 1 |
| `\section{}` | 二级，四号宋体加粗 | 1.1 |
| `\subsection{}` | 三级 | 1.1.1 |
| `\subsubsection{}` | 四级 | 1.1.1.1 |

规范规定最多四级标题，且两种编号体系（“一、（一）”与“1、1.1”）不能混用。模板用的是数字体系。

## 图、表、公式

```latex
\begin{figure}[htbp]
  \centering
  \includegraphics[width=0.8\textwidth]{figs/result.png}
  \caption{注水压力对驱油效率的影响}      % 图题在图下方，不加标点
  \label{fig:pressure}
\end{figure}

\begin{table}[htbp]
  \centering
  \caption{各方法在测试集上的准确率}       % 表题在表上方，不加标点
  \label{tab:acc}
  {\jnutablecell                           % 表内五号字
  \begin{tabular}{lcc} \hline 方法 & 准确率 & F1 \\ \hline … \\ \hline \end{tabular}}
\end{table}

\begin{equation}
  E = mc^2 \label{eq:energy}
\end{equation}
```

- 图表自动按章编号为“图 3-1”“表 3-1”，公式编号为“（式 3-1）”。
- 正文引用写 `图~\ref{fig:pressure}`、`表~\ref{tab:acc}`；**公式直接写 `\eqref{eq:energy}`**，不要再加“式”字，否则会印成“式（式 3-1）”。
- 表格中空缺的数据格填“—”，不要写“同上”。

## 参考文献

1. 把条目写进 `refs.bib`。中文作者写成 `作者1 and 作者2`，英文作者写成 `Last, First`。
2. 正文引用用 `\supercite{key}`，印成带方括号的上标，放在标点前，如“……机制\supercite{vaswani2017}。”；连续多篇会自动压缩为 [1-3]。
3. 启用 `\printbibliography[heading=jnubib]`，并删掉模板里“参考文献规范”的说明文字。
4. 格式由 biblatex-gb7714-2015 按 GB/T 7714-2015 顺序编码制生成，不要手动排文献表。

只引用用户提供或能核实的文献。拿不准的条目标出来让用户确认，不要补全不存在的卷期页码。

## 从 Word 迁移

- 图片：从 Word 里另存为 PNG/PDF 放进 `figs/`，矢量图优先保存为 PDF。
- 表格：重建为 `tabular`；合并单元格用 `\multirow` / `\multicolumn`（模板已加载）。
- 公式：改写为 LaTeX，行内用 `$…$`，独立公式用 `equation`。
- 文献：Word 里手写的文献表要逐条转成 `refs.bib` 条目，然后改用 `\supercite`。
