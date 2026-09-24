# 实验结果表与 LaTeX 三线表

## 什么时候使用

正式组会、论文汇报和答辩中，如果需要并列比较多个方法与多个指标，优先使用三线表。只有一两个关键数字时用 `kpi`，需要看趋势时用 `chart`，不要为了显得正式把所有结果都塞进表格。

选择顺序：

1. 当前宿主要求证据对象可编辑：创建 PowerPoint 原生表格，并把样式做成三线表（无竖线，只保留顶线、表头分隔线和底线）。
2. 使用本地 Python 引擎，且用户更在意论文式排版：使用 `three-line-table`，自动尝试 LaTeX `booktabs`。
3. LaTeX 编译器、中文字体、`booktabs` 或 `pdftocairo` 缺失：自动回退为 PowerPoint 原生可编辑表格，不中断整份 PPT。

LaTeX 成功时，PPT 中的表格默认以 600 dpi 透明 PNG 插入，不可直接编辑单元格；结构化数据仍保存在 `deck.json`，可以修改后重新生成。用户要求现场改数、协作编辑或无障碍读取时，应直接使用原生表格。

## 内容规则

- 建议不超过 6 行、5 列。超过时只保留与当前结论有关的方法和指标，完整表放备份页。
- 标题说明比较得到的结论，`lead` 交代数据集、随机种子、均值/方差或统计口径。
- 单位写进列名，例如 `Accuracy (%)`、`Latency (ms)`。
- 同一列使用一致的小数位。不要把 `0.803` 和 `80.3%` 混在一列。
- 只加粗真正需要听众比较的最佳值；并列最优应同时标出。不要用颜色制造不存在的显著性。
- 必须写 `source`。用户自己的实验要注明运行次数和评测口径；论文数据要给论文与表号。

## 本地工具链

检测：

```bash
python <SKILL_DIR>/scripts/render_latex_table.py --check
```

默认寻找 `xelatex`、`lualatex`、`pdflatex`，以及 `pdftocairo`。也可以用环境变量指定：

- `JNU_PPT_LATEX_ENGINE`
- `JNU_PPT_PDFTOCAIRO`
- `JNU_PPT_DISABLE_LATEX=1`：强制测试原生表格回退

脚本不会自动安装 TeX 包，也不会启用 shell escape。所有表格单元格按普通文本转义，不接受用户提供的原始 LaTeX 命令。

## deck.json 示例

```jsonc
{
  "layout": "three-line-table",
  "title": "本文方法在三个指标上均高于基线",
  "lead": "均值 ± 标准差，3 个随机种子。",
  "header": ["方法", "Accuracy (%)", "Macro-F1", "耗时 (h)"],
  "rows": [
    ["Baseline", "78.2 ± 0.3", "0.761 ± 0.004", "4.1"],
    ["Ours", "81.6 ± 0.2", "0.803 ± 0.003", "4.6"]
  ],
  "align": ["left", "center", "center", "center"],
  "bold_cells": [[2, 2], [2, 3]],
  "caption": "最佳结果加粗",
  "source": "本实验，3 个随机种子"
}
```

`bold_rows` 使用从 1 开始的数据行编号；`bold_cells` 使用 `[数据行, 列]`，同样从 1 开始。`latex_required: true` 会在 LaTeX 失败时终止构建，默认不要开启，因为本技能的正常策略是安全回退。
