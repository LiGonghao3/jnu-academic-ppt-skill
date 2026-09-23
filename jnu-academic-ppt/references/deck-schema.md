# deck.json 规范

一份 `deck.json` 描述**内容**，一套主题描述**样式**。
两者完全解耦：换 `--theme` 重跑，同一份内容就是另一套皮。

```jsonc
{
  "theme": "jnu-teal",              // 可选，也可以用 --theme 覆盖
  "asset_dir": "./素材",             // 可选，图片相对路径的根；默认 deck.json 所在目录
  "meta": { ... },                  // 封面和落款信息
  "chapters": [ ... ],              // 章节，决定导航条和目录页
  "slides": [ ... ]                 // 页面序列
}
```

## meta

```jsonc
"meta": {
  "title": "多模态对比学习在病理切片分类中的应用",  // 封面主标题
  "subtitle": "第 12 周组会汇报",                  // 封面副标题
  "presenter": "张同学",                           // 汇报人（示例请用中性姓名）
  "affiliation": "信息科学技术学院 2024 级",
  "advisor": "李老师 教授",                        // 有就写，答辩场景必写
  "date": "2026-09-22",
  "duration_minutes": 10,                           // 用于规划页数
  "audience": "导师与课题组同学",
  "presentation_mode": "compact",                 // compact 会跳过 toc/section 实体页
  "file_name": "第12周组会"                        // 不指定 --out 时的文件名
}
```

封面署名默认按主题的 `meta_format` 输出。需要完全自定义时写：

```jsonc
"meta_lines": ["汇报人：张同学｜信息科学技术学院", "指导教师：李老师", "2026 年 5 月"]
```

`presentation_mode` 可选 `compact` 或 `standard`。紧凑模式仍保留章节归属和导航高亮，只是不生成目录与章节过渡页。不写时按 `duration_minutes` 判断：10 分钟以内默认 `compact`，其余默认 `standard`；显式填写的值永远优先。

## chapters

决定**目录页内容**和**每页导航条上的章节标记**。

```jsonc
"chapters": [
  {"title": "本周主线", "title_en": "This Week"},   // title_en 可选，jnu-rigor 会显示
  {"title": "卡点与求助"}
]
```

不写 `chapters` 时，会自动从所有 `layout: "section"` 的页面推导。
**章节数建议 3–5，最多 6**（各主题导航条容量见 `themes.md`）。

## slides

数组，按顺序出页。每项必有 `layout`。
所有版式都支持这两个通用字段：

| 字段 | 说明 |
|---|---|
| `notes` | 演讲备注，写进 PPT 备注栏。放细节、数据出处、口播提示 |
| `chapter` | 手动指定本页属于第几章（1-based）。不写就继承上一个 `section` |
| `source` | 本页证据或图片来源；显示在页脚并自动并入备注 |

内容页（非 cover/toc/section/closing）还支持：

| 字段 | 说明 |
|---|---|
| `title` | 页标题。**写结论，不写话题** |
| `kicker` | 标题上方的小字标签，如「结论先行」「可复现性」 |
| `rule` | 标题下的分隔线，默认 `true` |

---

# 16 种版式

## cover —— 封面

```jsonc
{"layout": "cover"}
```
默认全部取 `meta`。要覆盖时可写 `title` / `subtitle` / `meta`（字符串数组）。
颜色由主题的 `defaults.cover` 决定，一般不用管。

## toc —— 目录

```jsonc
{"layout": "toc", "title": "本次汇报", "kicker": "CONTENTS"}
```
内容自动来自 `chapters`。超过 4 章会自动分两栏。

## section —— 章节过渡页

```jsonc
{"layout": "section", "title": "本周主线", "subtitle": "把 baseline 跑通并定位了掉点原因"}
```
序号自动按出现顺序编。`subtitle` 可选（也可用 `title_en`）。
**每个 section 会把后续内容页归到这一章**，导航条据此高亮。

## bullets —— 结论 + 要点（最常用）

```jsonc
{"layout": "bullets",
 "title": "掉点来自数据增强，不是模型",
 "kicker": "结论先行",
 "lead": "把裁剪 scale 下限从 0.08 提到 0.4，Top-1 回到 78.2%。",  // 可选，高亮的一句话结论
 "bullets": ["第一条", "第二条", "第三条"],
 "bullet_char": "▪"}                                              // 可选
```
**要点 3–5 条，每条不超过 40 字**。超了就拆页。

## two-col —— 双栏要点

```jsonc
{"layout": "two-col", "title": "消融实验：哪一项在起作用",
 "columns": [
   {"title": "有效的改动", "bullets": ["...", "..."]},
   {"title": "试过但没用的", "bullets": ["...", "..."]}
 ]}
```
每栏也可以用 `"text": "一段话"` 代替 `bullets`。

## compare —— 对比（前后 / 我们 vs 别人）

```jsonc
{"layout": "compare", "title": "改动前后的失败案例对比",
 "sides": [
   {"title": "改动前：正样本对失配", "bullets": ["...", "..."], "color": "risk"},
   {"title": "改动后", "bullets": ["...", "..."], "color": "ok"}
 ]}
```
不写 `color` 时左侧默认 `risk`（红）、右侧 `ok`（绿）。

## cards —— 并列卡片（2–4 张）

```jsonc
{"layout": "cards", "title": "三件事想听听大家意见",
 "cards": [
   {"tag": "01", "title": "显存不够", "bullets": ["batch 512 时 OOM", "梯度累积还是换卡？"]},
   {"tag": "02", "title": "标注噪声", "text": "约 3% 切片标签存疑"}
 ]}
```
`tag` 不写就自动编号。卡片高度跟着内容走，内容少不会撑满留白。

## steps —— 有序流程

```jsonc
{"layout": "steps", "title": "标注复核的三步走",
 "steps": [
   {"title": "置信度排序", "text": "按 margin 取最可疑的 200 张"},
   {"title": "双人盲标", "text": "我和师兄各标一遍"}
 ]}
```
卡片之间自动画箭头。**3–5 步最佳**。

## timeline —— 时间轴

```jsonc
{"layout": "timeline", "title": "下周四天排期",
 "steps": [
   {"label": "周一", "title": "切片级评测", "text": "对齐论文口径"},
   {"label": "周二—三", "title": "标注复核", "text": "抽 200 张人工过"}
 ]}
```
节点在轴线上下交替排布。**3–6 个节点**，多了会挤。

## kpi —— 关键数字（2–4 个）

```jsonc
{"layout": "kpi", "title": "三个种子的复现结果",
 "items": [
   {"value": "78.2%", "label": "Top-1 准确率", "note": "论文 78.5%"},
   {"value": "±0.2", "label": "三种子标准差", "note": "此前 ±1.9", "color": "ok"}
 ]}
```
`value` 会自动放大到合适字号。`note` 用来放对照基线。

## table —— 表格

```jsonc
{"layout": "table", "title": "与三个公开方法的对比",
 "header": ["方法", "Top-1", "F1", "训练时长"],
 "rows": [["SimCLR", "71.4%", "0.68", "4.0h"],
          ["本周复现", "78.2%", "0.76", "4.1h"]]}
```
**不超过 6 行 4 列**。更大的表格应该截图原表，或者只挑关键行。
表头自动用主题主色，隔行浅底。

## chart —— 原生可编辑图表

```jsonc
{"layout": "chart", "title": "修正增强后准确率恢复到论文水平",
 "lead": "三个随机种子的均值为 78.2%。",
 "chart": {
   "type": "line",                         // line / column / bar
   "categories": ["原始复现", "修正增强", "论文报告"],
   "series": [{"name": "Top-1 (%)", "values": [71.4, 78.2, 78.5]}],
   "x_title": "设置", "y_title": "Top-1 (%)",
   "show_legend": false
 },
 "source": "本实验，3 个随机种子；论文报告值见 Table 2"}
```

只能填入真实数据。标题、单位、基线、样本量/随机种子和来源应在 `lead`、轴标题、`source` 或备注中说清。更复杂的误差条、散点图和统计标记优先由宿主原生演示文稿工具创建；本地引擎的 `chart` 版式只覆盖常用折线、柱形和条形图。

## image-text —— 图文对照

```jsonc
{"layout": "image-text", "title": "注意力可视化",
 "image": "cam.png",              // 相对 asset_dir
 "image_side": "right",           // right(默认) / left
 "image_ratio": 0.5,              // 图占正文区宽度比例，默认 0.5
 "image_mode": "fit",             // fit(默认，保比例不裁) / fill(裁剪填满)
 "caption": "改动前后的 CAM 对比",
 "lead": "注意力从背景收敛到病灶中心。",
 "bullets": ["左：改动前", "右：改动后"]}
```
文字侧也可以用 `"text": "一段话"`。图找不到时默认停止导出，避免把占位框误交付；
只有显式使用 `--allow-missing-assets` 调试时才会画占位框并报警告。

## image-full —— 整图

```jsonc
{"layout": "image-full", "title": "总体技术路线",
 "image": "pipeline.png", "caption": "图 1 系统总体框架"}
```

## quote —— 一句话结论 / 引用

```jsonc
{"layout": "quote", "title": "本周最大的教训",
 "text": "复现掉点先查数据管线，别一上来就改模型。",
 "source": "第 12 周复盘"}
```

## blank —— 只有标题的空页

```jsonc
{"layout": "blank", "title": "留白页"}
```
自己后期往里放东西时用。

## closing —— 结尾

```jsonc
{"layout": "closing", "title": "谢谢", "subtitle": "欢迎提问 / 张同学 · 2026-09-22"}
```
`jnu-crisp` 的结尾外壳自带「谢谢观看」行楷艺术字，标题位已经被占了，
所以这套主题会忽略 `title` 只补副标题。确实要在艺术字上再叠一行，写 `"force_title": true`
（体检会报「压到模板文字」，那是预期内的）。

---

# 颜色 token

所有 `color` 类字段都可以写 token 名或 `#RRGGBB` 直写。各主题都提供这套 token：

`primary` `primary_dark` `secondary` `accent` `surface` `surface_alt`
`line` `ink` `muted` `white` `ok` `warn` `risk`

**优先用 token 而不是硬编码颜色**，否则换主题时颜色不会跟着变。

---

# 完整示例

见 `examples/组会示例.deck.json`（每周组会）和 `examples/答辩示例.deck.json`（毕业答辩）。

## qa —— 模拟导师问答

顶层可写结构化 `qa`。构建器会把精简版自动合并进第一张 `closing` 页的备注，但 Agent 仍应在交付对话中给出完整版。

- 问答备份单独占 600 字预算，按整题取舍：每条回答最多保留 120 字，放不下的题会在备注末尾注明「另有 N 题见对话」并在构建时报警告。
- 页面自己的 `notes`（讲稿备注）原样保留，不计入这 600 字，也不会被截断。
- 没有 `closing` 页时，问答不会写进任何备注，构建时会报警告。

```jsonc
"qa": [
  {"slide": 7, "question": "为什么只跑三个随机种子？",
   "answer": "说明资源约束、方差和补充实验计划。",
   "gap": "确认是否需要五种子结果"}
]
```
