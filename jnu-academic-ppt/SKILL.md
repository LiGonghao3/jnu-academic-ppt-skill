---
name: jnu-academic-ppt
description: 为暨南大学的组会、论文精读、开题、中期、毕业或竞赛答辩创建和修改学术 PPT。使用学校主题、整理周记或素材、生成可编辑幻灯片并准备模拟导师问答。仅在用户明确需要暨大风格或使用本技能提供的主题时触发；普通商务或其他学校 PPT 不使用。
---

# 暨大学术汇报 PPT

把零散材料整理成能讲清楚、能追溯证据的学术汇报，并使用 `themes/` 中的暨大主题。正文、表格和主要内容对象应保持可编辑；照片、论文原图和模板背景仍是图片。

## 先判断运行环境

开始制作前读 [平台兼容与工具路由](references/platforms.md)，遵守宿主已有的演示文稿和生图规则。

- Codex：优先使用宿主的 Presentations 与 image generation 能力，导入相应 `stencil.pptx` 并读取 `theme.json`。不要绕过宿主要求的演示文稿工具。
- Claude 或允许本地 Python 的环境：可以使用 `scripts/build_deck.py` 和 `scripts/check_deck.py`。
- 没有原生生图工具时，才使用 `scripts/gen_figure.py` 调 OpenAI Images API。

## 工作流

### 1. 识别场景和时长

| 场景 | 叙事主线 | 默认主题 |
|---|---|---|
| 每周组会 / 进展汇报 | 进展、证据、卡点、下一步 | `jnu-teal` |
| 论文精读 / 文献汇报 | 问题、方法、证据、局限、对本组的启发 | `jnu-teal` |
| 开题 / 中期 / 项目答辩 | 背景、研究内容、技术路线、进度与计划 | `jnu-defense` |
| 毕业论文答辩 | 背景、方法、结果、总结与不足 | `jnu-rigor` |
| 党团 / 思政 / 活动汇报 | 背景、做法、成效、下一步 | `jnu-red` |
| 社团或轻松场合 | 按材料决定 | `jnu-crisp` |

短组会默认使用紧凑模式：10 分钟以内不单独放目录和章节过渡页；答辩和长报告可保留。

### 2. 收集材料

首次使用或用户不知道怎么准备材料时，读 [用户使用教程](references/user-guide.md)，在对话中简短说明四种入口：周记流水账、素材文件夹、单篇论文、直接口述。用户已经给足材料时不要重复询问，只补问会改变结果的问题。

开始追问前读 [内容采集协议](references/intake.md)。一次问 2–4 个相关问题，优先确认：

- 汇报对象、时长和场景
- 本次最重要的主线
- 结论对应的数字、图表或来源
- 卡点、失败尝试和希望得到的帮助
- 下周可交付物或答辩风险

不要编数字、结论、引用或实验条件。信息确实不足时，做短版或请用户补材料，不要注水。

### 3. 确认叙事

读 [叙事与页数](references/narrative.md)，先在对话中给出章节、每页标题和页面目的。只在结构仍有实质分歧时等待确认；用户已经明确要求直接制作时可继续。

- 结果页优先用有证据支持的结论标题。
- 背景、定义、实验设置和流程页可以用直接的主题标题，不强行写成结论。
- 内容页至少应包含“结论或主题 + 证据/机制 + 解释或用途”中的两项。
- 卡点和求助是组会的独立重点；局限与风险是答辩的独立重点。

### 4. 准备内容与视觉

使用本地引擎时，按 [deck.json 规范](references/deck-schema.md) 写内容，并遵守 [排版与信息密度](references/slide-rules.md)。图表、公式、截图和生成图都要保留来源。

需要概念图、流程图或封面装饰时读 [配图规范](references/figures.md)。优先使用宿主原生生图工具；AI 图不得伪装成实验结果，图注必须标明“示意图（AI 生成）”。

### 5. 生成、校验和预览

本地 Python 路线使用当前环境可用的 Python，不写死机器路径。脚本在**本技能目录**下，而命令通常在用户的项目目录里执行，所以要用本 SKILL.md 所在目录拼出脚本的绝对路径（下文记作 `<SKILL_DIR>`）；deck.json 和输出文件留在用户的工作目录：

```bash
python <SKILL_DIR>/scripts/build_deck.py .jnuppt/report.deck.json --out report.pptx --theme jnu-teal
python <SKILL_DIR>/scripts/check_deck.py report.pptx --theme jnu-teal
```

第一次运行前确认依赖：`python -m pip install -r <SKILL_DIR>/requirements.txt`。Windows 上 `python` 可能只是微软商店的占位程序，打不开时改用 `py`、conda 或虚拟环境里的解释器。

有 PowerPoint 时可以追加 `--deep`。报告必须明确区分“静态检查”和“PowerPoint 实测”；深度检查不可用时不能宣称已经实测。

- ERROR 必须修复。
- 缺图默认导致构建失败；只有草稿阶段才可显式使用 `--allow-missing-assets`。
- 预览每一页，检查文字、裁图、图注、页码、主题匹配和节奏。
- 12 页以上至少使用三种适合内容的版式，但不要为了变化滥用卡片。

### 6. 交付与模拟导师问答

读 [预测提问规范](references/qa.md)。问答以对话框为主，PPT 备注为便携备份：

1. 在对话中给完整版。组会 5–8 题，答辩 8–12 题；每题写问题、回答思路、仍需补什么，并尽量标页码。
2. 将精简版写入结尾页备注，控制在 600 字内。用户导出 PDF 或主要使用 WPS/手机时，提醒备注可能不可见，不得只靠备注交付。
3. 不替用户编答案。缺少实验口径、样本量或对比依据时明确标“需补充”。

默认只交付 PPTX；`deck.json` 保留在工作目录用于换主题和后续修改。主题切换后必须重新校验，语义内容通常可复用，但标题换行、导航容量和图片裁切可能需要微调。

## 按需读取

- [平台兼容与工具路由](references/platforms.md)：每次制作前
- [用户使用教程](references/user-guide.md)：首次使用或用户需要帮助时
- [内容采集协议](references/intake.md)：采集材料时
- [叙事与页数](references/narrative.md)：设计大纲时
- [deck.json 规范](references/deck-schema.md)：使用本地引擎时
- [排版与信息密度](references/slide-rules.md)：写页面内容时
- [主题规范](references/themes.md)：选主题、维护预览或接入新模板时
- [配图规范](references/figures.md)：取图、生图或编辑图片时
- [预测提问规范](references/qa.md)：交付前
