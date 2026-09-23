# 平台兼容与工具路由

这套 skill 的内容方法、主题 token、模板外壳和 `deck.json` 是通用层；真正生成 PPTX 和图片时，必须遵守当前宿主的工具约束。

## 能力检测顺序

1. 查看宿主是否已经提供演示文稿技能或工具。
2. 查看宿主是否提供原生图片生成/编辑工具。
3. 只有宿主允许本地脚本、且没有更高优先级工具要求时，才运行本项目的 Python 引擎。
4. 运行任何命令前使用当前环境的解释器，不假设磁盘盘符或固定安装路径。

## Codex / ChatGPT 桌面

- 同时使用本 skill 与宿主的 `Presentations` skill。
- 宿主要求使用 `@oai/artifact-tool` 时，以该要求为准；不要直接用 `python-pptx` 创建或编辑最终 PPTX。
- 将 `themes/<id>/stencil.pptx` 作为模板导入，读取同目录的 `theme.json` 获取画布、字体、颜色、正文区和导航规格。
- 复用模板外壳和母版，按 `deck.json` 或已确认大纲填入内容。生成后走宿主要求的渲染、结构检查与逐页视觉复核。
- 需要图片时优先调用 Codex 的 image generation 工具。图片生成后仍须写图注和出处记录。

Codex 会读取 `agents/openai.yaml` 作为界面元数据，但该文件不是其他宿主运行本 skill 的前提。

## Claude Code / Claude 桌面或普通本地 Agent

如果宿主允许 Python 脚本，可以运行：

```bash
python <SKILL_DIR>/scripts/build_deck.py deck.json --theme jnu-teal --out report.pptx
python <SKILL_DIR>/scripts/check_deck.py report.pptx --theme jnu-teal
```

`<SKILL_DIR>` 是本技能的安装目录（SKILL.md 所在处），不是用户的项目目录。

若命令名不是 `python`，使用当前环境实际可用的 `python3`、虚拟环境解释器或用户明确给出的解释器。不要把个人机器路径写进 skill。

## 图片生成路由

1. 宿主原生生图/修图工具可用：直接使用。它通常不需要用户额外配置 API key。
2. 原生工具不可用，但已配置 OpenAI API：使用 `scripts/gen_figure.py`。
3. 两者都不可用：保留明确的图片需求说明，使用真实现有图片或请用户补图；不要静默放占位图交付。

无论哪条路线都要遵守：AI 只能生成概念示意、封面装饰和不冒充现实证据的插图；实验图、结果曲线、混淆矩阵和任何具体数据必须来自真实数据或原始资料。

## 可移植性边界

- PowerPoint 深度测量和本机渲染是可选增强，不是所有平台都有。
- WPS、PDF 和手机预览可能看不到备注，因此模拟问答必须同时出现在对话中。
- 模板字体可能在其他系统回退。去陌生电脑汇报时应另存 PDF 备用，并在 PowerPoint 中检查字体替换。
- 换主题不保证像素级无改动；必须重新检查换行、裁图和导航容量。

