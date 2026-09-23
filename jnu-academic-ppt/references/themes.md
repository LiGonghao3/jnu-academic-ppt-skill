# 主题包：选用、切换、收编新模板

## 五套主题

| id | 名称 | 主色 | 画布 | 导航形态 | 最适合 | 别用在 |
|---|---|---|---|---|---|---|
| `jnu-teal` | 大气通用 | `#145D68` 墨青 | 13.33×7.5 | 左侧竖排数字牌（≤6 章） | **每周组会**、信息密集的内容页 | 需要喜庆感的活动 |
| `jnu-defense` | 蓝色简约 | `#0F6A7B` 深青 | 13.33×7.5 | 顶部 Tab（≤5 章） | 开题/中期/竞赛答辩 | 章节超过 5 个 |
| `jnu-rigor` | 严谨有序 | `#4A639F` 蓝紫 | **20×11.25** | 顶部横排（≤5 章） | 毕业论文答辩、正式学术报告 | 轻松的周汇报 |
| `jnu-crisp` | 清爽简洁 | `#1D3F69` + 金 | 13.33×7.5 | 左上角「第N部分」标签 | 活动汇报、社团展示 | 严肃答辩（背景抢正文） |
| `jnu-red` | 红色主题 | `#C91313` | **20×11.25** | 左上角标签 | 党团/思政/红色文化类 | **学术组会和论文答辩** |

`jnu-rigor` 和 `jnu-red` 的画布是 20×11.25 英寸（仍是 16:9，只是基准更大）。
引擎按画布自动缩放全部间距，写 deck.json 时不需要关心这个差异。

## 主题资产与署名规范

每个 `theme.json` 都应包含：

- `source_template`：原始模板文件名
- `provenance`：模板来源，例如“暨南大学官方模板”
- `photo_credit`：背景照片/插画的署名状态；无法从原模板确认时写“沿用原模板，未提供独立摄影署名”，不要猜作者
- `meta_format`：封面汇报人、学院、导师和日期的格式

封面默认使用三行格式：

```text
汇报人：张同学｜信息科学技术学院 · 计算机科学与技术
指导教师：李老师 教授
2026 年 5 月
```

用户提供了姓名、职称或单位时原样保留，不擅自补“教授”“博士”等头衔。允许在 `meta.meta_lines` 中提供字符串数组完全覆盖默认格式。

### 预览图

- `preview/cover.png` 和 `preview/content.png` 使用中性示例名，不保留模板原作者、制作人或真实学生姓名。
- 五套主题的预览使用同一标题、同一汇报人、同一学院和同一日期，便于横向比较。
- 预览图只展示主题，不充当事实案例；示例数据必须标明为演示数据。
- 更换模板背景图时保持比例，不拉伸校徽，不裁掉校名和必要标识。
- 重新生成方法：用 `examples/组会示例.deck.json` 构建该主题，再用 `render_preview.ps1 -W 1280 -H 720` 导出。第 1 页存为 `cover.png`，标题为「修正增强后，三个种子都稳定恢复」的图表页存为 `content.png`。改动引擎的绘制逻辑或主题几何后，五套预览要一起更新。

## 切换主题

```bash
python <SKILL_DIR>/scripts/build_deck.py .jnuppt/xxx.deck.json --theme jnu-rigor --out "答辩.pptx"
```

语义内容通常可以复用。用户说「换成红色那套」「答辩那套」就重跑一次；换完仍要检查标题换行、导航容量、图片裁切和对比度，必要时缩短措辞。
换完**要重新体检**——深色主题上的浅色文字、大画布上的字号都可能变化。

## 字体风险

主题包已经把原模板内嵌的字体剥掉了（`jnu-rigor` 85MB、`jnu-red` 67MB，
每份组会 PPT 都背这个体积不现实）。代价是字形会退到系统字体：

| 主题 | 原模板字体 | 实际会用 |
|---|---|---|
| `jnu-teal` / `jnu-defense` | 微软雅黑 | 微软雅黑（中文 Windows/WPS 必有，零风险） |
| `jnu-crisp` | 微软雅黑 + 华文行楷（封面艺术字） | 同左（Windows 自带；**Mac 上行楷会退成楷体**） |
| `jnu-rigor` | 思源黑体 1 + Akzidenz-Grotesk | Source Han Sans CN → 微软雅黑 |
| `jnu-red` | 可画寒风体-简 + 思源宋体 | 思源宋体 → 宋体 |

**要去陌生电脑上讲**（答辩机房、比赛现场）时提醒用户两条保险做法：
1. PowerPoint：文件 → 选项 → 保存 → 勾「将字体嵌入文件」
2. 或者导出一份 PDF 备用

## 各主题的注意事项

- **jnu-defense**：目录页/章节页/结尾页是深青满底，文字必须浅色；内容页是白底。
  主题的 `defaults` 已配好，别手动改成深色。
- **jnu-rigor**：目录页和章节页背后有校门水印，文字别压在水印最浓的地方；
  支持章节英文副行（`chapters` 里写 `title_en`）。
- **jnu-crisp**：内容页背景是校园实景，信息密集的页会显吵；
  正文多的页建议改用 `cards` 或 `compare`，让白底卡片把文字托起来。
- **jnu-red**：背景是暖橘渐变 + 校门水印，**正文绝不能用浅色**；
  文字一律铺 `surface` 卡片再写，否则对比度不过关。

---

# 收编一套新模板

同学看到好看的模板想用，不必从头做——跑一遍下面的流程就能变成新主题包。
这是这个 skill 的「泛化」所在：我们做的不是五套硬编码模板，是一套把任意
.pptx 拆成「外壳 + token」的方法。

## 原理

每套主题 = 三个东西：

```
themes/<id>/
├── theme.json     色板 / 字体回退链 / 字号阶梯 / 版面几何 / 导航规格
├── stencil.pptx   「外壳页」——背景、装饰、校徽、导航底板，不含任何正文
└── shells.json    外壳页索引（make_stencils.py 自动生成）
```

出片时每页先克隆一张外壳（保真），再用 python-pptx 把正文画上去（可编辑）。
模板里没有的版式（时间轴、KPI、对比）照样画得出来，因为正文本来就是生成的。

## 步骤

### 1. 解剖原模板

```bash
python scripts/inspect_template.py "新模板.pptx" --all
```

输出画布尺寸、每页形状名与坐标、实际用到的颜色和字体频次、媒体清单。
据此判断：哪几页可以当封面 / 内容底 / 目录 / 结尾的外壳，哪些形状是装饰、哪些是正文。

### 2. 写外壳提炼规则

在 `scripts/stencil_spec.json` 里加一节：

```jsonc
"my-theme": {
  "source": "新模板.pptx",
  "shells": [
    {"id": "cover", "slide": 1, "keep": ["底图", "logo"]},        // 白名单
    {"id": "base",  "slide": 4, "drop": ["正文框1", "正文框2"]},   // 或黑名单
    {"id": "toc",   "slide": 2, "keep": ["..."],
     "drop_nested": ["组合里的某个文本框"],                        // 穿透组合删除
     "clear_text":  ["保留形状但清空文字的组合"]}
  ]
}
```

原则：**背景、装饰、校徽、导航底板留下；一切正文和章节文字删掉**——
因为导航标签要按用户的章节重画，正文更是全新生成的。

### 3. 生成并瘦身

```bash
python scripts/make_stencils.py --templates "<模板文件夹>" --theme my-theme
python scripts/slim_pptx.py themes/my-theme/stencil.pptx --inplace
```

`make_stencils.py` 会给外壳里每个形状加 `shell_` 前缀——体检靠它区分
「模板自带的装饰」和「我们生成的内容」，否则满幅出血的底图会被当成排版事故一路报错。

`slim_pptx.py` 删掉不可达的孤儿 part、剥离内嵌字体、重采样大图。

### 4. 写 theme.json

照着 `themes/jnu-teal/theme.json` 改。必填：

- `palette` —— 从解剖报告的高频色里挑，至少给全 13 个 token
- `fonts` —— 每个角色一条回退链，**首选必须是目标机器上大概率存在的字体**
- `type_scale` —— 20 寸画布的字号要比 13.33 寸大约 1.5 倍
- `geometry` —— 各区域的绝对英寸坐标（`canvas` 必须和模板实际尺寸一致）
- `nav` —— `rail-left` / `tabs-top` / `bar-top` / `header-label` / `none` 五选一
- `defaults` —— 各类页的默认文字色。**深色外壳上一定要配浅色文字**，
  这是最容易翻车的地方
- `provenance` / `photo_credit` —— 模板和背景资产来源；不知道就明确写未知，不猜署名
- `meta_format` —— 汇报人、学院、导师和日期的统一标签、分隔符和顺序

### 5. 验收

```bash
python scripts/build_deck.py examples/组会示例.deck.json --theme my-theme --out .jnuppt/t.pptx
python scripts/check_deck.py .jnuppt/t.pptx --theme my-theme --deep
powershell -File scripts/render_preview.ps1 -Path .jnuppt/t.pptx -Out .jnuppt/prev
```

示例 deck 覆盖了全部 15 种版式，体检 0 ERROR + 预览图肉眼过一遍，就算收编成功。

### 常见坑

- **克隆页不能带 `notesSlide` 关系**：备注页会反向指回原页，原页一删就成了
  循环引用，PowerPoint 直接拒绝打开文件。`ppt_kit.clone_slide` 已经处理。
- **形状名字会骗人**：见过模板里把青色竖条命名成「大白条」。以渲染结果为准。
- **组合内部的坐标不是页面坐标**：group 有自己的 chOff/chExt 变换，
  `inspect_template.py` 打出来的子形状坐标只能参考，不能直接当页面位置用。
- **模板自带的背景往往是低分辨率大图**：解剖时看一眼 dpi，太糊的话考虑换底图。
