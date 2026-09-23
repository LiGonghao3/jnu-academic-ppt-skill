# Font Setup / 字体配置

[中文](#中文) · [English](#english)

## 中文

官方 Word 模板主要依赖宋体、黑体和 Times New Roman。LaTeX 要尽量接近 Word，关键是中文字体。

### 推荐方案

| 场景 | 文档类选项 | 说明 |
| --- | --- | --- |
| Windows 本地 | `fontset=windows` | 使用系统 SimSun/SimHei |
| TeXPage / Overleaf / macOS / Linux | `fontset=bundled` | 从本目录加载字体文件，最适合最终稿 |
| macOS 草稿 | `fontset=mac` | 使用 Songti SC/Heiti SC，外观接近 |
| 零配置草稿 | `fontset=auto` 或 `fontset=fandol` | 优先保证可编译，字体不保证完全一致 |

### 使用 bundled 字体

从 Windows 的 `C:\Windows\Fonts\` 复制以下文件到本目录：

```text
fonts/
├── README.md
├── simsun.ttc
├── simhei.ttf
├── simkai.ttf
└── simfang.ttf
```

然后在主文件中使用：

```latex
\documentclass[numbering=arabic,fontset=bundled]{jnuthesis}
```

### 检查字体

```bash
fc-list :lang=zh family | sort -u
fc-match "Songti SC"
fc-match "Noto Serif CJK SC"
```

注意：Word 中的“宋体加粗”通常是宋体伪粗，不等于黑体。模板已经为宋体配置 `AutoFakeBold`，正文里的 `\textbf{...}` 会尽量模拟 Word 的宋体加粗效果。

## English

The official Word template mainly relies on SimSun, SimHei, and Times New Roman. For the closest LaTeX output, Chinese font selection matters most.

### Recommended Settings

| Scenario | Class option | Notes |
| --- | --- | --- |
| Windows local build | `fontset=windows` | Uses system SimSun/SimHei |
| TeXPage / Overleaf / macOS / Linux | `fontset=bundled` | Loads font files from this directory; best for final drafts |
| macOS draft | `fontset=mac` | Uses Songti SC/Heiti SC |
| Zero-config draft | `fontset=auto` or `fontset=fandol` | Prioritizes successful builds, not exact Word matching |

### Bundled Fonts

Copy these files from `C:\Windows\Fonts\` into this directory:

```text
fonts/
├── README.md
├── simsun.ttc
├── simhei.ttf
├── simkai.ttf
└── simfang.ttf
```

Then use:

```latex
\documentclass[numbering=arabic,fontset=bundled]{jnuthesis}
```

Word's bold SimSun is usually faux-bold SimSun, not SimHei. This class uses `AutoFakeBold` for the body CJK font to better match Word behavior.
