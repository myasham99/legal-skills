---
name: docx-official-document
description: 按中文字档排版规范生成 Word 文档（.docx）。适用于中文公文（工作证明、业绩证明、结案报告、法律意见书、通知、函、培训材料）与诉讼文书（民事答辩状、起诉状、证据清单、质证意见）。内置页面设置、四级标题字体层级（方正小标宋简体／黑体／楷体／仿宋_GB2312）、西文与数字统一 Times New Roman、固定行距与首行缩进、页脚居中页码、导航窗格标题等完整规范，并附 python-docx 实现示例与 OOXML 常见踩坑修复方法。当用户要求"按规范排版""格式要与既有文书一致""公文格式"或需要生成上述文档时使用。
---

# 中文公文 / 诉讼文书 Word 排版规范（docx-official-document）

本技能包提供**两套可直接落地的中文文档排版规范**与实现方法，用于生成格式统一、可直接交付的 `.docx`。

## 一、两套规范

| 规范 | 适用文档 | 文件 |
|---|---|---|
| 公文类通用规范 | 工作证明、业绩证明、结案报告、法律意见书、通知、函、培训材料 | [`references/公文格式规范.md`](references/公文格式规范.md) |
| 诉讼文书规范 | 民事答辩状、起诉状、证据清单、质证意见 | [`references/诉讼文书格式规范.md`](references/诉讼文书格式规范.md) |

两套规范的**共同底座**一致：A4 纸张、页脚居中页码、西文与数字一律 Times New Roman、标题使用 Word 内置 `heading` 样式（保证导航窗格与目录可用）。差异主要在标题字体与文书骨架。

## 二、为什么需要它

Word 文档"看起来差不多"却总是对不齐，通常源于三类问题，本规范都给出明确做法：

1. **字体有三层来源**：样式（`docDefaults` / `Normal` / `Heading`）、run 直接格式、以及 **theme 主题字体**。只改前两层时，Word 可能按 theme 回退成 MS Gothic 或宋体——必须显式删除 `*Theme` 字体属性。
2. **导航窗格 / 目录只认内置标题样式**：用"加粗＋大字号"伪装的标题不会出现在导航窗格，也无法生成目录。
3. **缩进、行距、XML 顺序各有陷阱**：`w:ind` 必须在 `w:jc` 之前、`pPrDefault` 不能重复 append、固定行距与多倍行距的单位不同等。

## 三、快速上手

[`scripts/example_generate.py`](scripts/example_generate.py) 是一份可直接运行的 python-docx 示例，生成带四级标题、固定行距、首行缩进、页脚页码与可导航标题的合规文档：

```bash
pip install python-docx
python scripts/example_generate.py 输出.docx
```

## 四、核心约束（两套规范通用）

- **页面**：A4（21 × 29.7 cm），页边距上下 2.54 cm、左右 3.17 cm。
- **西文与数字**：一律 `Times New Roman`（每个 run 的 `ascii` / `hAnsi`）。
- **中文字体按层级**：题目／文书名 → `方正小标宋简体`；一级标题 → `黑体`；二级标题 → `楷体`；三级标题与正文 → `仿宋_GB2312`。
- **导航可用**：标题必须使用 Word **内置** `heading 1`–`heading 4` 样式（保留 `outlineLvl`），不要自定义样式或伪标题。
- **页码**：成稿预计 ≥ 2 页时，页脚加居中 PAGE 域。
- **编号**：`一、`/`（一）`/`1.` 手打，不用自动编号。
- **表格内文字**：一般随正文（公文类可用小四）；诉讼文书表格随正文小三。

## 五、常见踩坑（生成后务必自检）

| 现象 | 原因 | 修复 |
|---|---|---|
| 标题显示成宋体 / MS Gothic | 样式残留 `asciiTheme` / `eastAsiaTheme` 等主题字体属性 | 删除 `rFonts` 上所有以 `Theme` 结尾的属性，只留显式字体 |
| 标题 4 变成"加粗斜体" | Word 内置 `Heading 4` 自带 `w:i` | 显式写 `<w:i w:val="0"/>` 与 `<w:iCs w:val="0"/>` |
| 缩进在不同阅读器不一致 | 只依赖 `docDefaults` 的 `firstLineChars` | 每段显式写 `w:ind w:firstLine="600"`（标题段 602，顶格段 0） |
| XSD 报 "ind not expected" | `w:ind` 被排到 `w:jc` 之后 | `jc.addprevious(ind)`，保证 `spacing → ind → jc` 顺序 |
| XSD 报 "pPr not expected" | 在 `pPrDefault` 里 append 了第二个 `w:pPr` | 复用已有 `pPr`（清空子元素后写入），或插在 `w:rPrDefault` 之后 |
| 导航窗格空白 | 标题未使用内置 heading 样式 | 改用 `Heading 1`–`Heading 4`，或自定义样式设 `outlineLevel` |
| 表格宽度错乱 / 变黑 | 只设了列宽没设单元格宽；或用了 `ShadingType.SOLID` | 表格与单元格同时给 DXA 宽度；底纹用 `CLEAR` |

## 六、安装

把本文件夹整体放入你的 AI 客户端 skills 目录：

| 客户端 | 目标路径 |
|---|---|
| Claude Code | `~/.claude/skills/docx-official-document/` |
| WorkBuddy | `%USERPROFILE%\.workbuddy\skills\docx-official-document\` |
| 其他（支持 SKILL.md 约定） | 放入对应 skills 目录 |

放入后重启客户端，即可在技能列表中看到 `docx-official-document`。

## 七、目录结构

```
docx-official-document/
├── SKILL.md                          ← 本文件
├── README.md                         ← 项目说明
├── LICENSE                           ← MIT
├── references/
│   ├── 公文格式规范.md                ← 公文类完整规范
│   └── 诉讼文书格式规范.md            ← 诉讼文书完整规范
└── scripts/
    └── example_generate.py           ← python-docx 可运行示例
```

## 八、许可

MIT License，详见 [LICENSE](LICENSE)。
