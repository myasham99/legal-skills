# docx-official-document

> 中文公文与诉讼文书的 Word 排版规范（含可运行的 python-docx 实现示例）。
> 一套可直接落地的 `.docx` 格式约定，解决"字体老是跑掉、导航窗格空白、缩进在各阅读器不一致"等常见问题。

## 这套规范解决什么问题

中文 Word 文档的格式"看起来差不多"，但对齐到像素级别时经常出问题，原因通常是：

| 问题 | 根因 |
|---|---|
| 明明设了仿宋/黑体，Word 里却显示宋体或 MS Gothic | 样式里残留 **theme 主题字体**（`asciiTheme`/`eastAsiaTheme`），Word 优先按 theme 解析 |
| 标题在导航窗格里不出现、目录生成不出来 | 标题用的是"加粗＋大字号"伪装，而非 Word **内置 heading 样式** |
| 缩进量在不同阅读器/版本里不一致 | 只依赖 `docDefaults` 的字符缩进，未在段落上显式写 `w:ind` |
| 生成的 docx 报 XSD 错误 | OOXML 子元素顺序错误（如 `w:ind` 排到了 `w:jc` 之后） |

本仓库把这些问题连同修复方法一起写进规范，并给出示例脚本。

## 内容结构

```
docx-official-document/
├── SKILL.md                       # 技能定义（含核心约束与踩坑速查表）
├── README.md                      # 本文件
├── LICENSE                        # MIT
├── references/
│   ├── 公文格式规范.md             # 公文类：页面 / 字体层级 / 表格 / 实现要点
│   └── 诉讼文书格式规范.md         # 诉讼文书：答辩状 / 质证意见 / 证据清单骨架 + 实现陷阱
└── scripts/
    └── example_generate.py        # 可运行示例（python-docx）
```

## 快速开始

```bash
pip install python-docx
python scripts/example_generate.py 示例.docx
```

生成的文档包含：四级标题层级、固定行距 22 磅、首行缩进 2 字符、页脚居中页码、可在导航窗格展开的标题树。

## 作为 AI 技能使用

把整个文件夹放入 AI 客户端的 skills 目录：

- Claude Code：`~/.claude/skills/docx-official-document/`
- WorkBuddy：`%USERPROFILE%\.workbuddy\skills\docx-official-document\`

之后让 AI「按规范起草一份工作证明 / 民事答辩状」，它会自动读取 `references/` 下的规范执行。

## 核心约定（速览）

- **页面**：A4，上下 2.54 cm、左右 3.17 cm
- **西文数字**：一律 Times New Roman
- **中文字体**：题目 `方正小标宋简体`；一级标题 `黑体`；二级 `楷体`；三级与正文 `仿宋_GB2312`
- **正文**：首行缩进 2 字符、固定行距 22 磅、两端对齐
- **标题**：使用 Word 内置 `heading 1`–`heading 4`，保证导航与目录可用
- **页码**：≥ 2 页时页脚居中 PAGE 域
- **编号**：`一、`/`（一）`/`1.` 手打，不用自动编号

## 说明

- 本仓库内容为**通用排版规范与实现方法**，示例中的当事人、案号、单位名称均已占位化（如 `（2026）粤XXXX民初XXXX号`）。
- 规范基于对真实 Word 文档 OOXML（`styles.xml` / `document.xml` / `footer*.xml`）的解析整理，适用于 Microsoft Word 与 WPS。

## 许可

[MIT](LICENSE)
