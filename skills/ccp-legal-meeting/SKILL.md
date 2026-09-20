---
name: ccp-legal-meeting
description: 批量生成同版式 Word / PPT 文档的技术骨架（通知、会议记录、签到表、背景板），并含"按屏幕内容给照片归档命名"的实操流程与 Word/PPT 排版踩坑修复清单。当需要一次产出多份版式一致的 docx/pptx（会议类、台账类、报表类），或需要把大量会议/活动照片按截图证据归类改名时使用。
---

# 批量文档生成技术骨架（ccp-legal-meeting）

本 skill 只保留**与具体单位无关的技术能力**，可直接用于任何"多份同版式文档 + 照片归档"场景。

> 原版本包含特定单位的会议流程与人员名册，已从开源版中移除（涉及内部组织信息）。

## 一、能力范围

| 能力 | 说明 | 文件 |
|---|---|---|
| 批量 docx / pptx 生成 | 固定版式下产出多份文档：通知、记录、签到表、背景板；含页脚页码、标题层级、表格样式 | [`scripts/ccp_docx.py`](scripts/ccp_docx.py) |
| 通过 Word 转换 | 某些特殊版式（如复杂艺术字/域）用 python 生成后再经 Word 转换，保证兼容 | [`scripts/convert_via_word.ps1`](scripts/convert_via_word.ps1) |
| 照片归档流程 | 以"屏幕内容"为第一证据对照片归类，按 `会议照片【类型】月份<内容>#序号` 规则改名 | [`references/photo-filing.md`](references/photo-filing.md) |
| 排版踩坑清单 | Word / PPT 批量生成中的字体、行距、页码、域、转换等常见坑与修复 | [`references/tech-pitfalls.md`](references/tech-pitfalls.md) |

## 二、典型用法

```bash
# 生成当月多份同版式文档（示例：按月份循环产出通知 / 记录 / 签到表）
python scripts/ccp_docx.py --month 9 --out D:/输出目录
```

照片归档的三条判据（优先级由高到低）：
1. **屏幕标题**写明会议名称 → 直接归属（最可靠）
2. **画面中的文字**（背景板、横幅、PPT 标题）
3. **人数与座位量级** —— 仅作辅助，不足以推翻前两条

## 三、注意事项

- 批量生成前先确认版式模板（字体、页边距、页脚页码），避免逐个文件返工。
- 生成后必做一次渲染检查，具体见 [`references/tech-pitfalls.md`](references/tech-pitfalls.md)。
