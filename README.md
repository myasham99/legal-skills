# legal-skills

> 面向法律实务的 AI Agent Skills 集合（中文优先）。
> 覆盖法律文书起草、债务计算、案件材料命名归档、履约判决跟踪、引注核验、账户解冻核对等场景。

## 包含的 skills

| skill | 用途 | 主要依赖 |
|---|---|---|
| [`answer-statement`](skills/answer-statement/) | 民事答辩状 / 答辩意见起草与复核（首部字段、标题层级、落款规范） | — |
| [`debt-calc`](skills/debt-calc/) | 从判决书提取债务要素，按 LPR 分段算息，生成公式联动 Excel（含加倍迟延利息、履约保证金） | `openpyxl`、`PyMuPDF`（可选） |
| [`naming-n-filing`](skills/naming-n-filing/) | 案件材料智能命名与归档（提取案号 / 相对方 / 项目 → 规范文件名 → 归入目录） | `pypdf`、`PyMuPDF`、`Pillow`、`imagehash`、Tesseract OCR（可选） |
| [`performance-of-obligations`](skills/performance-of-obligations/) | 履约判决跟踪台账全流程（底稿 → 多张派生台账、付款核对、交叉核对） | `openpyxl` |
| [`legal-citation-redline`](skills/legal-citation-redline/) | 法律引注核验，并**在原有 Word 文件上加批注完成就地修订**（改编自上游开源项目，见 NOTICE） | — |
| [`ccp-legal-meeting`](skills/ccp-legal-meeting/) | 批量生成同版式 Word / PPT 的技术骨架，含照片按屏幕内容归档流程与排版踩坑清单 | `python-docx`、`python-pptx` |
| [`unfreeze-account-assets`](skills/unfreeze-account-assets/) | 账户解冻情况核对与台账 P/Q 列校准（以财务实时台账为基准） | `openpyxl` |

## 安装

把需要的 skill 目录复制到你的 AI 客户端 skills 目录，重启客户端即可被识别：

| 客户端 | 目标路径 |
|---|---|
| Claude Code | `~/.claude/skills/<skill-name>/` |
| WorkBuddy | `%USERPROFILE%\.workbuddy\skills\<skill-name>\` |
| 其他（支持 SKILL.md 约定） | 放入对应 skills 目录 |

```
~/.claude/skills/
├── answer-statement/
├── debt-calc/
├── naming-n-filing/
└── ...
```

## 使用前请适配

本仓库内容已做通用化处理，直接使用前请按自身环境调整：

- **路径**：文中出现的 `D:/【用户名】`、`D:/【姓名】` 等均为占位示例，请替换为你的实际目录；脚本中的路径常量同理。
- **数据表**：`project-codes.md`、`project-mapping.md` 等仅保留**结构与虚构示例**，请用本单位台账导出结果替换。
- **业务术语**：部分术语按通用口径改写（如「专项借款台账」「履约判决跟踪底稿」「对上级月度报送」），请对照本单位口径理解。
- **法院代码**：`court-codes.md` 为公开的法院案号代码对照，可直接使用。

## 合规说明

- 本仓库**不含**任何真实当事人姓名、案号、单位名称、项目清单或内部台账数据。
- `naming-n-filing`、`debt-calc` 等含第三方组件依赖，均通过标准包管理器安装，未捆绑第三方二进制。
- 第三方来源与许可声明见 [NOTICE.md](NOTICE.md)。

## 免责声明

本项目提供的是**通用方法与工作流骨架**，不构成法律意见。使用者应对所生成文件的准确性、合规性自行核验。

## 许可

[MIT](LICENSE)
