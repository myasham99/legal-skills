# Notice

## 一、第三方来源

本仓库 `skills/legal-citation-redline` **改编自**开源项目：

- 项目：[NEU-ZHA/legal-ai-skills](https://github.com/NEU-ZHA/legal-ai-skills)
- 原 skill：`skills/legal-citation-comprehensive`
- 许可：MIT License, Copyright (c) 2026 Legal AI Skills contributors

**本仓库所做的修改**：

1. 核验结论**不再输出为新的 Word 文档**，改为直接写入**原 Word 文件的批注**（就地修订）；
2. 移除对上游配套 skill `legal-citation-automator` 的出稿依赖；
3. 名称由 `legal-citation-comprehensive` 改为 `legal-citation-redline`，以与上游区分；
4. 去除本地环境相关的内容（单位与个人路径、示例案号等）。

上游项目的 NOTICE 亦声明不重新分发私人模板、个人身份数据、API 凭证与第三方版权作品；本仓库遵循同一原则（见下）。

其余 skills 为本仓库作者原创。

## 二、已排除的内容

为遵守许可并保护隐私，本仓库**不包含**：

- 第三方版权作品（如《法学引注手册》等出版物的 PDF、OCR 文本或完整规则索引）
- 任何真实当事人姓名、案号、单位名称、项目清单、内部台账数据
- 个人身份信息、内部组织流程材料（如人员名册、党内会议具体材料）
- API 凭证与任何密钥

`project-codes.md`、`project-mapping.md` 等文件仅保留**数据结构说明与虚构示例行**，请在本地用本单位台账替换后再使用。

## 三、使用提示

- 本项目提供的是通用方法论与工作流骨架，**不构成法律意见**；使用者应自行核验所生成文件的准确性与合规性。
- 涉及具体案件、台账与主体的业务数据，请在本单位合规框架内处理，勿将内部数据提交至公开仓库。
