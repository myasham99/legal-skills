---
name: legal-citation-redline
description: 全面法学引注诊断、补全与格式化工具（本地手册版本：《法学引注手册》2019年11月版，规则1—100）。用于检查混乱脚注、识别缺失要素、说明应去哪里查找缺失信息，并在用户提供来源文件、法条、案例或网页信息时生成符合《法学引注手册》的中文/外文法律脚注。适用于法学引注、脚注检查、引注修正、引用格式转换、法律文献引用、法条脚注、案例脚注、论文引注。
---

# Legal Citation Redline（引注核验 · 就地修订）

> **来源与许可**：本 skill 改编自开源项目 [NEU-ZHA/legal-ai-skills](https://github.com/NEU-ZHA/legal-ai-skills) 中的 `legal-citation-comprehensive`（MIT License, Copyright (c) 2026 Legal AI Skills contributors）。
> **本地改动**：核验结论**不再输出为新的 Word 文档**，而是直接写入**原 Word 文件的批注**（就地修订），并弃用配套的 `legal-citation-automator` 出稿通道。
> 原项目与本项目均采用 MIT 许可；第三方版权资料（如《法学引注手册》PDF）不随本仓库分发。



## 参考资料自备说明（重要）

本仓库**不包含**《法学引注手册》等第三方版权资料，也不包含由其生成的 OCR 全文或完整规则索引（遵循上游项目同一原则）。

如需完整的规则数据层，请：
1. 合法取得该手册（纸质或正版电子版）；
2. 放入 `assets/` 目录（文件名见 `references/README_REFERENCE_DATA.md`）；
3. 运行 `python scripts/build_reference_data.py` 在**本地**生成 `handbook_raw.md`、`handbook_rule_index.*`、`citation_rules.json`。

未生成上述文件时，本 skill 仅依赖 `references/` 中随仓库分发的方法说明文件工作。

## Role

This skill is the citation brain. It diagnoses and generates citation text. It does not edit Word files directly; use `legal-citation-automator` for DOCX insertion.

Primary rule: never invent missing bibliographic facts. If a required element is absent and cannot be extracted from a provided source, mark it as `[待补: 要素名]` and tell the user where to find it.

Completeness rule: the simplified templates are not the whole handbook. Before giving a final answer for any non-trivial citation, consult a legally obtained copy of the citation handbook or a user-provided rule index. This public repository does not redistribute third-party handbook PDFs or OCR dumps. Use 本地生成的 `citation_rules.json` only as a fast execution layer when the user or local installation provides it.

Weak-agent guardrail: if you are unsure, do not improvise. Follow `references/operator_guardrails.md` exactly. A citation answer is not complete unless it states the source type, missing elements, lookup path, suggested format, and relevant handbook rule numbers.

## 诉讼文书模式（答辩状 / 起诉状 / 质证意见 · 强制流程）

当核验对象是**诉讼文书**（答辩状、起诉状、质证意见、代理词等）时，不适用脚注体例，改走
[`references/诉讼文书引注核验流程.md`](references/诉讼文书引注核验流程.md)：**取官方原文 → 出《引注核验表》→ 以同一序号把核验结果写进文书批注**。

要点（用户 2026-09-11 确认）：
- 核验表格式：标题「××状引注核验表」＋**逐行**表头信息（文书／案号／受诉法院／案由／核验日期／核验方式）＋ 4 列表格（`序号 | 答辩状引注 | 类型 | 核验来源与结论`）；**表尾不再写核验结论、核验说明、复核依据等段落**；全文黑色。
- **序号联动**：核验表的行号＝文书批注中的序号，批注写作 `【引注核验 #N】…（法条原文）… 核验来源…；结论…`。
- 批注作者统一 **`reviewer`**；批注不限标题，正文各引注处、庭审要点、待补证据清单、提交提示等都应写批注，正文从简。
- 引注形式遵《法学引注手册》2019年版：**行文援引**用阿拉伯数字并省略括号（规则61）；**引号内原文引用不改写**；司法解释标文件号（规则64）、首次出现处说明简称（规则59）；判例写案例名称＋审判法院全名＋案号＋文书类型（规则70、71）。

## Quick Workflow

1. Classify the citation type.
   - Search the user-provided handbook/rule index for the source category and rule numbers.
   - Use 本地生成的 `citation_rules.json` for required elements, templates, and handbook anchors when that optional local reference file is available.
   - If the repository lacks full reference data, stop at diagnosis/placeholders and tell the user what source material is needed.
2. Extract existing elements from the messy footnote or supplied source.
3. Diagnose missing required and optional elements.
4. Give source lookup guidance from `references/source_lookup_guide.md`.
5. Generate a standard citation.
   - If complete: output the final footnote.
   - If incomplete: output a placeholder version and a precise checklist.
6. For Chinese Civil Law homework, apply the course-specific statute rule:
   - If正文 only names a statute article, the footnote should include the article text.
   - If正文 already quotes the full article and the footnote adds nothing, omit the footnote.
7. Run `python scripts/self_test.py` after changing this skill or before distributing it. Full handbook coverage tests require locally supplied reference data.

## Deterministic Helper

For a first-pass report, run:

```bash
python scripts/citation_diagnose.py --text "王名扬《美国行政法》第18页"
```

Or:

```bash
python scripts/citation_diagnose.py footnotes.txt --json
```

The script provides classification, present/missing elements, source lookup hints, and a template-based corrected citation when local rule data is available. Treat script output as a first pass; verify edge cases against the handbook.

## Standard Answer Shape

When responding to a user, use this shape unless they ask for something else:

```text
识别结果：
- 类型：
- 置信度：

已有要素：
- ...

缺失/需确认要素：
- ...

去哪里找：
- ...

建议格式：
...

注意：
- ...
```

## Source Extraction Rules

- PDF book: inspect cover, copyright page, title page, table of contents area. Use copyright page for publisher, year, edition.
- PDF article: inspect first page and page headers/footers. Use article first page for author, title, journal, year, issue, start page.
- Statute or judicial interpretation: prefer official databases or PKULaw MCP when available. Capture full official name, article number, paragraph/item, current validity, and article text if needed.
- Case: capture case name, court, docket number, document type, source series/database. Public bulletin, guiding case, and People’s Court Case Library entries use their own formats.
- Web source: capture author if available, article title, platform/site, publication date, URL, and access date.

## Important Rules

- Chinese books: `作者：《书名》，出版社年份年版，第X页。`
- Chinese journal articles: `作者：《论文名》，载《期刊名》年份年第X期，第X页。`
- Edited collections: `作者：《文章名》，载编者主编：《书名》，出版社年份年版，第X页。`
- Statutes: `《法律名称》第X条第Y款第Z项。` Use Arabic numerals.
- Normative documents: include document number when available, e.g. `《文件名称》（发文字号）。`
- Cases: include case name, court, docket number, and document type unless citing a special source such as a bulletin or guiding case.
- Chinese repeat citations use `同前注〔X〕，第Y页。` or a clear short form. Do not use `前引`, `同上`, `supra`, or `Ibid.` for Chinese sources.
- Direct quotation: no `参见`; indirect or conceptual borrowing: use `参见`; secondary citation: use `转引自`.
- Statute quotation vs paraphrase: if a footnote says `第X条规定，...`, first decide whether the following words are the statute's original text. If yes, use `第X条规定：“……”`; if no, rewrite as a paraphrase such as `依第X条，可以说明……` or `参见第X条` and do not make the paraphrase look like quoted law text.
- Multiple sources supporting the same sentence or proposition belong in one footnote, separated by semicolons. Do not create two adjacent footnotes at the same sentence-final position merely because there are two sources.
- Use separate footnotes only when the sentence has distinct citation anchors, such as one statute supporting the first clause and another statute supporting a later clause. In that case, place each footnote immediately after the specific article, claim, or clause it supports.
- Explanatory notes are allowed when they explain the cited source's relevance, but do not hide major legal reasoning in a citation footnote. If the explanation is doing argumentative work, move it into the body and leave the footnote mostly as source support.

## Statute Original Text vs Paraphrase

Legal article footnotes often fail when an agent writes a paraphrase as if it were the statute's exact words. Apply this gate before finalizing any statute footnote:

- Exact or near-exact statutory text: use a colon and Chinese quotation marks.
  - Shape: `《上海市城市管理综合行政执法条例》第11条第2款规定：“……”`
- Paraphrase or synthesis: do not use `规定，...` followed by an unquoted sentence. Use `依/根据/参见` and make clear that the writer is summarizing.
  - Shape: `参见《上海市城市管理综合行政执法条例》第11条第2款。依该款，街道办事处具体实施哪些行政执法事项，应以市人民政府确定并公布的事项为准。`
- If the agent cannot verify the wording against an official source, do not add quotation marks around invented text. Mark `[待补: 条文原文]` or keep the note as a clearly labeled paraphrase.

## Footnote Bundling and Placement

When formatting homework or article citations, decide the citation point before writing the footnote:

- Same proposition, several sources: one footnote, sources separated by `；`.
  - Example shape: `参见作者甲：《文章名》，载《期刊名》2023年第2期，第276页；P. G. Biddle, Tree Root Damage to Buildings, ... p.131.`
- Different propositions inside one sentence: split the sentence-level anchors and place each footnote after the clause or legal article it supports.
  - Example shape: `通过《民法典》第580条〔脚注〕处理继续履行排除，通过第583条〔脚注〕处理损害赔偿。`
- Never output two bracketed footnotes or two DOCX footnote references back-to-back after the same punctuation unless the user explicitly asks to preserve an existing defective draft for comparison.

## Full Handbook Coverage

The rule index is mandatory for careful work:

- Rules 1-23: 引注的一般规范（引注基本要求、排版与符号位置、夹注、标点、重复引用、论文部件）。
- Rules 24-46: 引用纸质出版文献（作者、文献名称、标题、版本、出版信息、页码、章节）。
- Rules 47-51: 引用网络、电视、音像文献。
- Rules 52-57: 引用未发表文献（访谈、私人通讯、内部资料、会议论文、学位论文、档案文献）。
- Rules 58-69: 引用法律文件（名称、缩写、版本、**条款序数**、条文排版、法律法规规章、**规范性文件**、国家标准、立法说明、会议决议、外国法与国际公约、台湾地区法律文件）。
- Rules 70-73: 引用司法案例（案例名称、案件文号、案例来源、裁判时间）。
- Rules 74-76: 引用统计数据。
- Rules 77-100: 外文引注体例（一般要求、英文 78-85、法文 86-90、德文 91-95、日文 96-100）。

For any source type not explicitly represented in `citation_rules.json`, search the user-provided handbook or rule index by rule number or source name. If the OCR text is unclear, verify against the legally obtained original document.

## Reference Navigation

- `references/README_REFERENCE_DATA.md`: explains which optional reference files are intentionally not redistributed.
- `本地生成的 handbook_rule_index.json`: optional local full machine-readable index of rules 1-100. Load/search this for exhaustive coverage if the user provides it.
- `本地生成的 handbook_rule_index.md`: optional readable version of the full index with raw line ranges.
- 本地生成的 `citation_rules.json`: optional machine-readable rule source. Load this first if present.
- `references/missing_elements_matrix.md`: quick checklist for missing elements by type.
- `references/source_lookup_guide.md`: where to find missing bibliographic facts.
- `references/operator_guardrails.md`: strict operating rules for other AI agents.
- `references/common_errors.md`: common mistakes and course-specific footnote traps.
- `本地生成的 citation_handbook_structured.md`: structured digest of the handbook.
- `assets/Law_Journal_Citation_Handbook_2019.pdf`: 本地手册原件（《法学引注手册》2019年11月版，规则1—100），必要时用于最终核对。

## Reference Data Build

When a user has their own legally obtained copy of the handbook PDF, prefer the deterministic local builder before asking an AI to improvise the reference files:

```bash
python3 scripts/build_reference_data.py
```

The builder checks whether the PDF is searchable, extracts local Markdown, builds the 1-100 rule index, writes `citation_rules.json`, and then runs validation. 本技能的本地手册为 **《法学引注手册》2019年11月版（规则1—100）**，全部规则条号均以该版为准。If it refuses to overwrite existing files, rerun with `--force` only after the user confirms they want to rebuild local reference data. If extraction is incomplete, run:

```bash
python3 scripts/build_reference_data.py --print-ai-contract
```

Then use the printed contract to repair the local files.

## Reference Data Audit

When a user rebuilds handbook reference data from their own PDF/OCR, do not trust the generated files merely because they exist. If the builder was not used, run:

```bash
python3 scripts/audit_reference_data.py
python3 scripts/self_test.py
```

`audit_reference_data.py` checks required files, rule count, continuous rule numbers, category ranges, raw line ranges, common rule hints, OCR/mojibake markers, and `citation_rules.json` shape. If it reports `ERRORS`, stop and repair the reference data before claiming full handbook coverage.

## Relationship to Other Skills

- `legal-citation-automator`: use after this skill has produced or approved footnote text; it handles DOCX insertion and compatibility checks.
- `legal-homework-formatter`: use for overall Civil Law homework layout, footnote style, and Word/WPS formatting.
- `pkulaw-legal-search`: use when missing law or case source facts require legal database retrieval.
