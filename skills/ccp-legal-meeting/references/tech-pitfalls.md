# 技术坑清单（写代码前必读）

## Word COM（.doc→.docx、docx→PDF；soffice.py 在 Windows 不可用，报 AF_UNIX）

- 本机 Office 2019：`C:\Program Files\Microsoft Office\Root\Office16\WINWORD.EXE`，用 PowerShell COM（`New-Object -ComObject Word.Application`），需跳出沙箱执行（dangerouslyDisableSandbox，用户会收到授权确认）。
- **PS1 脚本必须 ASCII-only**：中文路径经文件清单传入——Python 写 UTF-8-BOM（`utf-8-sig`）清单，PS1 用 `[System.IO.File]::ReadAllLines($path, [System.Text.Encoding]::UTF8)` 读取。直接在 ps1 里写中文路径会因编码错乱报"Word 无法读取此文档，文档可能已损坏"（假错）。
- SaveAs 的 `[ref]` 参数目标必须 `[string]` 强转：`$dst = [string](Join-Path ...)`，否则报 psobject 绑定失败。
- 格式代码：docx=16，pdf=17。打开用只读 `Open($src, $false, $true)`，转完 `Close($false)`，最后 `$word.Quit()`；结果写 log 文件再用 Python 读（UTF-16）。
- log 文件若 Read 工具报二进制，用 Python `open(..., encoding='utf-16')` 读。

## python-docx

- **w:ind 必须用 `pPr.get_or_add_ind()`**，手动 `OxmlElement('w:ind')` append 到 pPr 末尾会破坏 OOXML 子元素顺序（ind 须在 jc 前），Word 打开报"文档损坏"。
- 首行缩进 2 字符：`w:firstLineChars="200"` + `w:firstLine="600"`。
- 页脚页码：footer 段落里放 begin/instrText(PAGE)/end 三个 run（OxmlElement 构造）。
- 替换模板段落文字：`runs[0].text = new; 其余 runs[i].text=''`（保第一个 run 格式）。
- **幂等替换不能先全局查"new 是否已存在"**：会被其他段落误判（如"2026年9月11日"已在开会时间行→落款"2021年11月9日"漏改）。逐段落匹配旧 key 才安全。
- 合并单元格：`row.cells` 按网格列返回会重复/错位，先按 `id(c._element)` 去重得 unique 列表，再按内容（如"地点"）+ 相邻关系定位目标格。
- 填表格单元格前先判 `if not cell.paragraphs[0].text.strip()` 再 add_run，防重复运行时叠加。

## python-pptx（PPT背景改字、滚动资料PPT）

- 保格式替换：段落内只改 `runs[0].text` 清空其余 runs。
- **新行数 > 现有段落数时，必须 deepcopy 克隆最后一个有文字的段落元素**（`tf.paragraphs[-1]._element.addnext(new_el)`）再填文字；直接把多行用全角空格拼进一段观感差，add_paragraph() 会丢格式。
- 匹配原文用去空白归一化（`''.join(s.split())`），模板里常有多余空格/换行。
- 滚动资料PPT：16:9（12192000×6858000 EMU）；学习内容页 = word/pdf 整页转图片后插入，双页横排约 10.7×7.5 英寸占满版心；封面/标题页复用背景版式（背景图 + 标题文本框）。

## PDF 抽验（pypdf）

- **方正小标宋简体字形在 pypdf 中提取不到**——标题"看似缺失"实为正常，勿误判；仿宋/黑体/宋体可正常提取。
- 扫描件 PDF 无文字层（pypdf 返回空），如签到表扫描件、入党申请书，无法提取要素，只能让用户对照纸质件补填。
- Word 自动在西文与中文间加空隙，提取文本带空格（"2026 年 9 月 11 日"），匹配时去空白再比。
