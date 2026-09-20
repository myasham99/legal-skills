#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
债务计算表生成器（通用版）v5.0
============================================
输入：判决书关键信息（案号、项目、本金、付款记录等）
输出：与"××××公司 债务计算表"格式一致的 .xlsx

算法规则（依法、固定）：
  1. 利率：1年期LPR，按公布日自动分段；2025-05-20至今固定3.00%
  2. 天数：含头含尾（终止日 - 起始日 + 1）
  3. 普通利息 = 本金 x 年利率 / 360 x 天数
     加倍迟延利息 = 基数 x 日利率(1.75/10000) x 天数
  4. 冲抵顺序：付款先冲利息，余下冲本金；利息不足时滚存、本金不变
  5. 履约保证金不计普通利息，但计入加倍迟延基数

用法：修改底部 __main__ 中的 params 字典后运行即可。
"""

# ============================================================
# ★ 路径配置（必须放在最前面，任何逻辑执行前就位）
# ============================================================
import sys
import os
import shutil
import re
from pathlib import Path
from datetime import datetime, timedelta, date

# OCR / naming_n_filing 模块路径
OCR_SKILL_DIR = Path("C:/Users/ACME/.claude/skills/ocr")
NAMING_SKILL_DIR = Path("C:/Users/ACME/.workbuddy/skills/naming-n-filing")
sys.path.insert(0, str(OCR_SKILL_DIR))
sys.path.insert(0, str(NAMING_SKILL_DIR))

try:
    from ocr import extract_text as ocr_extract_text
    _OCR_AVAILABLE = True
except ImportError:
    _OCR_AVAILABLE = False

try:
    import naming_n_filing as nf
    _NAMING_AVAILABLE = True
except ImportError:
    _NAMING_AVAILABLE = False

# 归档根目录
D_ROOT = Path("D:/【用户名】")
D_DEBT_CALC = D_ROOT / "【诉讼】" / "【执行履行】" / "债务计算"


def to_full_project_name(name):
    """项目简称 → 项目全称（查 naming-n-filing/references/project_mapping.md）。
    映射表列：| 序号 | code12 | 全称 | 简称 | 利润中心 |；无匹配时原样返回。"""
    if not name:
        return name
    try:
        mp = NAMING_SKILL_DIR / "references" / "project_mapping.md"
        if not mp.exists():
            return name
        name_s = str(name).strip()
        best = None
        for line in mp.read_text(encoding="utf-8").splitlines():
            line = line.strip()
            if not line.startswith("|"):
                continue
            cols = [c.strip() for c in line.split("|")]
            # 表结构：| # | code12 | 项目全称 | 简称 | 利润中心 |
            if len(cols) < 6 or cols[1] in ("", "#"):
                continue
            full, short = cols[3], cols[4]
            if not full or not short:
                continue
            if short == name_s:
                return full
            if (name_s.startswith(short) or short in name_s) and best is None:
                best = full
        return best or name
    except Exception:
        return name

# ============================================================
# 第三方库导入
# ============================================================
import openpyxl
from openpyxl.styles import Font, Border, Side, Alignment, PatternFill
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.properties import PageSetupProperties

# ============================================================
# 1. LPR 历史数据（1年期，%）
# ============================================================
LPR_TABLE = [
    (date(2019, 8, 20), 4.25), (date(2019, 9, 20), 4.20),
    (date(2019,10, 20), 4.20), (date(2019,11, 20), 4.15),
    (date(2019,12, 20), 4.15), (date(2020, 1, 20), 4.15),
    (date(2020, 2, 20), 4.05), (date(2020, 3, 20), 4.05),
    (date(2020, 4, 20), 3.85), (date(2020, 5, 20), 3.85),
    (date(2020, 6, 20), 3.85), (date(2020, 7, 20), 3.85),
    (date(2020, 8, 20), 3.85), (date(2020, 9, 20), 3.85),
    (date(2020,10, 20), 3.85), (date(2020,11, 20), 3.85),
    (date(2020,12, 20), 3.85), (date(2021, 1, 20), 3.85),
    (date(2021, 2, 20), 3.85), (date(2021, 3, 20), 3.85),
    (date(2021, 4, 20), 3.85), (date(2021, 5, 20), 3.85),
    (date(2021, 6, 20), 3.85), (date(2021, 7, 20), 3.85),
    (date(2021, 8, 20), 3.85), (date(2021, 9, 20), 3.85),
    (date(2021,10, 20), 3.85), (date(2021,11, 20), 3.85),
    (date(2021,12, 20), 3.80), (date(2022, 1, 20), 3.70),
    (date(2022, 2, 20), 3.70), (date(2022, 3, 20), 3.70),
    (date(2022, 4, 20), 3.70), (date(2022, 5, 20), 3.70),
    (date(2022, 6, 20), 3.70), (date(2022, 7, 20), 3.70),
    (date(2022, 8, 20), 3.65), (date(2022, 9, 20), 3.65),
    (date(2022,10, 20), 3.65), (date(2022,11, 20), 3.65),
    (date(2022,12, 20), 3.65), (date(2023, 1, 20), 3.65),
    (date(2023, 2, 20), 3.65), (date(2023, 3, 20), 3.65),
    (date(2023, 4, 20), 3.65), (date(2023, 5, 20), 3.65),
    (date(2023, 6, 20), 3.55), (date(2023, 7, 20), 3.55),
    (date(2023, 8, 20), 3.45), (date(2023, 9, 20), 3.45),
    (date(2023,10, 20), 3.45), (date(2023,11, 20), 3.45),
    (date(2023,12, 20), 3.45), (date(2024, 1, 20), 3.45),
    (date(2024, 2, 20), 3.45), (date(2024, 3, 20), 3.45),
    (date(2024, 4, 20), 3.45), (date(2024, 5, 20), 3.45),
    (date(2024, 6, 20), 3.45), (date(2024, 7, 20), 3.35),
    (date(2024, 8, 20), 3.35), (date(2024, 9, 20), 3.35),
    (date(2024,10, 20), 3.10), (date(2024,11, 20), 3.10),
    (date(2024,12, 20), 3.10), (date(2025, 1, 20), 3.10),
    (date(2025, 2, 20), 3.10), (date(2025, 3, 20), 3.10),
    (date(2025, 4, 20), 3.10), (date(2025, 5, 20), 3.00),
    # 2025-05-20 至今固定 3.00%，无需继续添加
]

def get_lpr_rate(dt):
    """获取 dt 当日适用的1年期LPR（百分数，如 3.45 表示 3.45%）"""
    rate = None
    for d, r in LPR_TABLE:
        if dt >= d:
            rate = r
        else:
            break
    return rate

# ============================================================
# 2. 日期工具
# ============================================================
def fmt_date(d):
    """将 date 对象格式化为 'YYYY年MM月DD日' 字符串"""
    return d.strftime("%Y年%m月%d日")

def date_to_serial(d):
    """date -> Excel 序列号（1900日期系统）"""
    return (d - date(1899, 12, 30)).days

# ============================================================
# 3. 利息计算核心
# ============================================================
def calc_interest_segment(principal, start_date, end_date, rate_schedule=None):
    """
    计算 [start_date, end_date] 区间内的利息（算头不算尾：起始日计入，截止日不计入）
    按LPR（或自定义利率表）自动分段。
    返回: [(seg_start, seg_end, days, rate%, interest), ...]
    """
    if start_date > end_date or principal <= 0:
        return []
    segments = []
    current = start_date
    while current <= end_date:
        if rate_schedule:
            next_change = None
            for chg_date, _ in sorted(rate_schedule):
                if chg_date > current:
                    next_change = chg_date
                    break
            seg_end = min(end_date, next_change - timedelta(days=1)) if next_change else end_date
            rate = None
            for chg_date, r in sorted(rate_schedule):
                if chg_date <= current:
                    rate = r
            if rate is None:
                raise ValueError(f"无法确定 {current} 对应的利率")
        else:
            # LPR 每月20号调整
            y, m = current.year, current.month
            if current.day <= 20:
                cand = date(y, m, 20)
                if cand > current:
                    next_lpr = cand
                elif m < 12:
                    next_lpr = date(y, m + 1, 20)
                else:
                    next_lpr = date(y + 1, 1, 20)
            else:
                if m < 12:
                    next_lpr = date(y, m + 1, 20)
                else:
                    next_lpr = date(y + 1, 1, 20)
            seg_end = min(end_date, next_lpr - timedelta(days=1))
            rate = get_lpr_rate(current)
            if rate is None:
                raise ValueError(f"无法获取 {current} 的LPR")
        days = (seg_end - current).days + 1   # 含头含尾：截止日 - 起息日 + 1
        interest = round(principal * rate / 100 * days / 360, 2)
        segments.append((current, seg_end, days, rate, interest))
        current = seg_end + timedelta(days=1)
    return segments

def calc_double_delay(principal, start_date, end_date):
    """加倍部分债务利息：基数 × 1.75/10000 × 天数（算头不算尾）。

    ★ 基数口径（2026-09-14 依权威口径修正）：仅"原本之债（本金）"——诉讼费（受理费）、
      保全费、执行费、鉴定费等一概不计入（最高人民法院执行局《理解与适用》）；一般债务
      利息亦不计（法释〔2014〕8号第1条）。
    ★ start_date 应为"生效法律文书确定的履行期间届满之日的次日"（届满之日不计入本日）。
    """
    if start_date > end_date or principal <= 0:
        return 0.0, 0
    days = (end_date - start_date).days + 1   # 含头含尾：截止日 - 起息日 + 1
    amount = round(principal * 1.75e-4 * days, 2)
    return amount, days

def calc_exec_fee(amount):
    """申请执行费（《诉讼费用交纳办法》（国务院令第481号）第14条第(一)项第2目）：
    执行金额 ≤1万：每件50元；1万~50万部分 1.5%；50万~500万部分 1%；
    500万~1000万部分 0.5%；>1000万部分 0.1%。分段累计。
    负担：依同办法第38条，申请执行费由被执行人负担。
    执行标的金额由用户提供；须先确认案件已进入执行程序方可计入。
    """
    if amount is None or amount <= 0:
        return 0.0
    if amount <= 10000:
        return 50.0
    fee = 50.0
    fee += (min(amount, 500000) - 10000) * 0.015
    if amount > 500000:
        fee += (min(amount, 5000000) - 500000) * 0.01
    if amount > 5000000:
        fee += (min(amount, 10000000) - 5000000) * 0.005
    if amount > 10000000:
        fee += (amount - 10000000) * 0.001
    return round(fee, 2)

# ============================================================
# 4. 样式工厂
# ============================================================
FONT_CN = "宋体"
FONT_EN = "Times New Roman"

def make_font(name=FONT_EN, size=20, bold=False, color="000000", italic=False, underline=None):
    return Font(name=name, size=size, bold=bold, color=color,
                italic=False, underline=underline)

# 边框元件
thin_side = Side(style="thin", color="000000")
thick_side = Side(style="medium", color="000000")
none_side = Side(style=None)

def make_border(left=none_side, right=none_side, top=none_side, bottom=none_side):
    return Border(left=left, right=right, top=top, bottom=bottom)

# 填充
FILL_EVENT = PatternFill(start_color="D9D9D9", end_color="D9D9D9",
                         fill_type="solid")

# 表头
FILL_HEADER = PatternFill(fill_type=None)

# 对齐
align_center = Alignment(horizontal="center", vertical="center", wrap_text=True)
align_left = Alignment(horizontal="left", vertical="center", wrap_text=True)
align_right = Alignment(horizontal="right", vertical="center", wrap_text=True)

# 会计格式（千分位、2位小数、负数括号、零显示"-"）
ACCT_FMT = '#,##0.00;(#,##0.00);"-"'
DATE_FMT = 'yyyy"年"mm"月"dd"日"'

# 百分比格式
PERCENT_FMT = '0.00%'

# ============================================================
# 4b. 工具函数
# ============================================================
def format_amount(x):
    """格式化金额，千分位+2位小数"""
    return f"{x:,.2f}"

# ============================================================
# 5. 主生成函数
# ============================================================
def generate_debt_excel(params, output_path):
    """
    v5.0 公式联动版。
    所有数值源头 = 基本信息表（第 5~10 行）。
    利息详单 C/F/G 列 = 公式。
    加倍迟延 C/F 列 = 公式。
    汇总区全部 = 公式引用。
    总计欠付 = SUM(汇总)。
    """
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "债务计算表"
    ws.sheet_properties.tabColor = None

    widths = {1: 30, 2: 30, 3: 30, 4: 30, 5: 30, 6: 30, 7: 30}
    for c, w in widths.items():
        ws.column_dimensions[get_column_letter(c)].width = w

    def sc(r, c, val, font=None, border=None, align=align_center,
           fill=None, number_format=None, row_height=None):
        cell = ws.cell(row=r, column=c, value=val)
        if font: cell.font = font
        if border: cell.border = border
        if align: cell.alignment = align
        if fill: cell.fill = fill
        if number_format: cell.number_format = number_format
        if row_height: ws.row_dimensions[r].height = row_height
        return cell

    def box_border(r1, r2, c1, c2):
        for ri in range(r1, r2 + 1):
            for ci in range(c1, c2 + 1):
                cell = ws.cell(row=ri, column=ci)
                b = make_border(thin_side, thin_side, thin_side, thin_side)
                if ri == r1: b.top = thick_side
                if ri == r2: b.bottom = thick_side
                if ci == c1: b.left = thick_side
                if ci == c2: b.right = thick_side
                cell.border = b

    r = 1

    # ---------- 板块0：标题 ----------
    ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=7)
    sc(r, 1, "××××公司 债务计算表",
       font=make_font(FONT_CN, 28, bold=True),
       align=align_center)
    box_border(r, r, 1, 7)
    r += 1

    # ---------- 板块1：基本信息（★数值源头★）----------
# 动态构建基本信息行
    row_map = {}  # 记录每个标签对应的行号
    basic_info = [
        ("案号",             params['case_no'], 'str'),
        ("案涉项目",         to_full_project_name(params['project']), 'str'),
        ("债权人",           params['payee'], 'str'),
    ]
    # 本金类（金额为 0/空的项不写入表格）
    for item in params.get('principal_items', []):
        if not item.get('amount'):
            continue
        label = item.get('label') or (item['name'] + ('（计息）' if item.get('with_interest') else '（不计息）'))
        basic_info.append((label, item['amount'], 'num'))
    # 履约保证金（判项无则不写）
    if params.get('deposit'):
        basic_info.append(("履约保证金（不计息）", params['deposit'], 'num'))
    # 费用类（金额为 0/空的项不写入表格；可含执行费等）
    for item in params.get('fee_items', []):
        if not item.get('amount'):
            continue
        basic_info.append((item['name'], item['amount'], 'num'))
    # 日期类
    basic_info.append(("利息起算日",       fmt_date(params['interest_start']), 'str'))
    if params.get('include_double_delay', True):
        basic_info.append(("加倍迟延起算日",   fmt_date(params['double_delay_start']), 'str'))
    basic_info.append(("利息计算截止日",   fmt_date(params['interest_end']), 'str'))
    # 利息上限（如有）
    interest_cap = params.get('interest_cap', None)
    if interest_cap:
        basic_info.append(("利息总额上限", interest_cap, 'num'))

    ROW_INFO_START = r
    for label, val, dtype in basic_info:
        ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=3)
        ws.merge_cells(start_row=r, start_column=4, end_row=r, end_column=7)
        sc(r, 1, label,
           font=make_font(FONT_CN, 20, bold=True),
           align=align_center)
        if dtype == 'num':
            sc(r, 4, val,
               font=make_font(FONT_EN, 20),
               align=align_center,
               number_format=ACCT_FMT)
        else:
            sc(r, 4, val,
               font=make_font(FONT_CN, 20),
               align=align_center)
        row_map[label] = r  # 记录行号
        r += 1
    ROW_INFO_END = r - 1
    box_border(ROW_INFO_START, ROW_INFO_END, 1, 7)

    # 从 row_map 中提取关键行号（供后续公式引用）
    def get_row(keywords):
        for lbl, row in row_map.items():
            if any(kw in lbl for kw in keywords):
                return row
        return None

    ROW_PRINCIPAL     = get_row(['本金（计息）']) or (ROW_INFO_START + 3)
    ROW_DEPOSIT       = get_row(['履约保证金']) or (ROW_INFO_START + 4)
    ROW_ACCEPTANCE    = get_row(['受理费']) or (ROW_INFO_START + 5)
    ROW_PRESERVATION  = get_row(['保全费']) or (ROW_INFO_START + 6)

    # ---------- 板块2：利息计算详单 ----------
    r += 1
    ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=7)
    sc(r, 1, "利息计算详单",
       font=make_font(FONT_CN, 20, bold=True),
       align=align_center)
    box_border(r, r, 1, 7)
    r += 1
    ROW_INT_START = r

    headers = ["起息日", "截止日", "计息本金", "天数", "年利率", "本段利息", "挂账利息"]
    for i, h in enumerate(headers, 1):
        sc(r, i, h,
           font=make_font(FONT_CN, 20, bold=True, color="000000"),
           align=align_center,
           fill=FILL_HEADER)
    r += 1

    # ---- 准备数据：LPR 分段 + 付款冲抵 ----
    principal0    = params['principal']
    deposit_amt   = params.get('deposit', 0)
    interest_start = params['interest_start']
    interest_end   = params['interest_end']
    double_delay_start = params['double_delay_start']
    payments      = sorted(params.get('payments', []), key=lambda x: x[0])
    rate_schedule = params.get('rate_schedule', None)
    interest_cap  = params.get('interest_cap', None)

    # 生成所有原始段（含付款冲抵）
    remaining_principal = principal0
    accrued_interest = 0.0  # 已挂账未付利息
    raw_segs = []
    prev_date = interest_start
    pmt_iter = iter(payments)
    next_pmt = next(pmt_iter, None)

    while prev_date <= interest_end:
        seg_end = interest_end
        if rate_schedule:
            for chg_date, _ in sorted(rate_schedule):
                if chg_date > prev_date:
                    seg_end = min(seg_end, chg_date - timedelta(days=1))
                    break
        else:
            y, m = prev_date.year, prev_date.month
            if prev_date.day <= 20:
                cand = date(y, m, 20)
                if cand > prev_date:
                    nxt = cand
                elif m < 12:
                    nxt = date(y, m + 1, 20)
                else:
                    nxt = date(y + 1, 1, 20)
            else:
                if m < 12:
                    nxt = date(y, m + 1, 20)
                else:
                    nxt = date(y + 1, 1, 20)
            seg_end = min(seg_end, nxt - timedelta(days=1))

        # 付款日截断本段
        if next_pmt and next_pmt[0] > prev_date:
            seg_end = min(seg_end, next_pmt[0] - timedelta(days=1))
        seg_end = min(seg_end, interest_end)

        if seg_end >= prev_date:
            segs = calc_interest_segment(remaining_principal, prev_date, seg_end, rate_schedule)
            for (ss, se, days, rate, inte) in segs:
                raw_segs.append({
                    'start': ss, 'end': se,
                    'days': days, 'rate': rate,
                    'base': remaining_principal,
                })
            # 本段利息累加到挂账
            seg_interest = sum(v for *_, v in segs)
            accrued_interest += seg_interest

            # 付款处理：先冲利息，余下冲本金
            if next_pmt and next_pmt[0] == seg_end + timedelta(days=1) and next_pmt[0] <= interest_end:
                pmt_date, pmt_amt = next_pmt
                if pmt_amt > 0:
                    if pmt_amt <= accrued_interest:
                        # 全部冲利息，本金不变
                        accrued_interest -= pmt_amt
                    else:
                        # 利息冲完，剩余冲本金
                        remainder = pmt_amt - accrued_interest
                        accrued_interest = 0.0
                        remaining_principal = max(0, remaining_principal - remainder)
                prev_date = pmt_date
                next_pmt = next(pmt_iter, None)
                continue
            prev_date = seg_end + timedelta(days=1)
        else:
            prev_date = seg_end + timedelta(days=1)

    # 合并同 base + 同 rate 的连续段
    merged = []
    i = 0
    while i < len(raw_segs):
        cur = raw_segs[i]
        j = i
        while (j + 1 < len(raw_segs) and
               raw_segs[j + 1]['base'] == cur['base'] and
               raw_segs[j + 1]['rate'] == cur['rate']):
            j += 1
        if j > i:
            total_days = sum(raw_segs[k]['days'] for k in range(i, j + 1))
            merged.append({
                'start': cur['start'],
                'end': raw_segs[j]['end'],
                'days': total_days,
                'rate': cur['rate'],
                'base': cur['base'],
                'merged': True,
            })
        else:
            merged.append(raw_segs[i])
        i = j + 1

    # ---- 写起始行（利息起算行：只填起息日）----
    sc(r, 1, interest_start,
       font=make_font(FONT_EN, 20), align=align_center, number_format=DATE_FMT)
    sc(r, 2, '', font=make_font(FONT_EN, 20))
    # C列 = 公式引用基本信息本金
    sc(r, 3, f"=D{ROW_PRINCIPAL}",
       font=make_font(FONT_EN, 20), align=align_center,
       number_format=ACCT_FMT)
    sc(r, 4, '', font=make_font(FONT_EN, 20))
    sc(r, 5, '', font=make_font(FONT_EN, 20))
    sc(r, 6, '', font=make_font(FONT_EN, 20))
    # G列起始 = 0（挂账利息起点）
    sc(r, 7, 0,
       font=make_font(FONT_EN, 20), align=align_center,
       number_format=ACCT_FMT)
    for c in range(1, 8):
        ws.cell(row=r, column=c).fill = FILL_EVENT
    r += 1

    # ---- 写合并后的利息段（D=公式, F=公式, G=累计公式）----
    prev_g_ref = f"G{ROW_INT_START + 1}"  # 起始行 G
    last_g_ref = prev_g_ref
    for seg in merged:
        # A=起息日，B=截止日（均为真日期值，便于 D 列相减）
        sc(r, 1, seg['start'],
           font=make_font(FONT_EN, 20), align=align_center, number_format=DATE_FMT)
        sc(r, 2, seg['end'],
           font=make_font(FONT_EN, 20), align=align_center, number_format=DATE_FMT)
        # C列 = 公式引用本金（若有付款减损，后续可手动改此公式）
        sc(r, 3, f"=D{ROW_PRINCIPAL}",
           font=make_font(FONT_EN, 20), align=align_center,
           number_format=ACCT_FMT)
        # ★ D = 截止日 - 起息日 + 1（含头含尾）
        sc(r, 4, f"=B{r}-A{r}+1",
           font=make_font(FONT_EN, 20), align=align_center)
        sc(r, 5, seg['rate'] / 100,
           font=make_font(FONT_EN, 20), align=align_center,
           number_format=PERCENT_FMT)
        # ★ F = C × D × E / 360
        sc(r, 6, f"=C{r}*D{r}*E{r}/360",
           font=make_font(FONT_EN, 20), align=align_center,
           number_format=ACCT_FMT)
        # ★ G = 上一行G + 本行F
        sc(r, 7, f"={prev_g_ref}+F{r}",
           font=make_font(FONT_EN, 20), align=align_center,
           number_format=ACCT_FMT)
        prev_g_ref = f"G{r}"
        last_g_ref = prev_g_ref
        r += 1

    ROW_INT_END = r - 1
    box_border(ROW_INT_START, ROW_INT_END, 1, 7)

    # ---------- 板块3：加倍迟延利息（include_double_delay=False 则整块省略）----------
    _include_dd = params.get('include_double_delay', True)
    ROW_DD_PRINCIPAL = None
    ROW_DD_DEPOSIT = None
    dd_days = 0
    if _include_dd:
        r += 1
        ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=7)
        sc(r, 1, "加倍迟延利息计算详单",
           font=make_font(FONT_CN, 20, bold=True),
           align=align_center)
        box_border(r, r, 1, 7)
        r += 1
        ROW_DD_START = r

        dd_headers = ["基数构成", "起息日", "截止日", "基数金额", "天数", "日利率", "金额"]
        for i, h in enumerate(dd_headers, 1):
            sc(r, i, h,
               font=make_font(FONT_CN, 20, bold=True, color="000000"),
               align=align_center,
               fill=FILL_HEADER)
        r += 1

        dd_base_principal = remaining_principal
        dd_base_deposit   = deposit_amt
        dd_amt_p, dd_days = calc_double_delay(dd_base_principal,
                                               double_delay_start, interest_end)
        dd_amt_d, _       = calc_double_delay(dd_base_deposit,
                                               double_delay_start, interest_end)

        # 本金行
        sc(r, 1, "本金",
           font=make_font(FONT_CN, 20), align=align_center)
        if dd_days > 0:
            sc(r, 2, double_delay_start,
               font=make_font(FONT_EN, 20), align=align_center, number_format=DATE_FMT)
            sc(r, 3, interest_end,
               font=make_font(FONT_EN, 20), align=align_center, number_format=DATE_FMT)
        else:
            sc(r, 2, '', font=make_font(FONT_EN, 20))
            sc(r, 3, '', font=make_font(FONT_EN, 20))
        # D列 = 公式引用基本信息本金
        sc(r, 4, f"=D{ROW_PRINCIPAL}",
           font=make_font(FONT_EN, 20), align=align_center,
           number_format=ACCT_FMT)
        # ★ E = 截止日 - 起息日 + 1（含头含尾）
        sc(r, 5, f"=C{r}-B{r}+1" if dd_days > 0 else '',
           font=make_font(FONT_EN, 20), align=align_center)
        sc(r, 6, "1.75/10000",
           font=make_font(FONT_EN, 20), align=align_center)
        if dd_days > 0:
            # ★ G = D × E × 1.75 / 10000（日利率，不除360）
            sc(r, 7, f"=D{r}*E{r}*1.75/10000",
               font=make_font(FONT_EN, 20), align=align_center,
               number_format=ACCT_FMT)
        else:
            sc(r, 7, 0,
               font=make_font(FONT_EN, 20), align=align_center,
               number_format=ACCT_FMT)
        ROW_DD_PRINCIPAL = r
        r += 1

        # 保证金行（仅当判项中存在履约保证金时写入；无则不列）
        ROW_DD_DEPOSIT = None
        if deposit_amt > 0:
            sc(r, 1, "履约保证金",
               font=make_font(FONT_CN, 20), align=align_center)
            if dd_days > 0:
                sc(r, 2, double_delay_start,
                   font=make_font(FONT_EN, 20), align=align_center, number_format=DATE_FMT)
                sc(r, 3, interest_end,
                   font=make_font(FONT_EN, 20), align=align_center, number_format=DATE_FMT)
            else:
                sc(r, 2, '', font=make_font(FONT_EN, 20))
                sc(r, 3, '', font=make_font(FONT_EN, 20))
            # D列 = 公式引用基本信息保证金
            sc(r, 4, f"=D{ROW_DEPOSIT}",
               font=make_font(FONT_EN, 20), align=align_center,
               number_format=ACCT_FMT)
            sc(r, 5, f"=C{r}-B{r}+1" if dd_days > 0 else '',
               font=make_font(FONT_EN, 20), align=align_center)
            sc(r, 6, "1.75/10000",
               font=make_font(FONT_EN, 20), align=align_center)
            if dd_days > 0:
                sc(r, 7, f"=D{r}*E{r}*1.75/10000",
                   font=make_font(FONT_EN, 20), align=align_center,
                   number_format=ACCT_FMT)
            else:
                sc(r, 7, 0,
                   font=make_font(FONT_EN, 20), align=align_center,
                   number_format=ACCT_FMT)
            ROW_DD_DEPOSIT = r
            r += 1

        ROW_DD_END = r - 1
        box_border(ROW_DD_START, ROW_DD_END, 1, 7)

    # ---------- 板块4：汇总（全部公式）----------
    r += 1
    ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=7)
    sc(r, 1, "汇总",
       font=make_font(FONT_CN, 20, bold=True),
       align=align_center)
    box_border(r, r, 1, 7)
    r += 1
    ROW_SUM_START = r

    # 按实际存在的项目动态生成汇总行（无金额项不列；每行 A:C 合并 + D:G 合并）
    sum_items = []          # [(标签, 值公式)]
    if principal0 > 0 and not params.get('principal_settled'):
        sum_items.append((params.get('principal_label') or "本金", f"=D{ROW_PRINCIPAL}"))
    if deposit_amt > 0:
        sum_items.append(("履约保证金", f"=D{ROW_DEPOSIT}"))
    if principal0 > 0:
        sum_items.append((params.get('interest_label', '利息'), f"={last_g_ref}"))
    if dd_days > 0 and _include_dd:
        dd_formula = f"=G{ROW_DD_PRINCIPAL}" + (f"+G{ROW_DD_DEPOSIT}" if ROW_DD_DEPOSIT else "")
        sum_items.append(("加倍迟延利息", dd_formula))
    for item in params.get('fee_items', []):
        if not item.get('amount'):
            continue
        lbl = item['name']
        if lbl in row_map:
            sum_items.append((lbl, f"=D{row_map[lbl]}"))

    for label, formula in sum_items:
        ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=3)
        ws.merge_cells(start_row=r, start_column=4, end_row=r, end_column=7)
        sc(r, 1, label,
           font=make_font(FONT_CN, 20, bold=True),
           align=align_center)
        sc(r, 4, formula,
           font=make_font(FONT_EN, 20), align=align_center,
           number_format=ACCT_FMT)
        r += 1

    # 总计欠付 = 各汇总行之和（全公式）
    ROW_SUM_GRAND = r
    ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=3)
    ws.merge_cells(start_row=r, start_column=4, end_row=r, end_column=7)
    sc(r, 1, "总计欠付",
       font=make_font(FONT_CN, 20, bold=True),
       align=align_center)
    parts = [f"D{ROW_SUM_START + i}" for i in range(len(sum_items))]
    sc(r, 4, "=" + "+".join(parts),
       font=make_font(FONT_EN, 20, bold=True), align=align_center,
       number_format=ACCT_FMT)
    ROW_SUM_END = r
    box_border(ROW_SUM_START, ROW_SUM_END, 1, 7)

    # ★ v5.0：不再有"待履行总计"行，总计欠付就是最后一行 ★

    # 打印设置
    ws.sheet_properties.pageSetUpPr = PageSetupProperties(fitToPage=True)
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.page_setup.orientation = 'landscape'
    ws.page_setup.paperSize = 9

    wb.save(output_path)
    print(f"[OK] 已生成：{output_path}")
    print(f" [公式联动] 基本信息表 → 计算详单 → 汇总表 → 总计欠付")
    return output_path

# ============================================================
# ★ v5.0 工作流函数
# ============================================================

# ------------------------------------------------------------
# 排除的己方关键词
# ------------------------------------------------------------
EXCLUDE_OUR_SIDE = [
    "××××公司", "公司", "珠海分公司", "××分公司", "广州分公司",
    "广东总承包", "基础设施分公司", "城市管网",
    "××××工程局", "××××有限公司",
]


# ------------------------------------------------------------
# B. 内置备用 OCR
# ------------------------------------------------------------
def _fallback_ocr(pdf_path):
    try:
        import fitz
        doc = fitz.open(str(pdf_path))
        texts = []
        for i in range(doc.page_count):
            t = doc[i].get_text().strip()
            if t:
                texts.append(f"===== 第{i+1}页 =====\n{t}")
        doc.close()
        if texts:
            return "\n\n".join(texts)
    except Exception:
        pass
    txt = Path(pdf_path).with_suffix('.txt')
    if txt.exists():
        return txt.read_text(encoding='utf-8')
    return ""


def extract_pdf_text(pdf_path):
    if _OCR_AVAILABLE:
        try:
            return ocr_extract_text(str(pdf_path), verbose=True)
        except Exception as e:
            print(f"[OCR] ocr.py调用失败: {e}，使用备用方案")
    return _fallback_ocr(pdf_path)


# ------------------------------------------------------------
# C. 判决书信息提取
# ------------------------------------------------------------
def _simplify_name(name):
    """简化当事人名称：去掉分公司、公司类型后缀"""
    if not name:
        return ""
    for s in ["（深圳）", "(深圳)", "（广州）", "(广州)", "（东莞）", "(东莞)",
              "有限公司", "有限责任公司", "股份有限公司", "集团"]:
        name = name.replace(s, "")
    return name.strip()

def _clean_case_num(s):
    s = s.replace(" ", "").replace("\u3000", "")
    return s.replace("(", "（").replace(")", "）")

def extract_case_number(text):
    norm = text.replace('(', '（').replace(')', '）')
    for pat in [
        r'（\s*\d{4}\s*）[\s\S]{0,20}?民初[\s\S]{0,10}?号',
        r'（\s*\d{4}\s*）[\s\S]{0,20}?民终[\s\S]{0,10}?号',
        r'（\s*\d{4}\s*）[\s\S]{0,30}?号',
    ]:
        m = re.search(pat, norm)
        if m:
            return _clean_case_num(m.group(0))
    return ""

def extract_all_case_numbers(text):
    raw = re.findall(r'[（(]\s*\d{4}\s*[）)][\s\S]{0,30}?号', text)
    return [_clean_case_num(m) for m in raw]

def extract_counterparty(text):
    patterns = [
        r'原告[（(][^）)]*[）)]\s*[：:]\s*([^\n,，;；]+)',
        r'上诉人[（(][^）)]*[）)]\s*[：:]\s*([^\n,，;；]+)',
        r'申请人[（(][^）)]*[）)]\s*[：:]\s*([^\n,，;；]+)',
        r'原告[：:]\s*([^\n,，;；]+)',
        r'上诉人[：:]\s*([^\n,，;；]+)',
    ]
    found = []
    for p in patterns:
        for m in re.findall(p, text):
            m = m.strip().rstrip(',，;；')
            if m and not any(ex in m for ex in EXCLUDE_OUR_SIDE):
                found.append(_simplify_name(m))
    seen, unique = set(), []
    for f in found:
        if f not in seen:
            seen.add(f)
            unique.append(f)
    return "、" .join(unique) if unique else ""

def extract_project_name(text):
    for pat in [
        r'[关关于][于\s]*([\u4e00-\u9fa5（）()]{4,30}?(?:项目|工程))',
        r'在([\u4e00-\u9fa5（）()]{4,30}?(?:项目|工程))',
        r'对([\u4e00-\u9fa5（）()]{4,30}?(?:项目|工程))',
    ]:
        m = re.search(pat, text)
        if m:
            return m.group(1).strip()
    return ""

def extract_principal(text):
    m = re.search(r'支付[^\n]{0,50}?(\d{1,10}(?:,\d{3})*\.?\d*)\s*元', text)
    if m:
        try:
            val = float(m.group(1).replace(',', ''))
            if val > 10000:
                return val
        except ValueError:
            pass
    for pat in [r'本金[为]*\s*([\d,]+\.?\d*)\s*元',
                r'[返支付还][还付]?\s*本金\s*([\d,]+\.?\d*)\s*元']:
        m = re.search(pat, text)
        if m:
            try:
                return float(m.group(1).replace(',', ''))
            except ValueError:
                pass
    nums = []
    for a in re.findall(r'([\d,]+\.?\d*)\s*元', text):
        try:
            nums.append(float(a.replace(',', '')))
        except ValueError:
            continue
    big = [n for n in nums if n > 10000]
    return max(big) if big else 0.0

def extract_deposit(text):
    for pat in [r'履约保证金[为]?\s*([\d,]+\.?\d*)\s*元',
                r'保证金[为]?\s*([\d,]+\.?\d*)\s*元']:
        m = re.search(pat, text)
        if m:
            try:
                return float(m.group(1).replace(',', ''))
            except ValueError:
                pass
    return 0.0

def extract_acceptance_fee(text):
    for pat in [r'受理费[为]?\s*([\d,]+\.?\d*)\s*元',
                r'案件受理费[为]?\s*([\d,]+\.?\d*)\s*元']:
        m = re.search(pat, text)
        if m:
            try:
                return float(m.group(1).replace(',', ''))
            except ValueError:
                pass
    return 0.0

def extract_preservation_fee(text):
    m = re.search(r'保全费[为]?\s*([\d,]+\.?\d*)\s*元', text)
    if m:
        try:
            return float(m.group(1).replace(',', ''))
        except ValueError:
            pass
    return 0.0

def extract_interest_cap(text):
    for pat in [
        r'利息总额[以用]?\s*不[超多]?\s*过?\s*([\d,]+\.?\d*)\s*元',
        r'以([\d,]+\.?\d*)\s*元\s*为\s*限',
        r'上限[为]?\s*([\d,]+\.?\d*)\s*元',
    ]:
        m = re.search(pat, text)
        if m:
            try:
                return float(m.group(1).replace(',', ''))
            except ValueError:
                pass
    return 0.0

def extract_interest_start_date(text):
    for pat in [
        r'自\s*(\d{4})\s*年\s*(\d{1,2})\s*月\s*(\d{1,2})\s*日[起开]',
        r'从\s*(\d{4})\s*年\s*(\d{1,2})\s*月\s*(\d{1,2})\s*日起',
        r'起诉之日[即为是]*\s*(\d{4})\s*年\s*(\d{1,2})\s*月\s*(\d{1,2})\s*日',
    ]:
        m = re.search(pat, text)
        if m:
            try:
                return date(int(m.group(1)), int(m.group(2)), int(m.group(3)))
            except ValueError:
                pass
    return None

def extract_judgment_date(text):
    for pat in [
        r'(\d{4})\s*年\s*(\d{1,2})\s*月\s*(\d{1,2})\s*日[，,。]?\s*判决如下',
        r'(\d{4})\s*年\s*(\d{1,2})\s*月\s*(\d{1,2})\s*日[作做]出',
    ]:
        m = re.search(pat, text)
        if m:
            try:
                return date(int(m.group(1)), int(m.group(2)), int(m.group(3)))
            except ValueError:
                pass
    ms = re.findall(r'(\d{4})\s*年\s*(\d{1,2})\s*月\s*(\d{1,2})\s*日', text)
    if ms:
        y, mth, d = ms[-1]
        try:
            return date(int(y), int(mth), int(d))
        except ValueError:
            pass
    return None

def is_second_instance(text, case_num=""):
    if "民终" in case_num:
        return True
    return False

def parse_iso_date(s):
    s = s.strip()
    if not s:
        return None
    for fmt in ["%Y-%m-%d", "%Y/%m/%d", "%Y年%m月%d日"]:
        try:
            return datetime.strptime(s, fmt).date()
        except ValueError:
            continue
    if s.isdigit() and len(s) == 8:
        try:
            return date(int(s[:4]), int(s[4:6]), int(s[6:8]))
        except ValueError:
            pass
    return None


# ------------------------------------------------------------
# D. 交互式信息确认
# ------------------------------------------------------------
def _prompt_float(label, default=0.0):
    while True:
        s = input(f"  {label} [{default:,.2f}]: ").strip()
        if not s:
            return default
        try:
            return float(s.replace(',', ''))
        except ValueError:
            print("    [!] 请输入有效数字")

def _prompt_date(label, default=None):
    ds = fmt_date(default) if default else "未定"
    while True:
        s = input(f"  {label} [{ds}]: ").strip()
        if not s:
            if default:
                return default
            print("    [!] 必须输入日期")
            continue
        d = parse_iso_date(s)
        if d:
            return d
        print("    [!] 日期格式无法识别")

def _prompt_str(label, default=""):
    s = input(f"  {label} [{default}]: ").strip()
    return s if s else default

def _confirm(label, default=True):
    d = "Y/n" if default else "y/N"
    s = input(f"  {label} [{d}]: ").strip().lower()
    if not s:
        return default
    return s in ("y", "yes", "是", "确认", "ok")

def interactive_collect_info(text, file_path):
    print(f"\n{'='*60}\n  Step 1: OCR识别 & 信息提取\n  文件: {file_path.name}\n{'='*60}")

    # 提取所有案号
    all_n = extract_all_case_numbers(text)
    cn = extract_case_number(text)
    cp = extract_counterparty(text)
    prj = extract_project_name(text)
    pri = extract_principal(text)
    dep = extract_deposit(text)
    acp = extract_acceptance_fee(text)
    prv = extract_preservation_fee(text)
    ist = extract_interest_start_date(text)
    jdt = extract_judgment_date(text)
    cap = extract_interest_cap(text)
    sec = is_second_instance(text, cn)

    print("\n[自动提取结果]")
    print(f"  案号:           {cn or '(未找到)'}   ★ 一律填一审案号（二审/终审案件也用其一审案号）")
    print(f"  对方当事人:     {cp or '(未找到)'}")
    print(f"  项目名:         {prj or '(未找到)'}")
    print(f"  本金:           {format_amount(pri) if pri else '(未找到)'}")
    print(f"  履约保证金:     {format_amount(dep) if dep else '(未找到)'}")
    print(f"  受理费:         {format_amount(acp) if acp else '(未找到)'}")
    print(f"  保全费:         {format_amount(prv) if prv else '(未找到)'}")
    print(f"  利息起息日:     {fmt_date(ist) if ist else '(未找到)'}")
    print(f"  判决作出日:     {fmt_date(jdt) if jdt else '(未找到)'}")
    print(f"  利息总额上限:    {format_amount(cap) if cap else '(未找到)'}")
    print(f"  是否二审:       {'是' if sec else '否（一审）'}")

    print(f"\n{'─'*40}\n  请确认/修正以下信息：\n{'─'*40}")
    if not cn or not _confirm(f"案号 '{cn}' 正确？", True):
        cn = _prompt_str("案号", cn)
    if not cp or not _confirm(f"对方当事人 '{cp}' 正确？", True):
        cp = _prompt_str("对方当事人", cp)
    if not prj or not _confirm(f"项目名 '{prj}' 正确？", True):
        prj = _prompt_str("项目名", prj)
    if not pri or not _confirm(f"本金 {format_amount(pri)} 正确？", True):
        pri = _prompt_float("本金", pri)
    if not dep or not _confirm(f"履约保证金 {format_amount(dep)} 正确？", True):
        dep = _prompt_float("保证金（无则0）", dep)
    if not acp or not _confirm(f"受理费 {format_amount(acp)} 正确？", True):
        acp = _prompt_float("受理费（无则0）", acp)
    if not prv or not _confirm(f"保全费 {format_amount(prv)} 正确？", True):
        prv = _prompt_float("保全费（无则0）", prv)
    if not cap or not _confirm(f"利息总额上限 {format_amount(cap)} 正确？", True):
        cap = _prompt_float("利息总额上限（无则0）", cap)
    if not ist or not _confirm(f"利息起息日 {fmt_date(ist)} 正确？", True):
        ist = _prompt_date("利息起息日", ist)
    if not jdt or not _confirm(f"判决作出之日 {fmt_date(jdt)} 正确？", True):
        jdt = _prompt_date("判决作出之日", jdt)

    print(f"\n  ★ 收到判决之日是计算加倍迟延利息起息日的关键参数 ★")
    rcv = _prompt_date("请输入收到判决书之日")

    if sec:
        print(f"  [二审/终审] 判决送达即生效")
        eff = 0
        if not _confirm("二审生效无需等待，是否正确？", True):
            eff = int(_prompt_str("生效等待天数", "0") or 0)
    else:
        print(f"  [一审] 民事判决上诉期为15日")
        eff = 15
        if not _confirm("一审生效期限为15日，是否正确？", True):
            eff = int(_prompt_str("生效期限天数", "15") or 15)

    perf = 5  # 默认5日
    ps = _prompt_str(f"履行期限天数（默认{perf}）", str(perf))
    perf = int(ps) if ps.isdigit() else perf

    if rcv:
        dds = rcv + timedelta(days=eff + perf)
        print(f"\n  [自动计算] 加倍迟延利息起息日 = {fmt_date(rcv)}(收到) +{eff}(生效) +{perf}(履行) = {fmt_date(dds)}")
    else:
        dds = None

    today = date.today()
    if _confirm(f"利息计算截止日设为今天 {fmt_date(today)}？", True):
        iend = today
    else:
        iend = _prompt_date("利息计算截止日", today)

    # 确认表
    print(f"\n{'='*60}\n  Step 2: 基础信息确认表\n{'='*60}")
    tbl = [
        ("案号", cn), ("案涉项目", prj), ("对方当事人", cp),
        ("本金（计息）", format_amount(pri)), ("履约保证金", format_amount(dep)),
        ("受理费", format_amount(acp)), ("保全费", format_amount(prv)),
        ("利息总额上限", format_amount(cap) if cap else "无"),
        ("利息起算日", fmt_date(ist) if ist else "待定"),
        ("判决作出之日", fmt_date(jdt) if jdt else "待定"),
        ("收到判决之日", fmt_date(rcv) if rcv else "待定"),
        ("生效期限", f"{eff}日"), ("履行期限", f"{perf}日"),
        ("加倍迟延起算日", fmt_date(dds) if dds else "待定"),
        ("利息计算截止日", fmt_date(iend) if iend else "待定"),
        ("审级", "二审/终审" if sec else "一审"),
    ]
    for k, v in tbl:
        print(f"  {k:<14s} {v}")
    if not _confirm("\n以上信息是否全部正确，可以开始计算？", True):
        print("[!] 用户取消")
        return None

    return {
        'case_no': cn, 'project': prj, 'payee': cp,
        'principal_items': [
            {'name': '工程款本金', 'amount': pri, 'with_interest': True},
            {'name': '履约保证金', 'amount': dep, 'with_interest': False},
        ],
        'fee_items': [
            {'name': '受理费', 'amount': acp},
            {'name': '保全费', 'amount': prv},
        ],
        'interest_cap': cap if cap else None,
        'interest_start': ist, 'interest_end': iend,
        'double_delay_start': dds,
        'payments': [], 'rate_schedule': None,
    }

# ------------------------------------------------------------
# E. 归档（调用 naming_n_filing 命名 + 依据文件不改名）
# ------------------------------------------------------------
def archive_debt_file(excel_path, source_files, params):
    """
    v5.0：调用 naming_n_filing.generate_new_filename() 生成标准命名，
    并在 D 盘搜索已有案件文件夹复用。依据文件保持原名不变。
    """
    actions = {"created_dirs": [], "copied_files": [], "errors": []}

    cn = params.get('case_no', '')
    cp = params.get('payee', '')
    prj = params.get('project', '')
    sec = "民终" in cn
    is_second = bool(sec)

    # ---- 用 naming_n_filing 生成标准文件夹名 ----
    if _NAMING_AVAILABLE:
        # 拼接所有已知案号（含OCR提取到的全部）
        all_case_nums = [cn] if cn else []
        # 尝试搜索已有文件夹
        selected_dir = None
        try:
            _, _, _, selected_dir = nf.interactive_search_and_backfill(
                case_nums=all_case_nums,
                counterparty=cp,
                project=prj,
                text="",
                dirs_only=True,
            )
        except Exception as e:
            print(f"  [搜索] naming_n_filing 搜索失败: {e}")

        folder_name = nf.generate_new_filename(
            prefix="债务计算",
            counterparty=cp,
            project=prj,
            case_num=cn,
            is_second=is_second,
            text="",
        )
        print(f"  [命名] naming_n_filing → {folder_name}")
    else:
        # 降级：自己拼（保持与原规则一致）
        folder_name = f"债务计算【{cp}】{prj}{cn}"
        folder_name = re.sub(r'[\\/*?:"<>|]', "", folder_name)
        print(f"  [命名] 降级模式 → {folder_name}")
        selected_dir = None

    folder_name = re.sub(r'[\\/*?:"<>|]', "", folder_name)
    target_dir = D_DEBT_CALC / folder_name

    try:
        target_dir.mkdir(parents=True, exist_ok=True)
        actions["created_dirs"].append(str(target_dir))
        print(f"  [创建] {target_dir}")
    except Exception as e:
        actions["errors"].append(f"创建文件夹失败: {e}")
        return actions

    # ---- Excel 命名为 "债务计算" + naming_n_filing 后缀 ----
    if excel_path.exists():
        ext = excel_path.suffix
        # 用 naming_n_filing 生成的完整文件名（含前缀+后缀）
        if _NAMING_AVAILABLE:
            xl_name = folder_name + ext  # 文件夹名 = Excel 名（去掉扩展名差异）
        else:
            xl_name = f"债务计算{ext}"
        target_xl = target_dir / xl_name
        try:
            shutil.copy2(str(excel_path), str(target_xl))
            actions["copied_files"].append(str(target_xl))
            print(f"  [复制] {excel_path.name} → {xl_name}")
        except Exception as e:
            actions["errors"].append(f"复制Excel失败: {e}")

    # ---- 依据文件：保持原始文件名，不改名 ----
    for sf in source_files:
        sp = Path(sf)
        if not sp.exists():
            actions["errors"].append(f"依据文件不存在: {sf}")
            continue
        try:
            target_f = target_dir / sp.name  # ★ 保持原名 ★
            shutil.copy2(str(sp), str(target_f))
            actions["copied_files"].append(str(target_f))
            print(f"  [复制] {sp.name} → 保持原名")
        except Exception as e:
            actions["errors"].append(f"复制依据文件失败: {e}")

    return actions


# ------------------------------------------------------------
# F. 行动报告
# ------------------------------------------------------------
def print_action_report(params, excel_path, archive_result):
    print(f"\n{'='*60}\n  ★ 行动报告 ★\n{'='*60}")
    print(f"\n  【案件信息】")
    print(f"    案号:     {params.get('case_no','')}")
    print(f"    项目:     {params.get('project','')}")
    print(f"    当事人:   {params.get('payee','')}")
    print(f"\n  【生成文件】\n    ✓ Excel: {excel_path.name}")
    print(f"\n  【归档操作】")
    if archive_result["created_dirs"]:
        print(f"    [创建文件夹] {len(archive_result['created_dirs'])}个:")
        for d in archive_result["created_dirs"]:
            print(f"      + {d}")
    if archive_result["copied_files"]:
        print(f"    [复制文件] {len(archive_result['copied_files'])}个:")
        for f in archive_result["copied_files"]:
            print(f"      + {f}")
    if archive_result["errors"]:
        print(f"    [错误] {len(archive_result['errors'])}个:")
        for e in archive_result["errors"]:
            print(f"      ✗ {e}")
    print(f"\n{'='*60}\n  处理完成 ✓\n{'='*60}\n")


# ------------------------------------------------------------
# G. 主流程
# ------------------------------------------------------------
def process_debt_calc(file_paths):
    if not file_paths:
        print("[错误] 未提供任何文件")
        return
    paths = [Path(p) for p in file_paths]
    for p in paths:
        if not p.exists():
            print(f"[错误] 文件不存在: {p}")
            return
    primary = paths[0]
    supporting = paths[1:]

    print(f"\n{'#'*60}")
    print(f"#   债务计算工作流 v5.0")
    print(f"#   主文件: {primary.name}")
    if supporting:
        print(f"#   依据文件: {len(supporting)}个")
    print(f"{'#'*60}")

    # Step 1: OCR
    print(f"\n>>> Step 1/4: OCR识别主文件...")
    text = extract_pdf_text(primary)
    if not text.strip():
        print("[错误] OCR未能提取到文本")
        return
    print(f"  [OK] 提取文本 {len(text)} 字符")

    all_text = text
    for sf in supporting:
        if sf.suffix.lower() == '.pdf':
            t2 = extract_pdf_text(sf)
            if t2:
                all_text += "\n" + t2

    # Step 2: 信息提取与确认
    print(f"\n>>> Step 2/4: 信息提取与确认...")
    params = interactive_collect_info(all_text, primary)
    if params is None:
        return

    # Step 3: 生成 Excel
    print(f"\n>>> Step 3/4: 生成债务计算表...")
    temp = Path.cwd() / "债务计算_temp.xlsx"
    try:
        xl = generate_debt_excel(params, str(temp))
        xl = Path(xl)
    except Exception:
        import traceback
        traceback.print_exc()
        return

    # Step 4: 归档（命名 + 建文件夹 + 复制依据文件）
    print(f"\n>>> Step 4/4: 归档文件...")
    all_src = [str(primary)] + [str(p) for p in supporting]
    res = archive_debt_file(Path(xl), all_src, params)
    print_action_report(params, Path(xl), res)


# ------------------------------------------------------------
# H. CLI 入口
# ------------------------------------------------------------
if __name__ == "__main__":
    if sys.platform == "win32":
        try:
            sys.stdout.reconfigure(encoding="utf-8")
        except Exception:
            pass
    print("=" * 60)
    print("  债务计算表生成器 - 完整工作流版 v5.0")
    print("  功能：OCR识别 → 信息确认 → 公式Excel → 归档 → 报告")
    print("=" * 60)
    if len(sys.argv) > 1:
        fps = sys.argv[1:]
    else:
        print("\n请输入文件路径（多个用空格分隔，q退出）：")
        ui = input("> ").strip()
        if ui.lower() == 'q':
            sys.exit(0)
        fps = ui.split()
    if fps:
        try:
            process_debt_calc(fps)
        except KeyboardInterrupt:
            print("\n[中断] 用户取消")
        except Exception:
            import traceback
            traceback.print_exc()
