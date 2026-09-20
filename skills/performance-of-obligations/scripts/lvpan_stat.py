# -*- coding: utf-8 -*-
"""
履约判决跟踪底稿统计脚本（skill: lvpan-dibang 配套）
用法: python lvpan_stat.py <底稿路径> [截止日]
按 局/公司 名义 × 时间桶 统计"生效待履约跟踪金额"（Q优先，Q空白用H）
"""
import openpyxl, sys, io, json
from datetime import datetime, timedelta

sys.stdout.reconfigure(encoding='utf-8', errors='replace')

SRC = sys.argv[1] if len(sys.argv) > 1 else r'D:/【用户名】/【诉讼】/【执行履行】/履约判决跟踪/【履约判决跟踪】2026年9月/202609分公司履约跟踪.xlsx'
# 分桶截止日（2026-08-31 起，每段+3/4个月）
CUTS = [
    (datetime(2026, 8, 31).date(), '截至8月'),
    (datetime(2026, 12, 31).date(), '2026.09-12'),
    (datetime(2027, 3, 31).date(), '2027.01-03'),
    (datetime(2027, 12, 31).date(), '2027.04-12'),
]

def edate(v):
    if v is None or str(v).strip() in ('', 'None', '/'):
        return None
    if isinstance(v, (int, float)):
        try:
            return (datetime(1899, 12, 30) + timedelta(days=float(v))).date()
        except:
            return None
    try:
        return datetime.strptime(str(v)[:10], '%Y-%m-%d').date()
    except:
        return None

def bucket(dt):
    if dt is None:
        return len(CUTS)  # 更晚/未定
    for i, (cut, _) in enumerate(CUTS):
        if dt <= cut:
            return i
    return len(CUTS)

def get_val(ws, r):
    """Q（未支付金额）优先，Q空白用 H（待履行金额）兜底"""
    Q = ws.cell(row=r, column=17).value
    if Q is not None and str(Q).strip() != '':
        try:
            return float(Q)
        except:
            pass
    H = ws.cell(row=r, column=8).value
    if H is not None and str(H).strip() != '':
        try:
            return float(H)
        except:
            return 0.0
    return 0.0

wb = openpyxl.load_workbook(SRC, data_only=True)
ws = wb['Sheet1']
max_row = ws.max_row

# 表头确认（列结构可能变）
assert '未支付金额' in str(ws.cell(row=2, column=17).value or ''), f'Q列表头异常: {ws.cell(row=2, column=17).value}'
assert '待履行金额' in str(ws.cell(row=2, column=8).value or ''), f'H列表头异常: {ws.cell(row=2, column=8).value}'

tot = {'局': [0]*(len(CUTS)+1), '公司': [0]*(len(CUTS)+1)}
cur = None
for r in range(3, max_row + 1):
    A = str(ws.cell(row=r, column=1).value or '').strip()
    B = str(ws.cell(row=r, column=2).value or '').strip()
    C = str(ws.cell(row=r, column=3).value or '').strip()
    D = str(ws.cell(row=r, column=4).value or '').strip()
    G = str(ws.cell(row=r, column=7).value or '').strip()
    # 案件识别：三列同非空 或 (A非空且B空且C='/')
    if A and B and C:
        cur = {'D': D}
    elif A and not B and C == '/':
        cur = {'D': D or (cur['D'] if cur else '局')}
    if cur is None:
        continue
    # 孤立金额行（A/B/C/G 全空、仅 Q 有值，无案件归属）→ 排除不统计（如 r144）
    if not A and not B and not C and not G:
        continue
    D2 = '公司' if cur['D'] == '公司' else '局'
    v = get_val(ws, r)
    if v == 0:
        continue
    b = bucket(edate(ws.cell(row=r, column=9).value))
    tot[D2][b] += v

print(f'底稿: {SRC.split("/")[-1]}')
print(f'{"名义":<6}', '  '.join(f'{lbl}={tot[k][i]:>14,.2f}' for k in ('局', '公司') for i, (_, lbl) in enumerate(CUTS)))
for D, lbl in (('局', '上级名义'), ('公司', '非上级名义')):
    parts = []
    for i, (_, name) in enumerate(CUTS):
        parts.append(f'{name}: {tot[D][i]:,.2f}')
    parts.append(f'更晚/未定: {tot[D][len(CUTS)]:,.2f}')
    parts.append(f'合计: {sum(tot[D]):,.2f}')
    print(f'{lbl}: ' + ' | '.join(parts))
