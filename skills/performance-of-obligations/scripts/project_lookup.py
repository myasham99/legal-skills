# -*- coding: utf-8 -*-
"""
项目编码查询（skill: performance-of-obligations 配套）
从 公司项目编码.xlsx 查项目，输出 code600 / code12 / 名称 / 分类 / 开工日期。

用法：
  python project_lookup.py <关键词>          # 按名称/简称/编码 模糊查询
  python project_lookup.py --code12 1220168774   # 按 12 开头旧码 精确查
  python project_lookup.py --code600 6000000155  # 按 600 开头新码 精确查
"""
import openpyxl, sys

sys.stdout.reconfigure(encoding='utf-8', errors='replace')

XMBM = r'D:/【用户名】/【诉讼】/案件统计/其他台账/公司项目编码.xlsx'

def load():
    wb = openpyxl.load_workbook(XMBM, data_only=True, read_only=True)
    ws = wb['Sheet1']
    rows = []
    for r in ws.iter_rows(min_row=2, values_only=True):
        if r[0] is None and r[3] is None:
            continue
        rows.append({
            'code600': str(r[0]).strip() if r[0] else '',
            'unit': str(r[1]).strip() if r[1] else '',
            'code12': str(r[2]).strip() if r[2] else '',
            'name': str(r[3]).strip() if r[3] else '',
            'region': str(r[4]).strip() if r[4] else '',
            'status': str(r[5]).strip() if r[5] else '',
            'cat': str(r[6]).strip() if r[6] else '',
            'start': str(r[8]).strip() if r[8] else '',
        })
    wb.close()
    return rows

def show(r):
    print(f'  code600={r["code600"]} | code12={r["code12"]} | {r["name"]}')
    print(f'    分类={r["cat"]} | 开工={r["start"]} | 状态={r["status"]} | 单位={r["unit"]}')

def main():
    rows = load()
    args = sys.argv[1:]
    if not args:
        print('用法: project_lookup.py <关键词|--code12 xxx|--code600 xxx>')
        return
    if args[0] == '--code12':
        key = args[1]
        hits = [r for r in rows if r['code12'] == key]
    elif args[0] == '--code600':
        key = args[1]
        hits = [r for r in rows if r['code600'] == key]
    else:
        key = args[0]
        hits = [r for r in rows if key in r['name'] or key in r['code12'] or key in r['code600']]
    print(f'匹配 {len(hits)} 条:')
    for r in hits[:20]:
        show(r)
    if len(hits) > 20:
        print(f'  ... 共 {len(hits)} 条，仅显示前 20')

if __name__ == '__main__':
    main()
