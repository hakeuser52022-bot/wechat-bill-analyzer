# -*- coding: utf-8 -*-
"""检查文档引用的数字与真实输出是否一致，并对生成的 Excel 报告做公式求值。

这个脚本补上了早期验证的一个缺口：
之前只从 docs/demo.md 抓数字比对文本，**完全没有校验 Excel 里的公式列**，
结果 timeseg.py 占比分母少一个 +1 的缺陷一直没被发现——
因为它不报错，docx/文本层面也看不出来，只有把那一列加起来才暴露。

用法: python tools/verify_numbers.py
"""
import os
import re
import sys

try:
    sys.stdout.reconfigure(encoding='utf-8', errors='replace')
except Exception:
    pass

ROOT = os.path.abspath(os.path.join(os.path.dirname(os.path.abspath(__file__)), '..'))

# ---------------------------------------------------------------------------
# 一、Excel 公式校验
# ---------------------------------------------------------------------------
def eval_cell(ws, coord, cache=None):
    """对 xlsx 单元格做真求值（支持 SUM 与四则运算）。"""
    if cache is None:
        cache = {}
    if coord in cache:
        return cache[coord]
    v = ws[coord].value
    if isinstance(v, (int, float)):
        cache[coord] = v
        return v
    if not (isinstance(v, str) and v.startswith('=')):
        cache[coord] = v
        return v
    body = v[1:]
    cache[coord] = 0
    m = re.fullmatch(r'SUM\(([A-Z]+\d+):([A-Z]+\d+)\)', body)
    if m:
        c0, r0 = re.match(r'([A-Z]+)(\d+)', m.group(1)).groups()
        c1, r1 = re.match(r'([A-Z]+)(\d+)', m.group(2)).groups()
        tot = 0.0
        for rr in range(int(r0), int(r1) + 1):
            for cc in range(ord(c0) - 65, ord(c1) - 65 + 1):
                x = eval_cell(ws, f'{chr(65 + cc)}{rr}', cache)
                if isinstance(x, (int, float)):
                    tot += x
        cache[coord] = tot
        return tot
    expr = re.sub(r'\$?([A-Z]{1,3}\$?\d+)',
                  lambda mm: str(eval_cell(ws, mm.group(1).replace('$', ''), cache)), body)
    try:
        if not re.fullmatch(r'[0-9eE+*/(). \-]+', expr): raise ValueError('不支持的公式')
        r = eval(expr, {'__builtins__': {}}, {})
    except Exception:
        r = None
    cache[coord] = r
    return r


def check_excel_percentages():
    """遍历 examples/output 下的 xlsx，校验每个"占比"列的合计是否为 100%。"""
    from openpyxl import load_workbook
    outdir = os.path.join(ROOT, 'examples', 'output')
    bad = 0
    checked = 0
    for fn in sorted(os.listdir(outdir)):
        if not fn.lower().endswith('.xlsx'):
            continue
        path = os.path.join(outdir, fn)
        wb = load_workbook(path)
        for ws in wb.worksheets:
            # 找出表头里含"占比"的列
            pct_cols = {}
            hdr_row = None
            for r in range(1, min(ws.max_row, 15) + 1):
                for c in range(1, ws.max_column + 1):
                    v = ws.cell(row=r, column=c).value
                    if isinstance(v, str) and '占比' in v:
                        pct_cols[c] = v
                        hdr_row = r
                if pct_cols:
                    break
            if not pct_cols:
                continue
            first = hdr_row + 1
            last = first
            while ws.cell(row=last, column=1).value not in (None, '合计'):
                last += 1
            if last <= first:
                continue
            cache = {}
            for c, name in pct_cols.items():
                s = 0.0
                for r in range(first, last):
                    x = eval_cell(ws, f'{chr(64 + c)}{r}', cache)
                    if isinstance(x, (int, float)):
                        s += x
                checked += 1
                ok = abs(s - 1) < 1e-6
                mark = 'OK  ' if ok else 'FAIL'
                print(f'  [{mark}] {fn} / {ws.title} / {name} '
                      f'({chr(64 + c)}{first}:{chr(64 + c)}{last - 1}) = {s * 100:.2f}%')
                if not ok:
                    bad += 1
    print(f'\n  共校验 {checked} 个占比列，异常 {bad} 个')
    return bad


# ---------------------------------------------------------------------------
# 二、文档数字一致性
# ---------------------------------------------------------------------------
demo = open(os.path.join(ROOT, 'docs', 'demo.md'), encoding='utf-8').read()

# 文档中明确声明"来自作者真实账单、无法在仓库复现"的数字，属于已声明出处
DECLARED_REAL_BILL = set()


def check_doc_numbers():
    docs = {}
    for rel in ('README.md', 'docs/presentation.md', 'docs/dev-log.md'):
        p = os.path.join(ROOT, rel)
        if os.path.exists(p):
            docs[rel] = open(p, encoding='utf-8').read()
    warn = 0
    print('\n=== 文档引用的 >1000 金额是否都能在 demo.md 找到 ===')
    for d, t in docs.items():
        nums = set(re.findall(r'\b(\d{4}\.\d{2})\b', t))
        missing = sorted(n for n in nums
                         if n not in demo and n not in DECLARED_REAL_BILL)
        if missing:
            print(f'  [需核对] {d}: {missing}')
            warn += 1
        else:
            print(f'  [OK] {d}')
    return warn


def main():
    print('=' * 70)
    print('一、Excel 报告公式求值校验')
    print('=' * 70)
    bad_excel = check_excel_percentages()
    print('\n' + '=' * 70)
    print('二、文档数字与真实输出一致性')
    print('=' * 70)
    bad_docs = check_doc_numbers()
    print(f'\n历史金额白名单（应为空）: {sorted(DECLARED_REAL_BILL)}')
    print()
    if bad_excel or bad_docs:
        print(f'结果：Excel 异常 {bad_excel} 处，文档需核对 {bad_docs} 处')
        return 1
    print('结果：全部通过')
    return 0


if __name__ == '__main__':
    sys.exit(main())
