# -*- coding: utf-8 -*-
"""
餐饮口径一致性核对。

历史上 meals.py 与 budget.py 各有一份"什么算餐饮"的判断，
对同一批数据给出不同答案（差 572.60 元）。现在两者都只用
analyze.chow_kind，这个脚本把一致性变成可随时验证的输出。

用法:
    python tools/check_food_consistency.py [账单.xlsx]

不带参数时使用 examples/data/demo_bill.xlsx。
退出码：0 = 一致，1 = 存在分歧。
"""
import argparse
import os
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.abspath(os.path.join(HERE, '..'))
sys.path.insert(0, os.path.join(ROOT, 'src'))

try:
    sys.stdout.reconfigure(encoding='utf-8', errors='replace')
except Exception:
    pass

import analyze      # noqa: E402
import meals        # noqa: E402
import budget       # noqa: E402


def main():
    ap = argparse.ArgumentParser(description='核对 meals 与 budget 的餐饮口径是否一致')
    ap.add_argument('bill', nargs='?',
                    default=os.path.join(ROOT, 'examples', 'data', 'demo_bill.xlsx'),
                    help='账单 xlsx（默认使用演示数据）')
    a = ap.parse_args()
    bill = a.bill
    if not os.path.exists(bill):
        raise SystemExit(f'文件不存在：{bill}')

    rows, _, _ = analyze.load_bill(bill)
    rules = analyze.compile_rules()
    for r in rows:
        r.setdefault('类别', analyze.classify(r, rules))

    m = meals.analyze_meals(rows)
    b = budget.analyze_budget(rows)
    meals_food = round(sum(v['金额'] for v in m['by_channel'].values()), 2)

    print(f'账单：{os.path.basename(bill)}（{len(rows)} 条记录）')
    print()
    print('餐饮渠道明细：')
    for k, v in sorted(m['by_channel'].items(), key=lambda kv: -kv[1]['金额']):
        print(f'   {k:<10}{v["金额"]:>10.2f}   {v["笔数"]:>3} 笔')
    print(f'   {"meals 合计":<10}{meals_food:>10.2f}')
    print(f'   {"budget 合计":<10}{b["food_recurring"]:>10.2f}')
    print()

    diff = round(b['food_recurring'] - meals_food, 2)
    same_set = abs(diff) < 0.005  # 金额按统一净额口径核对；模块使用副本，不依赖对象身份
    print(f'餐饮金额差异：{diff}')
    print(f'餐饮记录集合一致：{"是" if same_set else "否"}')

    lhs = round(b['food_recurring'] + b['other_recurring'], 2)
    ok_identity = abs(lhs - b['recurring']) < 0.005
    print(f'恒等式 餐饮 + 非餐饮 = 经常性支出：{b["food_recurring"]} + '
          f'{b["other_recurring"]} = {lhs}（经常性 {b["recurring"]}）'
          f'  {"成立" if ok_identity else "不成立"}')

    ok = abs(diff) < 0.005 and same_set and ok_identity
    print()
    print('结论：' + ('两个模块的餐饮口径一致 ✅' if ok else '存在口径分歧 ❌'))
    return 0 if ok else 1


if __name__ == '__main__':
    sys.exit(main())
