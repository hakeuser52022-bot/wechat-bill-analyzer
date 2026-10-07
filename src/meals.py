# -*- coding: utf-8 -*-
"""
吃饭预算：逐餐拆分 + 下月预算分档 + 娱乐活动预算 + 总预算。

用法:
    python meals.py <账单.xlsx> [--from YYYY-MM-DD] [--to YYYY-MM-DD]
                    [--days 31] [--json 预算.json]

核心方法（重要）:
    1. 用"稳定期"锚点单价（食堂早/午/晚笔均），不要用含旅行/聚餐的月份
    2. 预算 = 食堂三餐锚点×天数 + 可选加料（奶茶/加餐/外卖）按"每周次数"折算
    3. 一次性支出（火车票、办证、教材）单列，不能摊进月预算，否则月预算虚高
    4. 娱乐活动单独留预算：偶发大额（演出票）建议"娱乐基金"按月存钱
"""
import sys, io, os, json, argparse, collections
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

# 餐饮口径统一在 analyze.py 里定义（唯一权威），这里直接复用，
# 避免与 budget.py 出现两套判断。
from analyze import load_bill, classify, compile_rules, _force_utf8_stdio, chow_kind, net_expenses, window_days  # noqa: E402

_force_utf8_stdio()

# 说明：NON_FOOD / CONV / CAFE / OUTFOOD 与 chow_kind() 已上移到 analyze.py，
# 作为全项目唯一的餐饮口径。这里不再保留副本——两套判断是问题 5 的同类错误。


def analyze_meals(rows, lo=None, hi=None, days=None):
    exp = net_expenses(rows, lo, hi)
    rules = compile_rules()
    for r in exp:
        r.setdefault('类别', classify(r, rules))
    food = []
    for r in exp:
        k = chow_kind(r)
        if k:
            r['渠道'], r['餐次'] = k
            food.append(r)
    if days is None: days = window_days([r for r in rows if (not lo or r['iso'] >= lo) and (not hi or r['iso'] < hi)], lo, hi)

    def agg(key):
        d = collections.defaultdict(lambda: {'笔数': 0, '金额': 0.0})
        for r in food:
            d[r[key]]['笔数'] += 1
            d[r[key]]['金额'] = round(d[r[key]]['金额'] + r['金额'], 2)
        return dict(d)

    by_meal, by_ch = agg('餐次'), agg('渠道')

    def avg(ch=None, meal=None):
        sub = [r for r in food
               if (ch is None or r['渠道'] == ch) and (meal is None or r['餐次'] == meal)]
        return (round(sum(r['金额'] for r in sub) / len(sub), 2) if sub else 0), len(sub)

    anchor = {}
    for nm, ml in (('早餐', '早餐'), ('午餐', '午餐'), ('晚餐', '晚餐')):
        anchor[nm] = avg(ch='食堂', meal=ml)
    unit = {
        '奶茶咖啡': avg(ch='奶茶咖啡'),
        '加餐夜宵': avg(meal='加餐/夜宵'),
        '外卖正餐': avg(ch='校外外卖', meal='午餐'),
    }
    total = round(sum(r['金额'] for r in food), 2)
    # "现状档"按整月投影：窗口不足一天数时，用窗口日均外推，避免与预算档位口径不一致
    projected = round(total / days * 30 if days else 0, 2)
    return {'days': days, 'food': food, 'total': total, 'projected_month': projected,
            'by_meal': by_meal, 'by_channel': by_ch,
            'anchor': anchor, 'unit': unit, 'daily': round(total / days, 2) if days else 0}


def budget_plans(m, base_days=31):
    """按"每周次数"生成吃饭预算档位"""
    if base_days <= 0: raise ValueError('预算天数必须大于零')
    a = m['anchor']
    base = round(a['早餐'][0] + a['午餐'][0] + a['晚餐'][0], 2)
    u_tea = m['unit']['奶茶咖啡'][0]
    u_snack = m['unit']['加餐夜宵'][0]
    u_out = m['unit']['外卖正餐'][0]
    plans = []
    for name, tw, sw, ow, note in [
        ('① 极简档', 0, 0, 0, '三餐全食堂；不喝奶茶、不吃夜宵'),
        ('② 经济档', 1, 2, 1, '每周1杯奶茶、2次加餐、1次外卖'),
        ('③ 推荐档', 1, 3, 2, '每周1杯奶茶、3次加餐、2次外卖（可持续）'),
        ('④ 宽裕档', 2, 5, 3, '每周2杯奶茶、5次加餐、3次外卖'),
        ('⑤ 舒适档', 3, 8, 5, '每周3杯奶茶、8次加餐、5次外卖'),
    ]:
        missing = [k for k,v in a.items() if v[1] == 0]
        missing += [k for k,n in [('奶茶咖啡',tw),('加餐夜宵',sw),('外卖正餐',ow)] if n and not m['unit'][k][1]]
        b = round(base * base_days - ow * a['午餐'][0] * base_days / 7, 2)
        t = round(tw * u_tea * base_days / 7, 2)
        s = round(sw * u_snack * base_days / 7, 2)
        o = round(ow * u_out * base_days / 7, 2)
        plans.append({'name': name, 'base': b, 'tea': t, 'snack': s, 'out': o,
                      'known_subtotal': round(b + t + s + o, 2), 'total': None if missing else round(b + t + s + o, 2), 'note': note,
                      'days': base_days, 'missing': missing, 'available': not missing,
                      'daily': None if missing else round((b + t + s + o) / base_days, 2)})
    return base, plans, {'奶茶咖啡': u_tea, '加餐夜宵': u_snack, '外卖正餐': u_out}


def print_meals(m, base, plans):
    print("=" * 80)
    print(f"吃饭分析（{m['days']} 天）  合计 {m['total']:.2f} 元  日均 {m['daily']:.2f}")
    print(f"\n{'餐次':<12}{'笔数':>5}{'金额':>10}{'笔均':>9}{'日均':>9}{'月化':>10}")
    for k in ('早餐', '午餐', '晚餐', '加餐/夜宵'):
        v = m['by_meal'].get(k)
        if not v:
            continue
        print(f"{k:<12}{v['笔数']:>5}{v['金额']:>10.2f}{v['金额']/v['笔数']:>9.2f}"
              f"{v['金额']/m['days']:>9.2f}{v['金额']/m['days']*30:>10.2f}")
    print(f"\n{'渠道':<12}{'笔数':>5}{'金额':>10}{'笔均':>9}")
    for k, v in sorted(m['by_channel'].items(), key=lambda kv: -kv[1]['金额']):
        print(f"{k:<12}{v['笔数']:>5}{v['金额']:>10.2f}{v['金额']/v['笔数']:>9.2f}")
    print(f"\n锚点单价：食堂早 {m['anchor']['早餐'][0]:.2f} / 午 {m['anchor']['午餐'][0]:.2f}"
          f" / 晚 {m['anchor']['晚餐'][0]:.2f} = {base:.2f} 元/天")
    print(f"          奶茶 {m['unit']['奶茶咖啡'][0]:.2f}  加餐 {m['unit']['加餐夜宵'][0]:.2f}"
          f"  外卖正餐 {m['unit']['外卖正餐'][0]:.2f}")
    print(f"\n下月吃饭预算（{plans[0]['days']} 天；外卖替代午餐）：")
    print(f"{'方案':<12}{'食堂三餐':>11}{'奶茶':>9}{'加餐':>9}{'外卖':>9}{'合计':>11}{'日均':>8}")
    for p in plans:
        if not p['available']:
            print(p['name'] + '：样本不足，暂不提供总预算；缺少 ' + '、'.join(p['missing']))
            continue
        print(f"{p['name']:<12}{p['base']:>11.2f}{p['tea']:>9.2f}{p['snack']:>9.2f}"
              f"{p['out']:>9.2f}{p['total']:>11.2f}{p['daily']:>8.2f}")
    m = {**m, 'projected_month': round(m['total'] / m['days'] * plans[0]['days'],2) if m['days'] else 0}
    print(f"{'⑥ 现状档':<12}{'—':>11}{'—':>9}{'—':>9}{'—':>9}"
          f"{m['projected_month']:>11.2f}{m['daily']:>8.2f}")
    if m['days'] != plans[0]['days']:
        print(f"   注：现状档 = 窗口 {m['days']} 天实际餐饮 {m['total']:.2f} 元，"
              f"按月投影（日均 {m['daily']:.2f} × {plans[0]['days']}）得到 {m['projected_month']:.2f} 元；"
              f"若窗口明显短于一个月，该值参考意义有限。")


def main():
    ap = argparse.ArgumentParser(description='吃饭预算')
    ap.add_argument('bill')
    ap.add_argument('--from', dest='lo')
    ap.add_argument('--to', dest='hi')
    ap.add_argument('--days', type=int, default=31, help='预算天数（默认31）')
    ap.add_argument('--json', help='输出 JSON')
    a = ap.parse_args()
    rows, _, _ = load_bill(a.bill)
    m = analyze_meals(rows, lo=a.lo, hi=a.hi)
    base, plans, unit = budget_plans(m, a.days)
    print_meals(m, base, plans)
    if a.json:
        with open(a.json, 'w', encoding='utf-8') as f:
            json.dump({'summary': {k: v for k, v in m.items() if k != 'food'},
                       'base_daily': base, 'plans': plans, 'unit': unit},
                      f, ensure_ascii=False, indent=1, default=str)
        print(f"\n已写出：{a.json}")


if __name__ == '__main__':
    main()
