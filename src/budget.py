# -*- coding: utf-8 -*-
"""
预算与省钱方案：把吃饭预算、其他支出、娱乐活动合成下月总预算。

用法:
    python budget.py <账单.xlsx> [--from YYYY-MM-DD] [--to YYYY-MM-DD]
                     [--month-days 31] [--entertainment 180] [--json 预算.json]

输出:
    1. 经常性 vs 一次性支出拆解（月预算只应用经常性部分外推）
    2. 经常性支出结构（月化）
    3. 各档吃饭预算（来自 meals.py 的方法）
    4. 娱乐活动预算分档
    5. 下月总预算与情形对比

核心方法:
    - 月预算 = 经常性支出日均 × 天数；一次性支出（火车票/办证/教材）不计入
    - 省钱建议要按"可压缩金额"排序，而不是按金额大小排序
    - 娱乐活动必须留预算；偶发大额建议"娱乐基金"按月存
"""
import sys, io, os, json, argparse, collections, datetime
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from analyze import load_bill, classify, compile_rules, _force_utf8_stdio, net_expenses, window_days
from meals import analyze_meals, budget_plans

_force_utf8_stdio()

# 一次性支出识别（这些不该摊进月预算）
ONE_OFF = {
    'train': ('中铁网络', '12306'),
    'card_open': ('岭南通', '交通卡', '开卡'),
    'passport': ('出入境', '通行证', '护照'),
    'textbook': ('教材', '书店', '外研社'),
    'medical': ('医疗器械', '药业', '药店', '医院'),
    'uniform': ('军训',),
}
ONE_OFF_LABEL = {
    'train': '火车票/长途交通', 'card_open': '交通卡开卡充值（预付）',
    'passport': '办证', 'textbook': '教材', 'ai_sub': '订阅/API',
    'medical': '看病买药', 'uniform': '军训物资',
}
# 可压缩项：key 为对内键（food 子类或 ①-⑫ 类别前缀），label 为展示名
CUT = [
    ('网约车', '网约车', 0.60, '除赶时间/带重物/雨天，改地铁公交；提前 15 分钟出门'),
    ('网购消费', '网购消费', 0.30, '非急需先放购物车 24 小时；按清单买，不刷推荐流'),
    ('奶茶咖啡', '奶茶咖啡', 0.50, '每周≤2 杯，用券；不为凑起送费多加单'),
    ('便利店', '便利店/零食', 0.40, '自带水杯；零食每周集中买一次'),
    ('外卖正餐', '外卖正餐', 0.30, '外卖减到每周 2 次，午晚餐尽量食堂'),
    ('转账与红包', '转账与红包', 0.30, 'AA 及时结清，减少反复小额往来'),
]


def is_oneoff(r):
    t = str(r['交易对方']) + str(r['商品']) + str(r['类别'])
    for k, kws in ONE_OFF.items():
        if any(w in t for w in kws):
            return k
    return None


# 网约车识别（交通里最可压缩的一项）
RIDE = ('如祺', '阳光出行', 'T3出行', '享道', '首汽', '哈啰', '滴滴', '曹操', '花小猪')
# 网购识别
SHOP = ('淘宝', '天猫', '京东', '拼多多', '抖音电商', '闲鱼', '唯品会')


def analyze_budget(rows, lo=None, hi=None, month_days=31, entertainment=180,
                   rules=None, family=None):
    from analyze import chow_kind          # 唯一餐饮口径，见 analyze.py 顶部说明
    rules = compile_rules(rules)
    if month_days <= 0: raise ValueError('预算天数必须大于零')
    exp = net_expenses(rows, lo, hi)
    for r in exp:
        r.setdefault('类别', classify(r, rules))
        k = chow_kind(r)
        if k:
            r['渠道'], r['餐次'] = k
    # 没有任何支出时直接返回零值结构，避免 min()/max() 对空序列抛 ValueError
    if not exp:
        return {'days': 0, 'month_days': month_days, 'total': 0.0,
                'recurring': 0.0, 'one_off': 0.0, 'recurring_daily': 0.0,
                'recurring_month': 0.0, 'food_recurring': 0.0, 'food_month': 0.0,
                'other_recurring': 0.0, 'other_month': 0.0,
                'one_off_detail': {}, 'recurring_by_cat': {}, 'food_by_channel': {},
                'cuts': [], 'total_save': 0.0, 'after_save': 0.0,
                'entertainment': entertainment, 'empty': True}
    d0 = min(r['dt'].date() for r in exp)
    d1 = max(r['dt'].date() for r in exp)
    days = window_days([r for r in rows if (not lo or r['iso'] >= lo) and (not hi or r['iso'] < hi)], lo, hi)

    rec, once = [], collections.defaultdict(float)
    for r in exp:
        k = is_oneoff(r)
        if k:
            once[k] = round(once[k] + r['金额'], 2)
        else:
            rec.append(r)
    rec_tot = round(sum(r['金额'] for r in rec), 2)
    once_tot = round(sum(once.values()), 2)
    tot = round(rec_tot + once_tot, 2)

    # 经常性支出：餐饮（含子渠道）与非餐饮分开
    rag = collections.defaultdict(lambda: {'笔数': 0, '金额': 0.0})
    ch_amt = collections.defaultdict(float)
    ride_amt = shop_amt = 0.0
    for r in rec:
        rag[r['类别']]['笔数'] += 1
        rag[r['类别']]['金额'] = round(rag[r['类别']]['金额'] + r['金额'], 2)
        # 餐饮判定完全依据 analyze.chow_kind（唯一口径）。
        # 这里曾经额外排除"类别以 ⑫ 开头"的记录，导致与 meals.py 对同一批数据
        # 给出不同答案（差 572.60 元）——两套口径是问题 5 的同类错误，已删除。
        if r.get('渠道'):
            ch_amt[r['渠道']] = round(ch_amt[r['渠道']] + r['金额'], 2)
        else:
            blob = str(r['交易对方']) + str(r['商品'])
            if any(k in blob for k in RIDE) and r['类别'].startswith('②'):
                ride_amt = round(ride_amt + r['金额'], 2)
            if any(k in blob for k in SHOP):
                shop_amt = round(shop_amt + r['金额'], 2)
    rag = dict(sorted(rag.items(), key=lambda kv: -kv[1]['金额']))
    food_amt = round(sum(ch_amt.values()), 2)

    # 非餐饮经常性支出（这才是"其他支出"，与吃饭预算不重叠）
    other_rec = round(rec_tot - food_amt, 2)
    mo = lambda v: round(v / days * month_days, 2)

    # 可压缩项金额（按对内键）
    def cut_amount(key):
        if key == '网约车':
            return ride_amt
        if key == '网购消费':
            return shop_amt
        if key in ('奶茶咖啡', '便利店', '外卖正餐'):
            mapping = {'奶茶咖啡': '奶茶咖啡', '便利店': '便利店', '外卖正餐': '校外外卖'}
            return ch_amt.get(mapping[key], 0.0)
        if key == '转账与红包':
            return rag.get('⑦ 转账与红包', {}).get('金额', 0.0)
        return 0.0

    cuts = []
    for key, label, rate, how in CUT:
        amt = cut_amount(key)
        if amt <= 0:
            continue
        monthly = mo(amt)
        save = round(monthly * rate, 2)
        cuts.append({'键': key, '类别': label, '窗口金额': round(amt, 2), '月化': monthly,
                     '压缩幅度': rate, '月省': save, '压缩后': round(monthly - save, 2),
                     '做法': how})
    cuts.sort(key=lambda x: -x['月省'])
    total_save = round(sum(c['月省'] for c in cuts), 2)

    return {'days': days, 'month_days': month_days, 'total': tot,
            'recurring': rec_tot, 'one_off': once_tot,
            'recurring_daily': round(rec_tot / days, 2),
            'recurring_month': mo(rec_tot),
            'food_recurring': round(food_amt, 2),
            'food_month': mo(food_amt),
            'other_recurring': other_rec,
            'other_month': mo(other_rec),
            'one_off_detail': dict(sorted(once.items(), key=lambda kv: -kv[1])),
            'recurring_by_cat': rag, 'food_by_channel': dict(ch_amt),
            'cuts': cuts, 'total_save': total_save,
            'after_save': round(mo(rec_tot) - total_save, 2),
            'entertainment': entertainment}


def print_budget(b, m, base, plans):
    if not b['total']:
        print('本区间无净支出；缺少预算依据。')
        return
    print("=" * 82)
    print(f"窗口 {b['days']} 天  净支出 {b['total']:.2f}")
    print(f"  经常性 {b['recurring']:.2f}（{b['recurring']/b['total']*100:.0f}%）"
          f"  日均 {b['recurring_daily']:.2f}  月化({b['month_days']}天) {b['recurring_month']:.2f}")
    print(f"     其中餐饮   {b['food_recurring']:.2f}  月化 {b['food_month']:.2f}")
    print(f"     非餐饮     {b['other_recurring']:.2f}  月化 {b['other_month']:.2f}")
    print(f"  一次性 {b['one_off']:.2f}（{b['one_off']/b['total']*100:.0f}%）—— 不计入月预算")
    if b['one_off_detail']:
        print("  一次性明细：")
        for k, v in b['one_off_detail'].items():
            print(f"     {ONE_OFF_LABEL.get(k, k):<18}{v:>9.2f}")

    print(f"\n经常性支出结构（月化 {b['month_days']} 天）：")
    print(f"{'类别':<16}{'窗口金额':>10}{'日均':>9}{'月化':>10}{'笔数':>6}")
    for k, v in b['recurring_by_cat'].items():
        print(f"{k:<16}{v['金额']:>10.2f}{v['金额']/b['days']:>9.2f}"
              f"{v['金额']/b['days']*b['month_days']:>10.2f}{v['笔数']:>6}")
    print(f"{'合计':<16}{b['recurring']:>10.2f}{b['recurring_daily']:>9.2f}"
          f"{b['recurring_month']:>10.2f}")
    if b['food_by_channel']:
        print("  餐饮内部渠道：")
        for k, v in sorted(b['food_by_channel'].items(), key=lambda kv: -kv[1]):
            print(f"     {k:<12}{v:>10.2f}  月化 {v/b['days']*b['month_days']:>9.2f}")

    print(f"\n省钱方案（按可省金额排序，合计可省 {b['total_save']:.2f} 元/月）：")
    if not b['cuts']:
        print("   （本窗口未识别到可压缩项）")
    else:
        print(f"{'类别':<14}{'月化':>10}{'幅度':>8}{'月省':>10}  做法")
        for c in b['cuts']:
            print(f"{c['类别']:<14}{c['月化']:>10.2f}{c['压缩幅度']*100:>7.0f}%"
                  f"{c['月省']:>10.2f}  {c['做法']}")
    print(f"\n省后经常性月支出：{b['after_save']:.2f}")

    if any(not p['available'] for p in plans):
        print('餐饮样本不足，暂不生成总预算；请补充三餐与可选项目的单价样本。')
        return
    rec_plan = [p for p in plans if '推荐' in p['name']][0]
    econ_plan = [p for p in plans if '经济' in p['name']][0]
    mini_plan = [p for p in plans if '极简' in p['name']][0]
    wide_plan = [p for p in plans if '宽裕' in p['name']][0]
    # 非餐饮经常性支出（省钱后）
    other_save = sum(c['月省'] for c in b['cuts'] if c['键'] not in ('奶茶咖啡','便利店','外卖正餐'))
    other_net = round(b['other_month'] - other_save, 2)
    ent = b['entertainment']
    print(f"\n下月吃饭预算：")
    print(f"{'方案':<12}{'食堂三餐':>11}{'奶茶':>9}{'加餐':>9}{'外卖':>9}{'合计':>11}{'日均':>8}")
    for p in plans:
        print(f"{p['name']:<12}{p['base']:>11.2f}{p['tea']:>9.2f}{p['snack']:>9.2f}"
              f"{p['out']:>9.2f}{p['total']:>11.2f}{p['daily']:>8.2f}")

    print(f"\n下月总预算（吃饭 + 其他经常性 + 娱乐）：")
    print(f"   其中「其他经常性」= 非餐饮月化 {b['other_month']:.2f} − 非餐饮节省 {other_save:.2f}"
          f" = {other_net:.2f}")
    print(f"\n{'情形':<10}{'吃饭':>10}{'其他':>10}{'娱乐':>10}{'月总计':>11}")
    for nm, meal, e in [('极简', mini_plan['total'], 0),
                        ('推荐', rec_plan['total'], ent),
                        ('宽裕', wide_plan['total'], 300),
                        ('含演出', rec_plan['total'], 450),
                        ('底线', econ_plan['total'], 80)]:
        print(f"{nm:<10}{meal:>10.2f}{other_net:>10.2f}{e:>10.2f}{meal+other_net+e:>11.2f}")
    print(f"\n   注：吃饭预算来自 meals.py 的档位；其他经常性支出已扣除省钱方案的压缩额。")


def main():
    ap = argparse.ArgumentParser(description='下月预算与省钱方案')
    ap.add_argument('bill')
    ap.add_argument('--from', dest='lo')
    ap.add_argument('--to', dest='hi')
    ap.add_argument('--month-days', type=int, default=31)
    ap.add_argument('--entertainment', type=float, default=180)
    ap.add_argument('--json')
    a = ap.parse_args()
    rows, _, _ = load_bill(a.bill)
    b = analyze_budget(rows, a.lo, a.hi, a.month_days, a.entertainment)
    m = analyze_meals(rows, lo=a.lo, hi=a.hi)
    base, plans, unit = budget_plans(m, a.month_days)
    print_budget(b, m, base, plans)
    if a.json:
        with open(a.json, 'w', encoding='utf-8') as f:
            json.dump({'budget': b, 'meal_base_daily': base, 'meal_plans': plans},
                      f, ensure_ascii=False, indent=1, default=str)
        print(f"\n已写出：{a.json}")


if __name__ == '__main__':
    main()
