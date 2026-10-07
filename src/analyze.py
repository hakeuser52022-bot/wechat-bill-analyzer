# -*- coding: utf-8 -*-
"""
微信支付账单分析工具（skill: wechat-bill-analysis）

用法:
    python analyze.py <账单.xlsx> [--out 报告.xlsx] [--from YYYY-MM-DD] [--to YYYY-MM-DD]
                      [--json 输出.json] [--no-report] [--rules 规则.json]

功能:
    1. 自动识别微信账单表头（不假定固定行号）
    2. 逐笔归类到分类（内置一套通用规则，可用 --rules 覆盖）
    3. 正确剔除退款（负向支出），给出净支出
    4. 按可选时间区间分段统计
    5. 输出分类汇总、月度、收入构成、餐饮餐次细分
    6. 可选生成 Excel 报告

设计要点（踩过的坑，勿改）:
    - 退款必须从支出中剔除，否则总额虚高
    - "校园卡充值/交通卡开卡"等预付性支出会与后续消费重复，需单列提示
    - 扫码付款/二维码收款要按收款方关键词重新归到正确类别（如"羽毛球馆"）
    - 每个占比列的分母必须是本表合计行，且写完要校验各列合计=100%
"""
import sys, io, os, re, json, argparse, collections, datetime, math

def _force_utf8_stdio():
    """仅在需要时包装一次，避免被 import 时重复包装导致流被关闭。"""
    for name in ('stdout', 'stderr'):
        s = getattr(sys, name, None)
        if s is None or not hasattr(s, 'buffer'):
            continue
        enc = (getattr(s, 'encoding', '') or '').lower().replace('-', '')
        if enc == 'utf8':
            continue
        try:
            s.reconfigure(encoding='utf-8', errors='replace')
        except Exception:
            pass


_force_utf8_stdio()

# ----------------------------------------------------------------------------
# 一、解析：微信账单表头位置不固定，必须自动定位
# ----------------------------------------------------------------------------
# ---------------------------------------------------------------------------
# 餐饮口径（唯一权威定义）
#
# 这个判断曾经在 meals.py 和 budget.py 里各写了一份，结果两者对同一批数据
# 给出不同答案（差 572.60 元）：budget 额外排除了"类别属于 ⑫ 其他"的记录，
# meals 没有。两套判断 = 两套口径，是问题 5 的同类错误。
# 现在统一放在这里，meals 与 budget 都从这里取，保证恒等式成立：
#     餐饮 + 非餐饮 = 经常性支出
# ---------------------------------------------------------------------------
NON_FOOD = ('医疗器械', '药业', '药店', '医院', '诊所', '体检')
FOOD_CONV = ('便利', '小卖部', '售货机', '智慧零售', '超市', 'ELEVEN', '物联')
FOOD_CAFE = ('喜茶', '霸王茶姬', '茶百道', '瑞幸', '蜜雪', '星巴克', '咖啡', '奶茶',
             '糕点', '面包', '烘焙')
FOOD_OUT = ('美团', '饿了么', '麦当劳', '肯德基', '餐厅', '火锅', '烧烤', '快餐',
            '面馆', '粉店', '真功夫', '比萨')


def meal_slot(hour):
    """按小时判断餐次。"""
    if 5 <= hour < 10: return '早餐'
    if 10 <= hour < 14: return '午餐'
    if 14 <= hour < 19: return '晚餐'
    return '加餐/夜宵'


def chow_kind(row):
    """判断一笔支出属于哪类吃，返回 (渠道, 餐次)；不是餐饮返回 None。

    这是全项目唯一的餐饮判断入口。
    """
    t = str(row['交易对方']) + str(row['商品'])
    if any(k in t for k in NON_FOOD + ('充值', '校园卡', '打印', '教材', '团费', '洗衣', '热水', '饮水')):
        return None
    m = meal_slot(row['hour'])
    if '食堂' in t or '餐厅' in t and ('大学' in t or '学院' in t or '示例' in t):
        return ('食堂', m)
    if any(k in t for k in FOOD_CAFE):
        return ('奶茶咖啡', m)
    if any(k in t for k in FOOD_CONV):
        return ('便利店', m)
    if any(k in t for k in FOOD_OUT):
        return ('校外外卖', m)
    if str(row.get('类别') or '').startswith('①'):
        return ('其他餐饮', m)
    return None



def refund_amount(r):
    if r.get('收/支') != '支出': return 0.0
    status = str(r.get('当前状态') or '')
    if '退款' not in status: return 0.0
    if '全额退款' in status: return r['金额']
    match = re.search(r'已退款\s*[（(]?\s*[¥￥]?\s*(\d+(?:\.\d+)?)', status)
    if not match: raise ValueError(f'退款状态无法确定金额：{status}')
    amount = float(match.group(1))
    if not 0 <= amount <= r['金额']: raise ValueError('退款金额超出原交易金额')
    return round(amount, 2)

def net_expenses(rows, lo=None, hi=None):
    result=[]
    for r in rows:
        if r.get('收/支') != '支出' or (lo and r['iso'] < lo) or (hi and r['iso'] >= hi): continue
        amount=round(r['金额']-refund_amount(r),2)
        if amount > 0: result.append({**r, '原金额':r['金额'], '金额':amount})
    return result

def window_days(rows, lo=None, hi=None):
    start=datetime.date.fromisoformat(lo) if lo else min((r['dt'].date() for r in rows),default=None)
    end=datetime.date.fromisoformat(hi) if hi else max((r['dt'].date() for r in rows),default=None)
    if end is not None and not hi: end += datetime.timedelta(days=1)
    if start is None or end is None: return 0
    if end <= start: raise ValueError('结束日期必须晚于开始日期')
    return (end-start).days


def find_header_row(ws, max_scan=80):
    """找到包含 '交易时间' 与 '金额' 的表头行"""
    for i, row in enumerate(ws.iter_rows(min_row=1, max_row=max_scan, values_only=True), 1):
        cells = [str(c) for c in row if c is not None]
        if any('交易时间' in c for c in cells) and any('金额' in c for c in cells):
            return i
    raise ValueError("未找到微信账单表头行（应包含「交易时间」与「金额(元)」）")


def load_bill(path):
    from openpyxl import load_workbook
    wb = load_workbook(path, data_only=True)
    ws = wb[wb.sheetnames[0]]
    hr = find_header_row(ws)
    headers = [str(c.value).strip() if c.value is not None else '' for c in ws[hr]]
    required={'交易时间','金额(元)','收/支','当前状态'}
    if not required.issubset(headers): raise ValueError(f'缺少字段：{required-set(headers)}')
    rows = []
    for r in ws.iter_rows(min_row=hr + 1, values_only=True):
        if r[0] is None:
            continue
        d = dict(zip(headers, r))
        try:
            amt = float(d.get('金额(元)') or 0)
        except (TypeError, ValueError) as e:
            raise ValueError(f'无效金额：{d.get("金额(元)")}') from e
        if not math.isfinite(amt) or amt < 0: raise ValueError('交易金额必须为非负有限数值')
        d['金额'] = round(amt, 2)
        t = d.get('交易时间')
        if isinstance(t, str):
            try: t = datetime.datetime.fromisoformat(t.strip().replace('/', '-'))
            except ValueError as e: raise ValueError(f'无效交易日期：{t}') from e
        if not isinstance(t, datetime.datetime): raise ValueError(f'无效交易日期：{t}')
        d['dt'] = t
        d['iso'] = t.date().isoformat()
        d['hour'] = t.hour
        d['已退款'] = '退款' in str(d.get('当前状态') or '')
        rows.append(d)
    return rows, headers, hr


# ----------------------------------------------------------------------------
# 二、分类：关键词规则，顺序敏感（先具体后宽泛）
# ----------------------------------------------------------------------------
DEFAULT_RULES = [
    ('⑦ 转账与红包', ['微信红包', '群收款', '企业微信红包', '发给', '转账备注']),
    ('② 交通出行', ['中铁网络', '12306', '如祺出行', '阳光出行', 'T3出行', '享道出行',
                    '首汽约车', '哈啰出行', '滴滴出行', '羊城通', '岭南通', '地铁',
                    '轨道交通', '如约出行', '联合电子', '停车', '加油站', '修电动车']),
    ('⑥ 运动健身', ['羽毛球', '游泳', '健身房', '球馆', '示例运动', '体育']),
    ('④ 学习与订阅', ['校园卡', '信息中心', '教材', '书店', '打印', '图文', '军训',
                      'API', '会员', '数据库', '网课', '培训', '团费', '考试']),
    ('⑤ 通讯与网络', ['电信', '移动', '联通', '营业厅', '话费', '宽带', '流量']),
    ('⑧ 旅游与门票', ['景区', '门票', '博物院', '博物馆', '旅游', '文旅', '出入境',
                      '通行证', '大麦', '演出', '赛事']),
    ('③ 购物消费', ['淘宝', '天猫', '京东', '拼多多', '抖音电商', '闲鱼', '唯品会',
                    '品牌文化店', '百货', '网购']),
    ('① 餐饮美食', ['大学', '学院', '食堂', '餐厅', '美团', '饿了么', '麦当劳',
                    '肯德基', '喜茶', '霸王茶姬', '茶百道', '瑞幸', '蜜雪', '星巴克',
                    '便利', '零食', '面包', '蛋糕', '水果', '酸奶', '小卖部', '售货机',
                    '智慧零售', '咖啡', '烧烤', '火锅', '快餐', '面馆', '粉店', '饭',
                    '食', '饮', '奶茶', '酒', '超市']),
    ('⑨ 日用与洗衣', ['洗衣机', '饮水', '热水', '物联', '日用']),
    ('⑩ 医疗健康', ['药', '医院', '诊所', '医疗', '体检', '口腔', '眼科']),
    ('⑪ 公益捐赠', ['慈善', '公益', '捐赠', '基金会']),
]


def compile_rules(rules=None):
    if rules is None:
        return [(c, list(kws)) for c, kws in DEFAULT_RULES]
    if isinstance(rules, dict):
        return [(c, list(kws)) for c, kws in rules.items()]
    return [(c, list(kws)) for c, kws in rules]


def load_rules(path):
    """读取自定义分类规则。

    用 utf-8-sig 而不是 utf-8：Windows 的 PowerShell `Set-Content -Encoding utf8`
    默认会写入 BOM，用 utf-8 读取会抛
    "Unexpected UTF-8 BOM"——中文用户几乎必然会踩到。
    """
    try:
        with open(path, encoding='utf-8-sig') as f:
            data = json.load(f)
    except json.JSONDecodeError as e:
        raise SystemExit(
            f"规则文件不是合法 JSON：{path}\n"
            f"  {e}\n"
            f"  期望格式（对象形式，键为类别名，值为关键词数组）：\n"
            f'  {{"① 餐饮美食": ["食堂", "美团"], "② 交通出行": ["地铁"]}}'
        ) from e
    except OSError as e:
        raise SystemExit(f'无法读取规则文件：{path}\n  {e}') from e
    if not isinstance(data, (dict, list)):
        raise SystemExit(f'规则文件顶层应为对象或数组，实际是 {type(data).__name__}')
    # 显式跳过以下划线开头的说明键（如 "_说明"），避免它们被当成一个类别
    if isinstance(data, dict):
        cleaned = {}
        for k, v in data.items():
            if isinstance(k, str) and k.startswith('_'):
                continue
            if not isinstance(v, list) or not all(isinstance(x, str) for x in v):
                raise SystemExit(
                    f'规则 "{k}" 的值必须是字符串数组，实际是 {type(v).__name__}。\n'
                    f'  期望：{{"① 餐饮美食": ["食堂", "美团"]}}'
                )
            cleaned[k] = v
        if not cleaned:
            raise SystemExit(f'规则文件里没有任何有效类别（是否只有以下划线开头的说明键？）：{path}')
        data = cleaned
    return compile_rules(data)


def classify(row, rules):
    text = ' '.join(str(row.get(k) or '') for k in ('交易类型', '交易对方', '商品'))
    for cat, kws in rules:
        for kw in kws:
            if kw and kw in text:
                return cat
    return '⑫ 其他'


# ----------------------------------------------------------------------------
# 三、统计
# ----------------------------------------------------------------------------
def summarize(rows, rules=None, lo=None, hi=None, family=None):
    """lo/hi 为 'YYYY-MM-DD'（含 lo，不含 hi）。
    family: 视为"亲人/资助"的对方关键词列表，用于收入拆分。"""
    rules = compile_rules(rules)
    sel = [r for r in rows if (lo is None or r['iso'] >= lo) and (hi is None or r['iso'] < hi)]
    exp = [r for r in sel if r.get('收/支') == '支出']
    inc = [r for r in sel if r.get('收/支') == '收入']
    other = [r for r in sel if r.get('收/支') not in ('支出', '收入')]
    net_rows = net_expenses(exp)
    window_days(sel, lo, hi)
    ref_rows = [r for r in exp if r['已退款']]
    for r in net_rows:
        r['类别'] = classify(r, rules)
    for r in sel:
        r.setdefault('类别', None)

    gross = round(sum(r['金额'] for r in exp), 2)
    refund = round(sum(refund_amount(r) for r in ref_rows), 2)
    net = round(sum(r['金额'] for r in net_rows), 2)
    inc_tot = round(sum(r['金额'] for r in inc if '退款' not in str(r.get('交易类型'))), 2)
    cash_income = round(sum(r['金额'] for r in inc), 2)

    cat = collections.defaultdict(lambda: {'笔数': 0, '金额': 0.0})
    for r in net_rows:
        cat[r['类别']]['笔数'] += 1
        cat[r['类别']]['金额'] = round(cat[r['类别']]['金额'] + r['金额'], 2)
    cat = dict(sorted(cat.items(), key=lambda kv: -kv[1]['金额']))

    # 自检：分类必须完整覆盖净支出
    assert abs(sum(v['金额'] for v in cat.values()) - net) < 0.005, "分类合计 != 净支出"
    assert sum(v['笔数'] for v in cat.values()) == len(net_rows), "分类笔数 != 净支出笔数"

    mon = collections.defaultdict(lambda: collections.defaultdict(float))
    mon_tot, mon_cnt = collections.defaultdict(float), collections.defaultdict(int)
    for r in net_rows:
        m = r['dt'].strftime('%Y-%m')
        mon[m][r['类别']] = round(mon[m][r['类别']] + r['金额'], 2)
        mon_tot[m] = round(mon_tot[m] + r['金额'], 2)
        mon_cnt[m] += 1

    # 收入构成（可拆出"亲人/资助"）
    fam = list(family or [])
    inc_by = collections.defaultdict(lambda: {'笔数': 0, '金额': 0.0})

    def inc_kind(r):
        t, cp = str(r['交易类型']), str(r['交易对方'])
        if '退款' in t:
            return '退款回流'
        if fam and any(f in cp for f in fam):
            return '亲人/资助转入'
        if '红包' in t:
            return '红包收入'
        return '其他转入'

    for r in inc:
        k = inc_kind(r)
        inc_by[k]['笔数'] += 1
        inc_by[k]['金额'] = round(inc_by[k]['金额'] + r['金额'], 2)
    inc_by = dict(sorted(inc_by.items(), key=lambda kv: -kv[1]['金额']))

    def meal(r):
        if 5 <= r['hour'] < 10: return '早餐'
        if 10 <= r['hour'] < 14: return '午餐'
        if 14 <= r['hour'] < 19: return '晚餐'
        return '加餐/夜宵'

    meals = collections.defaultdict(lambda: {'笔数': 0, '金额': 0.0})
    for r in net_rows:
        if str(r['类别']).startswith('①'):
            k = meal(r)
            meals[k]['笔数'] += 1
            meals[k]['金额'] = round(meals[k]['金额'] + r['金额'], 2)
    meals = {k: meals[k] for k in ('早餐', '午餐', '晚餐', '加餐/夜宵') if k in meals}

    days = window_days(sel, lo, hi)

    return {
        'range': {'lo': lo, 'hi': hi,
                  'first': min((r['iso'] for r in sel), default=None),
                  'last': max((r['iso'] for r in sel), default=None)},
        'days': days,
        'counts': {'total': len(sel), 'expense': len(exp), 'income': len(inc),
                   'other': len(other), 'net_expense': len(net_rows), 'refunded': len(ref_rows)},
        'gross': gross, 'refund': refund, 'net': net, 'income': inc_tot,
        # 结余 = 收入 − 净支出。注意不能用 gross（含退款的流水）：
        # income 排除退款回流；另提供 cash_balance 表示本期现金流差额。
        'balance': round(inc_tot - net, 2),
        'cash_balance': round(cash_income - gross, 2),
        'refund_income': round(cash_income-inc_tot, 2),
        'daily': round(net / days, 2) if days else 0,
        'categories': cat, 'monthly': {m: {'金额': mon_tot[m], '笔数': mon_cnt[m],
                                           'by_cat': dict(mon[m])} for m in sorted(mon_tot)},
        'income_breakdown': inc_by, 'meals': meals,
        'refund_rows': ref_rows, 'net_rows': net_rows, 'income_rows': inc,
        'other_rows': other, 'all_rows': sel,
    }


# ----------------------------------------------------------------------------
# 四、控制台报告
# ----------------------------------------------------------------------------
def print_report(res, top=15):
    rc = res['range']
    print("=" * 78)
    days_txt = f"{res['days']} 天" if res['days'] else "无支出记录"
    print(f"区间：{rc['first']} ~ {rc['last']}（{days_txt}）")
    c = res['counts']
    print(f"记录 {c['total']} 笔：支出 {c['expense']}（其中已退款 {c['refunded']}）"
          f"、收入 {c['income']}、不计收支 {c['other']}")
    print(f"支出流水 {res['gross']:.2f} | 退款 {res['refund']:.2f} | 净支出 {res['net']:.2f} "
          f"| 日均 {res['daily']:.2f}")
    print(f"收入（不含退款回流）{res['income']:.2f} | 消费口径结余 {res['balance']:.2f}")
    print(f"本期退款回流 {res['refund_income']:.2f} | 本期现金流差额 {res['cash_balance']:.2f}（不是账户余额）")
    if res['refund_rows']:
        print("\n含退款的原交易明细（部分退款保留未退金额）：")
        for r in sorted(res['refund_rows'], key=lambda x: x['dt']):
            print(f"   {r['iso']} {str(r['交易对方'])[:18]:<20}{str(r['商品'])[:28]:<30}"
                  f"{r['金额']:>9.2f}  {r['当前状态']}")

    print(f"\n{'类别':<16}{'金额':>10}{'占比':>8}{'笔数':>6}{'笔均':>9}{'月化':>10}")
    n = res['net']
    for k, v in res['categories'].items():
        pct = v['金额'] / n * 100 if n else 0
        mo = v['金额'] / res['days'] * 30 if res['days'] else 0
        print(f"{k:<16}{v['金额']:>10.2f}{pct:>7.1f}%{v['笔数']:>6}"
              f"{v['金额']/v['笔数']:>9.2f}{mo:>10.2f}")
    print(f"{'合计':<16}{n:>10.2f}{100.0:>7.1f}%{res['counts']['net_expense']:>6}")

    print("\n月度：")
    for m, v in res['monthly'].items():
        print(f"   {m}  {v['笔数']:>4} 笔  {v['金额']:>9.2f}")

    if res['meals']:
        print("\n餐饮餐次细分（仅「① 餐饮美食」类别）：")
        tot = sum(v['金额'] for v in res['meals'].values())
        nt = sum(v['笔数'] for v in res['meals'].values())
        for k, v in res['meals'].items():
            print(f"   {k:<10}{v['笔数']:>4} 笔{v['金额']:>9.2f}  笔均 {v['金额']/v['笔数']:>6.2f}"
                  f"  日均 {v['金额']/res['days']:>6.2f}  月化 {v['金额']/res['days']*30:>8.2f}")
        print(f"   {'合计':<10}{nt:>4} 笔{tot:>9.2f}")

    print("\n收入构成：")
    for k, v in res['income_breakdown'].items():
        print(f"   {k:<18}{v['金额']:>10.2f}  {v['笔数']:>3} 笔")

    if res['other_rows']:
        print("\n不计收支（充值/提现等）：")
        for r in res['other_rows']:
            print(f"   {r['iso']} {str(r['交易类型']):<10}{r['金额']:>10.2f}")

    pre = [r for r in res['net_rows']
           if any(k in str(r['商品']) + str(r['交易对方']) for k in ('充值', '开卡', '校园卡'))]
    if pre:
        print("\n注意：以下为预付/充值类支出。后续消费可能不再出现在账单里（如校园卡刷卡），")
        print("      分析实际消费时不要与之重复计算：")
        for r in pre:
            print(f"   {r['iso']} {str(r['交易对方'])[:18]:<20}{str(r['商品'])[:26]:<28}{r['金额']:>9.2f}")


# ----------------------------------------------------------------------------
# 五、Excel 报告
# ----------------------------------------------------------------------------
def write_report(res, out_path, title='账单分析报告'):
    import xlsxwriter
    bk = xlsxwriter.Workbook(out_path, {'nan_inf_to_errors': True})
    F = {
        'title': bk.add_format({'font_size': 15, 'bold': True, 'font_color': '#1F4E79'}),
        'h2': bk.add_format({'font_size': 12, 'bold': True, 'font_color': '#1F4E79'}),
        'sub': bk.add_format({'font_size': 10, 'font_color': '#808080'}),
        'h': bk.add_format({'bold': True, 'bg_color': '#1F4E79', 'font_color': 'white',
                            'border': 1, 'align': 'center', 'valign': 'vcenter', 'text_wrap': True}),
        'lbl': bk.add_format({'bold': True, 'bg_color': '#DDEBF7', 'border': 1}),
        'txt': bk.add_format({'border': 1, 'text_wrap': True, 'valign': 'top'}),
        'money': bk.add_format({'num_format': '#,##0.00', 'border': 1}),
        'money_lbl': bk.add_format({'num_format': '#,##0.00', 'border': 1,
                                    'bg_color': '#DDEBF7', 'bold': True}),
        'pct': bk.add_format({'num_format': '0.0%', 'border': 1}),
        'int': bk.add_format({'num_format': '0', 'border': 1, 'align': 'center'}),
        'date': bk.add_format({'num_format': 'yyyy-mm-dd hh:mm', 'border': 1}),
    }
    s = bk.add_worksheet('统计概览')
    for c, w in [('A', 22), ('B', 16), ('C', 34)]:
        s.set_column(f'{c}:{c}', w)
    s.write('A1', title, F['title'])
    rc = res['range']
    s.write('A2', f"区间 {rc['first']} ~ {rc['last']}（{res['days']} 天）", F['sub'])
    r = 3
    for k, v in [('交易笔数', res['counts']['total']), ('支出笔数', res['counts']['expense']),
                 ('其中已退款', res['counts']['refunded']), ('支出流水', res['gross']),
                 ('退款', res['refund']), ('净支出', res['net']), ('日均净支出', res['daily']),
                 ('收入', res['income']), ('结余', res['balance'])]:
        s.write(r, 0, k, F['lbl'])
        if isinstance(v, float):
            s.write_number(r, 1, v, F['money_lbl'])
        else:
            s.write(r, 1, v, F['txt'])
        r += 1
    r += 1
    s.write(r, 0, '分类汇总', F['h2']); r += 1
    for i, x in enumerate(['类别', '金额(元)', '占比', '笔数', '笔均(元)', '月化(元)'], start=1):
        s.write(r, i - 1, x, F['h'])
    r += 1
    f0 = r
    ncat = len(res['categories'])
    total_row = f0 + ncat + 1          # 合计行（Excel 行号）：数据区起点 + 类别数
    for k, v in res['categories'].items():
        s.write(r, 0, k, F['txt'])
        s.write_number(r, 1, v['金额'], F['money'])
        s.write_formula(r, 2, f'=B{r+1}/$B${total_row}', F['pct'])
        s.write_number(r, 3, v['笔数'], F['int'])
        s.write_formula(r, 4, f'=B{r+1}/D{r+1}', F['money'])
        s.write_formula(r, 5, f'=B{r+1}/{res["days"]}*30', F['money'])
        r += 1
    s.write(r, 0, '合计', F['lbl'])
    s.write_formula(r, 1, f'=SUM(B{f0+1}:B{r})' if ncat else '=0', F['money_lbl'])
    s.write_formula(r, 2, '=1' if ncat else '=0', F['pct'])
    s.write_formula(r, 3, f'=SUM(D{f0+1}:D{r})' if ncat else '=0', F['lbl'])
    s.write_formula(r, 4, f'=B{r+1}/D{r+1}' if ncat else '=0', F['money_lbl'])
    s.write_formula(r, 5, f'=B{r+1}/{res["days"]}*30' if res['days'] else '=0', F['money_lbl'])

    s2 = bk.add_worksheet('全部明细')
    cols = ['交易时间', '收/支', '交易类型', '交易对方', '金额(元)', '分类',
            '支付方式', '当前状态', '商品']
    widths = [19, 8, 14, 20, 11, 15, 13, 14, 40]
    for i, (x, w) in enumerate(zip(cols, widths), start=1):
        s2.write(0, i - 1, x, F['h'])
        s2.set_column(i - 1, i - 1, w)
    for i, x in enumerate(sorted(res['all_rows'], key=lambda y: y['dt']), start=1):
        s2.write_datetime(i, 0, x['dt'], F['date'])
        s2.write(i, 1, str(x.get('收/支') or ''), F['txt'])
        s2.write(i, 2, str(x['交易类型']), F['txt'])
        s2.write(i, 3, str(x['交易对方']), F['txt'])
        s2.write_number(i, 4, x['金额'], F['money'])
        s2.write(i, 5, str(x.get('类别') or ''), F['txt'])
        s2.write(i, 6, str(x['支付方式']), F['txt'])
        s2.write(i, 7, str(x['当前状态']), F['txt'])
        s2.write(i, 8, str(x['商品'] or ''), F['txt'])
    s2.autofilter(0, 0, len(res['all_rows']), len(cols) - 1)
    s2.freeze_panes(1, 0)
    bk.close()


# ----------------------------------------------------------------------------
def main():
    ap = argparse.ArgumentParser(description='微信支付账单分析')
    ap.add_argument('bill', help='微信账单 xlsx 路径')
    ap.add_argument('--out', help='输出 Excel 报告路径')
    ap.add_argument('--json', help='输出 JSON 统计路径')
    ap.add_argument('--from', dest='lo', help='起始日期 YYYY-MM-DD（含）')
    ap.add_argument('--to', dest='hi', help='结束日期 YYYY-MM-DD（不含）')
    ap.add_argument('--rules', help='自定义分类规则 JSON（{"类别": ["关键词", ...]}，'
                                    '示例见 examples/rules.example.json）')
    ap.add_argument('--family', help='亲人/资助识别关键词，逗号分隔')
    ap.add_argument('--title', default='账单分析报告')
    ap.add_argument('--no-report', action='store_true', help='不输出 Excel')
    a = ap.parse_args()

    rows, headers, hr = load_bill(a.bill)
    print(f"已读取 {len(rows)} 条记录（表头在第 {hr} 行）")
    rules = load_rules(a.rules) if a.rules else None
    fam = [s.strip() for s in a.family.split(',') if s.strip()] if a.family else None
    res = summarize(rows, rules=rules, lo=a.lo, hi=a.hi, family=fam)
    print_report(res)

    if a.json:
        slim = {k: v for k, v in res.items()
                if k not in ('net_rows', 'income_rows', 'other_rows', 'all_rows', 'refund_rows')}
        with open(a.json, 'w', encoding='utf-8') as f:
            json.dump(slim, f, ensure_ascii=False, indent=1, default=str)
        print(f"\n已写出统计 JSON：{a.json}")
    if a.out and not a.no_report:
        write_report(res, a.out, title=a.title)
        print(f"已写出 Excel 报告：{a.out}")


if __name__ == '__main__':
    main()
