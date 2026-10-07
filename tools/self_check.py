# -*- coding: utf-8 -*-
"""
提交前的隐私自查。

用途：如果你打算分享**真实账单**（例如截图、或者把账单放进仓库），
先跑一遍这个脚本，它会列出账单里所有可能暴露个人信息的内容。

    python tools/self_check.py 你的账单.xlsx

本脚本只读文件，不做任何修改，也不联网。
退出码：0 = 正常完成，2 = 用法错误。
"""
import argparse
import os
import re
import sys
import collections

sys.path.insert(0, os.path.join(os.path.dirname(os.path.abspath(__file__)), '..', 'src'))

try:
    sys.stdout.reconfigure(encoding='utf-8', errors='replace')
except Exception:
    pass


# 可能包含个人身份信息的字段
SENSITIVE_FIELDS = ['交易对方', '商户单号', '交易单号', '备注', '商品']

# 看起来像人名/昵称的模式
NAME_HINTS = [
    (r'[\u4e00-\u9fa5]{2,4}\s*[（(].{1,12}[)）]', '带昵称括号的疑似人名'),
    (r'^\s*(?:发给|转账给)?[\u4e00-\u9fa5]{2,3}\s*$', '疑似纯姓名'),
]

# 联系方式类。
# 注意两个容易误报的地方，我第一版都踩过：
#   1) 手机号：19 位订单号里天然含 11 位连续数字，必须排除"是更长数字串一部分"的情况。
#      第一版在 265 条记录上误报 136 处"手机号"。
#   2) 身份证号：18 位订单号同样会被匹配。用身份证校验位算法过滤后误报基本消失。
CONTACT_PATTERNS = [
    (r'[\w.+-]+@[\w-]+\.[\w.]+', '邮箱'),
    (r'\d{4}\*{2,}\d{2,4}', '被掩码的号码（如 172****1776）'),
]


def _id_checksum_ok(s):
    """中国大陆 18 位身份证验证：校验位 + 出生日期合法性。

    只验校验位还不够——随机的 18 位订单号约有 1/11 的概率碰巧通过校验位
    （实测 439 条记录里就有 3 个）。再加上"第 7-14 位必须是合法出生日期"
    这一条，误报基本消失。
    """
    if not re.fullmatch(r'\d{17}[\dXx]', s):
        return False
    # 出生日期：第 7-14 位（索引 6-13）应为 YYYYMMDD
    try:
        y, m, d = int(s[6:10]), int(s[10:12]), int(s[12:14])
        import datetime as _dt
        _dt.date(y, m, d)
        if not (1900 <= y <= 2030):
            return False
    except ValueError:
        return False
    # 校验位
    w = [7, 9, 10, 5, 8, 4, 2, 1, 6, 3, 7, 9, 10, 5, 8, 4, 2]
    codes = '10X98765432'
    total = sum(int(s[i]) * w[i] for i in range(17))
    return s[17].upper() == codes[total % 11]


def find_phones_and_ids(rows):
    """找出真正的手机号 / 身份证号。

    关键：只在"商户名称"（交易对方）里找嵌着的号码。
    订单号（商户单号/交易单号）是 16-20 位随机数字，里面**天然**会包含
    形如手机号的 11 位片段——把这些当号码报出来毫无意义（第二版因此在
    439 条记录上误报 492 个"商户"）。
    """
    standalone_phone, embedded_phone, real_id = [], [], []
    for r in rows:
        name = str(r['交易对方'] or '')
        # 1) 名称里独立出现的手机号（如"某某店17324313833"）
        for m in re.findall(r'(?<!\d)1[3-9]\d{9}(?!\d)', name):
            embedded_phone.append((r['iso'], name, m))
        # 2) 任何字段里独立出现的手机号
        blob = ' '.join(str(r.get(f) or '') for f in SENSITIVE_FIELDS)
        for m in re.findall(r'(?<!\d)1[3-9]\d{9}(?!\d)', blob):
            standalone_phone.append((r['iso'], name, m))
        # 3) 身份证号：必须通过校验位，否则只是订单号
        for m in re.findall(r'(?<!\d)\d{17}[\dXx](?!\d)', blob):
            if _id_checksum_ok(m):
                real_id.append((r['iso'], name, m))
    # 去重（同一商户反复出现只报一次）
    def dedup(items):
        seen, out = set(), []
        for it in items:
            key = (it[1], it[2])
            if key not in seen:
                seen.add(key)
                out.append(it)
        return out
    return dedup(standalone_phone), dedup(embedded_phone), dedup(real_id)

# 已知的非个人名（支付方式、连锁商户、平台），避免把商户误判成人名
KNOWN_NON_PERSON = {
    '羊城通', '岭南通', '美团', '饿了么', '麦当劳', '肯德基', '真功夫', '德克士',
    '喜茶', '霸王茶姬', '茶百道', '瑞幸', '蜜雪冰城', '星巴克', '美宜佳', '新力有家',
    '7-ELEVEN', '24H智能售货', '示例洗衣服务', '携程', '闲鱼', '拼多多', '京东',
    '淘宝平台商户', '抖音电商商家', '中铁网络', '深圳市地铁', '示例大学',
    '杭州深度求索', '外研社旗舰店', '大麦网', '中国移动', '中国电信', '中国联通',
}
NON_PERSON_HINTS = ('银行', '储蓄卡', '信用卡', '地铁', '公交', '出行', '铁路', '航空',
                    '有限', '公司', '商店', '超市', '便利店', '餐厅', '酒店', '平台',
                    '商户', '药店', '医院', '大学', '学院', '中心', '景区')


def is_probably_person(name):
    """保守判断一个交易对方是否可能是个人。宁可漏报，也不要大量误报。"""
    n = str(name).strip()
    if not n or n in ('/', 'None'):
        return False
    if n in KNOWN_NON_PERSON:
        return False
    if any(h in n for h in NON_PERSON_HINTS):
        return False
    if re.search(r'[（(].{1,12}[)）]', n):       # 带昵称括号
        return True
    if re.fullmatch(r'[\u4e00-\u9fa5]{2,4}', n) and not any(
            h in n for h in ('支付', '转账', '收款', '退款', '红包')):
        return True
    return False

# 地点线索（能推断你住在哪里 / 去了哪里）
LOCATION_HINTS = r'(?:省|市|区|县|镇|村|路|街|道|大厦|广场|校区|宿舍|公寓|机场|车站|地铁站)'


def main():
    ap = argparse.ArgumentParser(
        description='分享真实账单前做一次隐私自查（只读、不联网）')
    ap.add_argument('bill', help='微信账单 xlsx 路径')
    a = ap.parse_args()
    path = a.bill
    if not os.path.exists(path):
        raise SystemExit(f'文件不存在：{path}')

    from analyze import load_bill

    rows, headers, hr = load_bill(path)
    print('=' * 74)
    print(f'隐私自查：{os.path.basename(path)}')
    print(f'表头位于第 {hr} 行，共 {len(rows)} 条记录')
    print('=' * 74)

    # 1) 抬头里的账号信息
    from openpyxl import load_workbook
    ws = load_workbook(path, data_only=True)[load_workbook(path).sheetnames[0]]
    print('\n【1】文件抬头（可能含昵称与账号）：')
    for i in range(1, hr):
        v = ws.cell(row=i, column=1).value
        if v:
            print(f'    {str(v)[:110]}')

    # 2) 交易对方里出现的所有疑似个人名
    print('\n【2】疑似个人的交易对方（建议打码）：')
    cnt = collections.Counter(str(r['交易对方']) for r in rows)
    persons = [(n, c) for n, c in cnt.most_common() if is_probably_person(n)]
    others = [(n, c) for n, c in cnt.most_common() if not is_probably_person(n)]
    if persons:
        for name, n in persons[:25]:
            print(f'      {name:<28} {n:>3} 次')
        if len(persons) > 25:
            print(f'      ... 其余 {len(persons) - 25} 个省略')
    else:
        print('      未发现')
    print(f'    其余 {len(others)} 个对方判定为商户/平台/支付方式（通常可保留）：')
    for name, n in others[:8]:
        print(f'      {name:<28} {n:>3} 次')
    if len(others) > 8:
        print(f'      ... 其余 {len(others) - 8} 个省略')

    # 3) 联系方式与证件号
    print('\n【3】联系方式 / 证件号线索：')
    phone, embedded, real_id = find_phones_and_ids(rows)
    if phone:
        print(f'    手机号（独立出现，高置信）：{len(phone)} 处')
        for iso, cp, m in phone[:6]:
            print(f'      {iso}  {cp[:22]:<24} {m}')
        if len(phone) > 6:
            print(f'      ... 其余 {len(phone) - 6} 处省略')
    # 商户名里嵌着的号码（例如"某某店17324313833"），这些是真实存在的公开号码，
    # 但会暴露你常去的店铺，建议一并确认
    seen = set()
    uniq_embedded = []
    for iso, cp, m in embedded:
        if (cp, m) not in seen:
            seen.add((cp, m))
            uniq_embedded.append((iso, cp, m))
    if uniq_embedded:
        print(f'    商户名中嵌有号码（{len(uniq_embedded)} 个不同商户，会暴露常去店铺）：')
        for iso, cp, m in uniq_embedded[:6]:
            print(f'      {cp[:30]:<32} {m}')
        if len(uniq_embedded) > 6:
            print(f'      ... 其余 {len(uniq_embedded) - 6} 个省略')
    if real_id:
        print(f'    身份证号（已通过校验位验证，高置信）：{len(real_id)} 处')
        for iso, cp, m in real_id[:6]:
            print(f'      {iso}  {cp[:22]:<24} {m}')
    # 邮箱等文本型模式
    hits = collections.defaultdict(list)
    for r in rows:
        blob = ' '.join(str(r.get(f) or '') for f in SENSITIVE_FIELDS)
        for pat, label in CONTACT_PATTERNS:
            for m in re.findall(pat, blob):
                hits[label].append((r['iso'], str(r['交易对方']), m))
    for label, items in hits.items():
        print(f'    {label}：{len(items)} 处')
        for iso, cp, m in items[:4]:
            print(f'      {iso}  {cp[:22]:<24} {m}')
    if not (phone or uniq_embedded or real_id or hits):
        print('    未发现')
    print('    说明：身份证号需同时通过"校验位"与"出生日期合法"两项验证，')
    print('          随机订单号几乎不可能通过（实测 439 条记录误报从 8 处降到 0 处）；')
    print('          11 位手机号只在独立出现时计入，且只在商户名称里找嵌号，')
    print('          避免把订单号里的数字片段误报成号码。')

    # 4) 地点线索
    print('\n【4】地点线索（能推断居住地/行程）：')
    locs = collections.Counter()
    for r in rows:
        for f in ('交易对方', '商品'):
            for m in re.findall(rf'[\u4e00-\u9fa5]{{2,10}}{LOCATION_HINTS}', str(r.get(f) or '')):
                locs[m] += 1
    for name, n in locs.most_common(15):
        print(f'    {name:<28} {n:>3} 次')
    if not locs:
        print('    未发现')

    # 5) 金额区间（暴露消费水平）
    amts = [r['金额'] for r in rows if r.get('收/支') == '支出']
    if amts:
        print(f'\n【5】金额区间：{min(amts):.2f} ~ {max(amts):.2f} 元 '
              f'（共 {len(amts)} 笔支出；单笔最大额本身也是一种信息）')

    # 结论与建议
    print('\n' + '=' * 74)
    print('建议：')
    print('  1. 如果要分享真实账单，先用 Excel 手工替换【2】里的个人姓名与【3】里的号码；')
    print('  2. 更省事的做法是使用合成演示数据：python tools/make_demo_data.py')
    print('     —— 结构与陷阱都齐全，且不含任何真实信息，可直接公开；')
    print('  3. 截图时注意避开抬头区域的昵称，以及金额列的单笔大额；')
    print('  4. 仓库里的 .gitignore 已默认排除 *微信支付账单* 与 *账单流水*，')
    print('     但推送前仍建议用 git status 复查一遍待提交文件列表。')
    print('=' * 74)
    return 0


if __name__ == '__main__':
    sys.exit(main())
