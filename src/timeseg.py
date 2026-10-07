# -*- coding: utf-8 -*-
"""
按自定义时段（学期/假期/军训等）分段统计支出。

用法:
    python timeseg.py <账单.xlsx> --segments 分段.json [--out 分段报告.xlsx]

分段.json 格式（数组，互斥且应连续覆盖账单区间）:
    [
      {"name": "A 暑假", "lo": "2026-07-03", "hi": "2026-08-29", "kind": "假期", "note": "..."},
      {"name": "D 军训", "lo": "2026-09-10", "hi": "2026-09-25", "kind": "在校", "note": "..."}
    ]

关键设计:
    - 分段必须互斥且连续；任何一笔支出落在所有分段之外都要报错并列出，
      否则合计对不上（这是最容易犯的错）。
    - 输出"假期/在校"汇总，并给出各段日均，日均比总额更有决策价值。
    - 可选：把某个分段内的时间窗单独拆出来（如"暑假中的旅行"）。
"""
import sys, io, os, json, datetime, argparse, collections
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from analyze import load_bill, classify, compile_rules, _force_utf8_stdio, net_expenses

_force_utf8_stdio()


def in_seg(iso, seg):
    return seg['lo'] <= iso < seg['hi']


def segment(rows, segments, rules=None):
    rules = compile_rules(rules)
    exp = [r for r in rows if r.get('收/支') == '支出']
    net_rows = net_expenses(exp)
    ordered=sorted(segments,key=lambda s:s['lo'])
    for i,s in enumerate(ordered):
        if s['hi'] <= s['lo']: raise ValueError('分段结束日期必须晚于开始日期')
        if i and s['lo'] < ordered[i-1]['hi']: raise ValueError('分段不允许重叠')
    for r in net_rows:
        r['类别'] = classify(r, rules)

    out = []
    assigned = set()
    for seg in segments:
        sub = []
        for r in net_rows:
            if in_seg(r['iso'], seg):
                sub.append(r)
                assigned.add(id(r))
        days = (datetime.date.fromisoformat(seg['hi']) - datetime.date.fromisoformat(seg['lo'])).days
        amt = round(sum(r['金额'] for r in sub), 2)
        by = collections.defaultdict(lambda: {'笔数': 0, '金额': 0.0})
        for r in sub:
            by[r['类别']]['笔数'] += 1
            by[r['类别']]['金额'] = round(by[r['类别']]['金额'] + r['金额'], 2)
        out.append({**seg, 'days': days, 'n': len(sub), 'amt': amt,
                    'daily': round(amt / days, 2) if days else 0,
                    'by_cat': dict(sorted(by.items(), key=lambda kv: -kv[1]['金额'])),
                    'items': sub})

    # 完整性校验：不允许有支出落在所有分段之外
    missed = [r for r in net_rows if id(r) not in assigned]
    total = round(sum(r['金额'] for r in net_rows), 2)
    seg_sum = round(sum(s['amt'] for s in out), 2)
    if missed:
        print(f"[警告] {len(missed)} 笔支出未被任何分段覆盖，合计 "
              f"{sum(r['金额'] for r in missed):.2f} 元 —— 分段有空隙：")
        for r in missed:
            print(f"    {r['iso']} {str(r['交易对方'])[:18]:<20}{r['金额']:>9.2f}")
    assert abs(seg_sum + sum(r['金额'] for r in missed) - total) < 0.005, "分段合计 != 净支出"

    kinds = collections.defaultdict(lambda: {'days': 0, 'n': 0, 'amt': 0.0})
    for s in out:
        k = kinds[s.get('kind') or '其他']
        k['days'] += s['days']; k['n'] += s['n']
        k['amt'] = round(k['amt'] + s['amt'], 2)
    for k, v in kinds.items():
        v['daily'] = round(v['amt'] / v['days'], 2) if v['days'] else 0

    return {'segments': out, 'by_kind': dict(kinds), 'total': total,
            'days_total': sum(s['days'] for s in out), 'missed': missed}


def print_seg(res):
    print("=" * 86)
    print(f"{'时段':<22}{'性质':<6}{'天数':>5}{'笔数':>6}{'净支出':>11}{'日均':>9}{'占净支出':>10}")
    tot = res['total']
    for s in res['segments']:
        print(f"{s['name']:<22}{s.get('kind',''):<6}{s['days']:>5}{s['n']:>6}"
              f"{s['amt']:>11.2f}{s['daily']:>9.2f}{(s['amt']/tot*100 if tot else 0):>9.1f}%")
    print(f"{'合计':<22}{'':<6}{res['days_total']:>5}"
          f"{sum(s['n'] for s in res['segments']):>6}{tot:>11.2f}")
    print("\n按性质汇总（日均比总额更有决策价值）：")
    for k, v in sorted(res['by_kind'].items(), key=lambda kv: -kv[1]['amt']):
        print(f"   {k:<8}{v['days']:>4} 天{v['n']:>5} 笔{v['amt']:>10.2f}  "
              f"日均 {v['daily']:>7.2f}  占 {(v['amt']/tot*100 if tot else 0):>5.1f}%")
    print("\n各时段分类明细：")
    for s in res['segments']:
        if not s['by_cat']:
            continue
        print(f"\n   {s['name']}（{s['lo']}~{s['hi']}，{s['n']} 笔 {s['amt']:.2f} 元）")
        for c, v in s['by_cat'].items():
            print(f"      {c:<16}{v['金额']:>10.2f}  {v['笔数']:>3} 笔")


def write_seg_report(res, path):
    import xlsxwriter
    bk = xlsxwriter.Workbook(path, {'nan_inf_to_errors': True})
    t = bk.add_format({'font_size': 15, 'bold': True, 'font_color': '#1F4E79'})
    h = bk.add_format({'bold': True, 'bg_color': '#1F4E79', 'font_color': 'white',
                       'border': 1, 'align': 'center', 'valign': 'vcenter', 'text_wrap': True})
    lbl = bk.add_format({'bold': True, 'bg_color': '#DDEBF7', 'border': 1})
    txt = bk.add_format({'border': 1, 'text_wrap': True, 'valign': 'top'})
    mny = bk.add_format({'num_format': '#,##0.00', 'border': 1})
    pct = bk.add_format({'num_format': '0.0%', 'border': 1})
    itg = bk.add_format({'num_format': '0', 'border': 1, 'align': 'center'})
    s = bk.add_worksheet('分段统计')
    for c, w in [('A', 24), ('B', 8), ('C', 7), ('D', 7), ('E', 12), ('F', 10), ('G', 10), ('H', 40)]:
        s.set_column(f'{c}:{c}', w)
    s.write('A1', '分时段支出统计', t)
    s.write('A2', '分段来自 segments.json；各段互斥，合计等于区间净支出', bk.add_format({'font_size': 10, 'font_color': '#808080'}))
    for i, x in enumerate(['时段', '性质', '天数', '笔数', '净支出(元)', '日均(元)', '占比', '说明'], start=1):
        s.write(2, i - 1, x, h)
    r = 3
    f0 = r
    segs = res['segments']
    for sgm in segs:
        s.write(r, 0, sgm['name'], txt)
        s.write(r, 1, sgm.get('kind', ''), itg)
        s.write_number(r, 2, sgm['days'], itg)
        s.write_number(r, 3, sgm['n'], itg)
        s.write_number(r, 4, sgm['amt'], mny)
        s.write_formula(r, 5, f'=E{r+1}/C{r+1}', mny)
        # 分母是合计行 = 数据区起点 + 时段数 + 1（漏掉 +1 会指向最后一个时段）
        s.write_formula(r, 6, f'=E{r+1}/$E${f0 + len(segs) + 1}' if res['total'] else '=0', pct)
        s.write(r, 7, sgm.get('note', ''), txt)
        r += 1
    s.write(r, 0, '合计', lbl)
    for c in (1,):
        s.write(r, c, '', lbl)
    s.write_formula(r, 2, f'=SUM(C{f0+1}:C{r})', lbl)
    s.write_formula(r, 3, f'=SUM(D{f0+1}:D{r})', lbl)
    s.write_formula(r, 4, f'=SUM(E{f0+1}:E{r})', lbl)
    s.write_formula(r, 5, f'=E{r+1}/C{r+1}', lbl)
    s.write_formula(r, 6, '=1' if res['total'] else '=0', lbl)
    s.write(r, 7, '', lbl)
    bk.close()


def main():
    ap = argparse.ArgumentParser(description='按自定义时段分段统计')
    ap.add_argument('bill')
    ap.add_argument('--segments', required=True, help='分段定义 JSON')
    ap.add_argument('--out', help='输出 Excel 路径')
    a = ap.parse_args()
    rows, _, _ = load_bill(a.bill)
    with open(a.segments, encoding='utf-8-sig') as f:
        segs = json.load(f)
    res = segment(rows, segs)
    if res['missed']:
        raise SystemExit('分段未覆盖全部支出；请补齐日期区间后再生成报告。')
    print_seg(res)
    if a.out:
        write_seg_report(res, a.out)
        print(f"\n已写出：{a.out}")


if __name__ == '__main__':
    main()
