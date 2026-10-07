# -*- coding: utf-8 -*-
"""一次性跑完全部脚本并保存真实输出到 docs/ 与 examples/output/。

用法: python tools/run_demo.py
"""
import os
import hashlib
import json
from pathlib import Path
import re
import subprocess
import sys
sys.stdout.reconfigure(encoding="utf-8", errors="replace")
sys.stderr.reconfigure(encoding="utf-8", errors="replace")

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.abspath(os.path.join(HERE, '..'))
BILL = os.path.join(ROOT, 'examples', 'data', 'demo_bill.xlsx')
SEG = os.path.join(ROOT, 'examples', 'segments.demo.json')
OUT = os.path.join(ROOT, 'examples', 'output')
DOCS = os.path.join(ROOT, 'docs')

STEPS = [
    ('analyze.py',
     ['src/analyze.py', BILL, '--family', '演示资助者A,演示资助者B',
      '--out', os.path.join(OUT, 'demo_report.xlsx'),
      '--json', os.path.join(OUT, 'demo_stats.json')]),
    ('timeseg.py',
     ['src/timeseg.py', BILL, '--segments', SEG,
      '--out', os.path.join(OUT, 'demo_segments.xlsx')]),
    ('meals.py',
     ['src/meals.py', BILL, '--from', '2024-04-16', '--days', '31',
      '--json', os.path.join(OUT, 'demo_meals.json')]),
    ('budget.py',
     ['src/budget.py', BILL, '--from', '2024-04-16', '--month-days', '31',
      '--entertainment', '180', '--json', os.path.join(OUT, 'demo_budget.json')]),
    # 等长时段对比：演示文档里"日均才是有决策价值的口径"一节引用的数字，
    # 放进 demo 输出以便任何人核对。
    ('analyze.py（等长对比：虚构阶段 A，15 天）',
     ['src/analyze.py', BILL, '--from', '2024-04-01', '--to', '2024-04-16',
      '--no-report', '--json', os.path.join(OUT, 'demo_window_summer.json')]),
    ('analyze.py（等长对比：虚构阶段 B，15 天）',
     ['src/analyze.py', BILL, '--from', '2024-04-16', '--to', '2024-05-01',
      '--no-report', '--json', os.path.join(OUT, 'demo_window_school.json')]),
    # 自定义分类规则：演示 --rules 用法（示例规则随仓库提供）
    ('analyze.py（自定义分类规则 --rules）',
     ['src/analyze.py', BILL, '--rules',
      os.path.join(ROOT, 'examples', 'rules.example.json'), '--no-report']),
    # 餐饮口径一致性核对：meals 与 budget 必须给出同一个餐饮总额
    ('餐饮口径一致性核对',
     ['tools/check_food_consistency.py', BILL]),
]


def relativize(text):
    """把输出里的绝对路径替换为相对路径，让文档在任何机器上都能复现。"""
    text = text.replace(ROOT + os.sep, '').replace(ROOT + '/', '')
    # 兜底：任何残留的绝对盘符路径都压成文件名
    text = re.sub(r'[A-Za-z]:\\[^\s"\']*[\\/]([^\\/\s"\']+)', r'\1', text)
    return text


def main():
    env = dict(os.environ, PYTHONIOENCODING='utf-8')
    blocks = []
    for name, args in STEPS:
        print(f'==> 运行 {name}')
        p = subprocess.run([sys.executable] + args, cwd=ROOT, env=env,
                           capture_output=True, text=True, encoding='utf-8')
        out = relativize((p.stdout or '') + (p.stderr or ''))
        shown = ' '.join(a if not os.path.isabs(a) else
                         os.path.relpath(a, ROOT).replace(os.sep, '/') for a in args)
        blocks.append((name, shown, out, p.returncode))
        print(out)
        print(f'    退出码 {p.returncode}')

    # 生成 docs/demo.md
    # 注意：这里直接写成 Markdown 链接（而不是 `path` 形式的 code span），
    # 否则每次重新生成本文件都会把链接退回成不可点击的代码片段，
    # 与 tools/dress_links.py 来回打架。
    lines = ['# 演示运行结果',
             '',
             '以下全部为**真实运行输出**，数据来自仓库内的合成演示账单',
             '（[demo_bill.xlsx](../examples/data/demo_bill.xlsx)，由 '
             '[make_demo_data.py](../tools/make_demo_data.py) 生成，',
             '不含任何真实交易信息）。',
             '',
             '复现方式：',
             '',
             '```bash',
             'python tools/make_demo_data.py   # 生成演示账单',
             'python tools/run_demo.py         # 跑全部脚本并重新生成本文件',
             '```',
             '']
    for name, args, out, rc in blocks:
        lines += [f'## {name}', '', '```bash', f'python {args}', '```', '',
                  f'退出码 `{rc}`。输出：', '', '```text', out.strip(), '```', '']
    with open(os.path.join(DOCS, 'demo.md'), 'w', encoding='utf-8') as f:
        f.write('\n'.join(lines))
    print(f'已写出 {os.path.join(DOCS, "demo.md")}')

    bad = [n for n, _, _, rc in blocks if rc != 0]
    if bad:
        print('以下脚本返回非零退出码：', bad)
        return 1
    hashes={}
    for folder in ('examples/data','examples/output'):
        for p in sorted((Path(ROOT)/folder).glob('*')):
            if p.is_file(): hashes[p.relative_to(ROOT).as_posix()]=hashlib.sha256(p.read_bytes()).hexdigest()
    (Path(ROOT)/'examples/artifacts.sha256.json').write_text(json.dumps(hashes,ensure_ascii=False,indent=2),encoding='utf-8')
    return 0


if __name__ == '__main__':
    sys.exit(main())
