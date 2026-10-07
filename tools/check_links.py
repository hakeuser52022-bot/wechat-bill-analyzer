# -*- coding: utf-8 -*-
"""校验项目内所有 Markdown 的本地链接是否都指向真实存在的文件。

用法: python tools/check_links.py
仅检查相对路径链接（忽略 http/https/mailto 与纯锚点）。
"""
import os
import re
import sys

try:
    sys.stdout.reconfigure(encoding='utf-8', errors='replace')
except Exception:
    pass

ROOT = os.path.abspath(os.path.join(os.path.dirname(os.path.abspath(__file__)), '..'))
LINK = re.compile(r'\[[^\]]*\]\(([^)]+)\)')
SKIP_SCHEME = ('http://', 'https://', 'mailto:', '#', 'data:')


def main():
    bad = []
    checked = 0
    files = []
    for dirpath, dirnames, filenames in os.walk(ROOT):
        dirnames[:] = [d for d in dirnames if d not in ('__pycache__', '.git')]
        for fn in filenames:
            if fn.lower().endswith('.md'):
                files.append(os.path.join(dirpath, fn))

    for path in sorted(files):
        rel = os.path.relpath(path, ROOT).replace(os.sep, '/')
        with open(path, encoding='utf-8') as f:
            text = f.read()
        # 跳过代码块，避免把示例命令里的方括号当链接
        text = re.sub(r'```.*?```', '', text, flags=re.S)
        for m in LINK.finditer(text):
            target = m.group(1).strip()
            if not target or target.startswith(('http://', 'https://', 'mailto:', 'data:')):
                continue
            anchor = ''
            if '#' in target:
                target, anchor = target.split('#', 1)
            target = target.split('?')[0]

            # 纯页内锚点（形如 (#标题)）也要校验：
            # 早期版本把 '#' 放进 SKIP_SCHEME，导致这类链接被整个跳过，
            # 改坏了也检不出来。
            if not target:
                if not anchor:
                    continue
                resolved = path
            else:
                resolved = os.path.normpath(os.path.join(os.path.dirname(path), target))
            checked += 1

            if not os.path.exists(resolved):
                bad.append((rel, m.group(1)))
            elif anchor and resolved.lower().endswith('.md'):
                with open(resolved, encoding='utf-8') as f:
                    body = f.read()

                def slugify(h):
                    """复现 GitHub 的锚点规则。

                    GitHub 的做法是：小写 → 去掉标点（保留中文、字母、数字、
                    空格与连字符）→ 空格转连字符 → 合并连续连字符。
                    注意**不是**把标点替换成连字符：标题
                    "本项目专用：复制粘贴版" 的锚点是
                    "本项目专用复制粘贴版"，全角冒号被直接删掉。
                    早期版本把标点替换成 "-"，导致正确锚点被误报为不存在。
                    """
                    h = re.sub(r'[`*_\[\]()]', '', h).strip().lower()
                    h = re.sub(r'[^\w\u4e00-\u9fa5\s-]', '', h)
                    h = h.replace(' ', '-')
                    return re.sub(r'-+', '-', h).strip('-')

                heads = re.findall(r'^#{1,6}\s+(.*)$', body, flags=re.M)
                if anchor.lower() not in [slugify(h) for h in heads]:
                    bad.append((rel, f'{m.group(1)}  (锚点不存在)'))

    print(f'检查了 {len(files)} 个 Markdown 文件、{checked} 个本地链接')
    if bad:
        print(f'\n失效链接 {len(bad)} 个：')
        for rel, t in bad:
            print(f'  {rel}  ->  {t}')
        return 1
    print('全部链接有效 ✅')
    return 0


if __name__ == '__main__':
    sys.exit(main())
