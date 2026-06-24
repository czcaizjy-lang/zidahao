#!/usr/bin/env python3
"""
构建独立自达号看板 HTML 文件
读取 zidahao.html（模板）和 zidahao_data.json（数据），
生成内嵌数据的 public/zidahao.html 和 public/index.html（GitHub Pages 入口）。
"""

import json
import os
import sys

BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
TEMPLATE_PATH = os.path.join(BASE_DIR, 'zidahao.html')
DATA_PATH = os.path.join(BASE_DIR, 'data', 'zidahao_data.json')
OUTPUT_PATHS = [
    os.path.join(BASE_DIR, 'public', 'zidahao.html'),
    os.path.join(BASE_DIR, 'public', 'index.html'),
]

OLD_FETCH = """fetch('data/zidahao_data.json?t=' + Date.now()).then(r => r.json()).then(d => {
  DATA = d;
  render();
}).catch(e => { console.error('Failed to load zidahao data:', e); });"""


def build():
    with open(TEMPLATE_PATH, 'r', encoding='utf-8') as f:
        template = f.read()

    if OLD_FETCH not in template:
        print('✗ 模板中未找到 fetch 语句，可能已经变更', file=sys.stderr)
        return False

    with open(DATA_PATH, 'r', encoding='utf-8') as f:
        data = json.load(f)

    data_str = json.dumps(data, ensure_ascii=False, separators=(',', ':'))

    parts = template.split(OLD_FETCH, 1)
    before = parts[0]
    after = parts[1]

    inline_block = (
        '</script>\n'
        '<script id="inline-data" type="application/json">' + data_str + '</script>\n'
        '<script>\n'
        'DATA = JSON.parse(document.getElementById("inline-data").textContent);\n'
        'setTimeout(render, 0);'
    )

    output = before + inline_block + after

    for path in OUTPUT_PATHS:
        os.makedirs(os.path.dirname(path), exist_ok=True)
        with open(path, 'w', encoding='utf-8') as f:
            f.write(output)

    summary = data.get('summary', {})
    gmv = summary.get('直播GMV', 0)
    anchors = len(data.get('anchors', []))
    subs = len(data.get('sub_agencies', []))
    size_kb = len(output) / 1024

    print(f'✓ 已生成 {len(OUTPUT_PATHS)} 个文件')
    for p in OUTPUT_PATHS:
        print(f'  → {os.path.basename(p)}')
    print(f'  模板: {os.path.basename(TEMPLATE_PATH)} ({len(template)} 字符)')
    print(f'  数据: {os.path.basename(DATA_PATH)} ({len(data_str)} 字符)')
    print(f'  输出: {size_kb:.0f} KB')
    print(f'  总GMV: ¥{gmv:,.2f} | 达人: {anchors} | 子机构: {subs}')
    return True


if __name__ == '__main__':
    ok = build()
    sys.exit(0 if ok else 1)
