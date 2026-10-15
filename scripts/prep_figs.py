#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""生成前处理图位：图片文件存在则嵌入，缺失则该图位降级为占位框（不写 file 字段）。
用法： python prep_figs.py 输入.json 输出.json
"""
import json, os, sys

src, dst = sys.argv[1], sys.argv[2]
base = os.path.dirname(os.path.abspath(src))
d = json.load(open(src, encoding='utf-8'))
have = miss = 0

def walk(node):
    global have, miss
    if isinstance(node, dict):
        if node.get('type') == 'figure' and node.get('file'):
            f = os.path.join(base, node['file'])
            if os.path.exists(f):
                have += 1
            else:
                node.pop('file'); miss += 1
        for v in node.values():
            walk(v)
    elif isinstance(node, list):
        for v in node:
            walk(v)

walk(d)
json.dump(d, open(dst, 'w', encoding='utf-8'), ensure_ascii=False, indent=1)
print('图位：已就绪 %d 个，缺图降级为占位框 %d 个' % (have, miss))
