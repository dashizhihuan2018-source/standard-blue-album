# -*- coding: utf-8 -*-
"""单栏版底栏贴底测量：B5/A4 自适应，页脚线按 PDF 页高动态计算，
以及底栏 y 值。聚焦 content 页，hybrid 页按 below_top 偏移单独处理。
用法: python measure_bottom_single.py <rendered.pdf> [page_types_single.json]
"""
import sys, os, json, subprocess, tempfile, statistics as st
from xml.etree import ElementTree as ET

PDF = sys.argv[1] if len(sys.argv) > 1 else ""
PTFILE = sys.argv[2] if len(sys.argv) > 2 else os.path.join(os.path.dirname(os.path.abspath(__file__)), 'page_types_single.json')
CM = 28.3465
NSX = 'http://www.w3.org/1999/xhtml'

TOP_Y = 4.2*CM
HEAD_LIMIT = 3.5*CM

_HERE = os.path.dirname(os.path.abspath(__file__))
PTYPES = json.load(open(PTFILE)) if os.path.exists(PTFILE) else None

bbox = tempfile.mktemp(suffix='.xml')
subprocess.run(['pdftotext', '-bbox-layout', PDF, bbox], check=True)
tree = ET.parse(bbox)

rows = []
page_no = 0
for page in tree.iter('{%s}page' % NSX):
    page_no += 1
    page_h = float(page.get('height'))
    foot_limit = page_h - 1.40*CM
    footer_line = page_h - 1.30*CM
    if page_no <= 4:
        continue
    _pt = PTYPES[page_no-1] if PTYPES and page_no <= len(PTYPES) else {'t':'content'}
    ptype = _pt['t'] if isinstance(_pt, dict) else _pt
    below_top = _pt.get('below_top') if isinstance(_pt, dict) else None
    if ptype in ('chapter', 'table'):
        continue
    col = []   # 单栏所有 body 行 (ymax, ymin)
    for line in page.iter('{%s}line' % NSX):
        xmin = float(line.get('xMin')); xmax = float(line.get('xMax'))
        ymin = float(line.get('yMin')); ymax = float(line.get('yMax'))
        h = ymax - ymin
        t = ''.join((w.text or '') for w in line.iter('{%s}word' % NSX))
        if not t.strip():
            continue
        if h > 40:
            continue
        if ymin < HEAD_LIMIT - 2.5 or ymin > foot_limit + 2.5:
            continue
        if ptype == 'hybrid' or (ptype == 'hybrid' and below_top is not None):
            pass
        if ptype == 'hybrid' and below_top is not None:
            if ymin < below_top*CM - 2.5:
                continue
        col.append((ymax, ymin))
    if not col:
        continue
    b = max(y for y, _ in col)
    bc = b/CM
    gap = (footer_line - b)/CM
    rows.append({'pg': page_no, 't': ptype, 'n': len(col), 'b': bc, 'g': gap})

os.unlink(bbox)

# 满页：正文行数 >= 18 行（单栏一页约 27 行；18 行以上视为"应贴底"的满页）
full = [r for r in rows if r['n'] >= 18]
tail = [r for r in rows if r['n'] < 18]

def summ(vals):
    vals = [v for v in vals if v is not None]
    if not vals: return "n=0"
    return f"n={len(vals)} mean={st.mean(vals):.2f} median={st.median(vals):.2f} max={max(vals):.2f} min={min(vals):.2f}"

print(f"=== 单栏 content/hybrid 页: {len(rows)} (满页 {len(full)} / 短页 {len(tail)}) ===")
print("满页 gap(距页脚线):", summ([r['g'] for r in full]))
print("满页 底栏 y:", summ([r['b'] for r in full]))

g_bad = [r for r in full if r['g'] is not None and r['g'] > 1.0]
print(f"\n满页 gap>1.0cm(未贴底)的页: {len(g_bad)}")
if g_bad[:12]:
    print("  gap 最差:", [(r['pg'], round(r['g'],2)) for r in sorted(g_bad, key=lambda x:-x['g'])[:12]])
print(f"短页(尾页)样本 {len(tail)}: 底栏 y median={st.median([r['b'] for r in tail if r['b']] or [0]):.2f}")
