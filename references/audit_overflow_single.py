# -*- coding: utf-8 -*-
"""单栏版栏边界溢出量化审计（bbox 级）——B5/A4 自适应。
单栏无中缝：整页为一个通栏文本块，版心边界按 PDF 页面宽高动态计算。
故只需三项几何断言：
  1) 正文行不得越过右版心（标点悬挂 bug 的回归检测，单栏右缘即版心右界）；
  2) 任何行不得越过右版心（与 1 同源，统一上报）；
  3) 正文不得压过版心底(23.4cm)侵入页脚区。
用法: python audit_overflow_single.py <rendered.pdf> [page_types_single.json]
退出码 0=PASS, 1=FAIL
"""
import sys, re, subprocess, tempfile, os, json
from xml.etree import ElementTree as ET

PDF = sys.argv[1] if len(sys.argv) > 1 else ""
PTFILE = sys.argv[2] if len(sys.argv) > 2 else os.path.join(os.path.dirname(os.path.abspath(__file__)), 'page_types_single.json')
NSX = 'http://www.w3.org/1999/xhtml'
CM = 28.3465

ML = 1.8*CM
TOP_Y = 4.2*CM            # 版心顶
TOL = 2.5                 # pt 容差

bbox = tempfile.mktemp(suffix='.xml')
subprocess.run(['pdftotext', '-bbox-layout', PDF, bbox], check=True)
tree = ET.parse(bbox)

_HERE = os.path.dirname(os.path.abspath(__file__))
PTYPES = json.load(open(PTFILE)) if os.path.exists(PTFILE) else None

x_over, y_over = [], []
page_no = 0
for page in tree.iter('{%s}page' % NSX):
    page_no += 1
    page_w = float(page.get('width'))
    page_h = float(page.get('height'))
    page_r = page_w - 1.6*CM
    bottom = page_h - 1.60*CM
    _pt = PTYPES[page_no-1] if PTYPES and 1 <= page_no <= len(PTYPES) else {'t':'content'}
    ptype = _pt['t'] if isinstance(_pt, dict) else _pt
    below_top = _pt.get('below_top') if isinstance(_pt, dict) else None
    if ptype in ('cover','blank','colophon','toc','chapter','table','back_cover'):
        continue
    for line in page.iter('{%s}line' % NSX):
        xmin = float(line.get('xMin')); xmax = float(line.get('xMax'))
        ymin = float(line.get('yMin')); ymax = float(line.get('yMax'))
        h = ymax - ymin
        t = ''.join((w.text or '') for w in line.iter('{%s}word' % NSX))
        if not t.strip():
            continue
        if h > 40:        # 设计元素/水印
            continue
        in_body = TOP_Y - TOL < ymin < bottom + TOL
        in_below = True
        if ptype in ('hybrid','gallery') and below_top is not None:
            in_below = ymin >= below_top*CM - TOL
        # 1) + 2) 单栏右缘即版心右界：正文行不得越过右侧
        if ptype in ('content','hybrid','gallery') and in_below and in_body and h < 25:
            if xmax > page_r + TOL:
                x_over.append((page_no, round((xmax-page_r)/CM, 2), t[:20]))
        # 3) 正文不得压过版心底侵入页脚区
        if ymin > bottom + TOL and h >= 13 and not re.fullmatch(r'\d{2,3}', t.strip()):
            y_over.append((page_no, round(ymin/CM, 2), t[:20]))

os.unlink(bbox)
print("越右版心/栏右缘行(X):", len(x_over), x_over[:8])
print("越版心底行(Y):", len(y_over), y_over[:8])
ok = not x_over and not y_over
print("SINGLE-COLUMN OVERFLOW AUDIT:", "PASS" if ok else "FAIL")
sys.exit(0 if ok else 1)
