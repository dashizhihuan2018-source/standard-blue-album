# -*- coding: utf-8 -*-
"""
双栏版栏边界溢出量化审计（bbox 级）——交付前必跑的标准检查项。
背景：DrawingML 段落默认 hangingPunct=1，行尾中文标点悬挂到文本框外，
导致"右缘不齐 + 每行溢出约一个全角字宽"，肉眼表现为未两端对齐/文字溢出。
python-pptx 层面的 algn=just 检查发现不了此问题，必须渲染 PDF 后用
pdftotext -bbox-layout 对每行 xMax/yMin 做几何断言。
用法: python audit_overflow.py <rendered.pdf>
退出码 0=PASS, 1=FAIL
"""
import sys, re, subprocess, tempfile, os
from xml.etree import ElementTree as ET

PDF = sys.argv[1] if len(sys.argv) > 1 else ""
if not PDF:
    sys.exit("用法: python audit_overflow.py <rendered.pdf>")
NSX = 'http://www.w3.org/1999/xhtml'
CM = 28.3465

# 与引擎 _apply_geometry(columns=2) 保持一致的几何
ML = 1.8*CM; COLW = 6.675*CM; GAP = 0.85*CM
C1L, C1R = ML, ML+COLW
C2L, C2R = ML+COLW+GAP, ML+COLW+GAP+COLW
PAGE_R = 16.0*CM          # 右版心边界
BOTTOM = 23.4*CM          # 版心底
FOOTER_Y = 23.6*CM        # 页脚区起点（其下文字属页脚，合法）
TOL = 2.5                 # pt 容差

bbox = tempfile.mktemp(suffix='.xml')
subprocess.run(['pdftotext', '-bbox-layout', PDF, bbox], check=True)
tree = ET.parse(bbox)

TOP_Y = 4.2*CM            # 版心顶（其上为页眉区，标题全宽，不做栏检查）

# 页面类型清单（引擎 build_album 生成 page_types.json）：content 与 hybrid(通栏图/表带+下方双栏)
# 页的下方双栏做栏级约束；chapter(全出血设计页)/table(全宽表格页) 不适用栏检查。
# hybrid 的通栏带为全宽，其行 xmax≈右版心，由下方 xmax 上限 guard 豁免，不会误报栏溢出。
import json, os
_HERE = os.path.dirname(os.path.abspath(__file__))
# 优先取命令行第二参；否则依次尝试 PDF 同目录、本脚本目录
_PT_CANDS = ([sys.argv[2]] if len(sys.argv) > 2 else []) + [
    os.path.join(os.path.dirname(os.path.abspath(PDF)), 'page_types.json'),
    os.path.join(_HERE, 'page_types.json'),
    os.path.join(os.path.dirname(_HERE), 'samples', 'page_types.json'),
]
PTYPES = None
for _c in _PT_CANDS:
    try:
        PTYPES = json.load(open(_c)); break
    except Exception:
        continue

x_over, y_over, wide_over = [], [], []
page_no = 0
for page in tree.iter('{%s}page' % NSX):
    page_no += 1
    _pt = PTYPES[page_no-1] if PTYPES and 1 <= page_no <= len(PTYPES) else {'t':'content'}
    ptype = _pt['t'] if isinstance(_pt, dict) else _pt
    below_top = _pt.get('below_top') if isinstance(_pt, dict) else None
    if ptype in ('cover','blank','colophon','toc','chapter','table','back_cover'):
        continue
    if ptype == 'chapter':          # 全出血设计页：只查是否越出页面本身
        continue
    for line in page.iter('{%s}line' % NSX):
        xmin = float(line.get('xMin')); xmax = float(line.get('xMax'))
        ymin = float(line.get('yMin')); ymax = float(line.get('yMax'))
        h = ymax - ymin
        t = ''.join((w.text or '') for w in line.iter('{%s}word' % NSX))
        if not t.strip():
            continue
        if h > 40:        # 章分隔页大数字/水印/大标题等设计元素，豁免
            continue
        in_body = TOP_Y - TOL < ymin < BOTTOM + TOL   # 仅正文区做栏级检查
        # hybrid 页：通栏图/表带位于 band 区(ymin<below_top)，其行全宽，豁免栏检查；
        # 仅 band 区下方的双栏正文(ymin>=below_top)参与栏级检查。
        in_below = True
        if ptype in ('hybrid','gallery') and below_top is not None:
            in_below = ymin >= below_top*CM - TOL
        # 1) 栏内正文行不得越过本栏右缘（标点悬挂 bug 的回归检测）——content/hybrid(下方双栏)页
        if ptype in ('content','hybrid','gallery') and in_below and in_body and h < 25:
            if C1L - TOL < xmin < C1R and xmax > C1R + TOL and xmax < C2L + COLW*0.5:
                x_over.append((page_no, round(xmax-C1R, 1), t[:20]))
            elif xmin >= C2L - TOL and xmax > C2R + TOL:
                x_over.append((page_no, round(xmax-C2R, 1), t[:20]))
        # 2) 任何行不得越过右版心
        if xmax > PAGE_R + TOL + 6:
            wide_over.append((page_no, round((xmax-PAGE_R)/CM, 2), t[:20]))
        # 3) 正文不得压过版心底侵入页脚区（页脚自身 7.5-10pt, h<13）
        if ymin > BOTTOM + TOL and h >= 13 and not re.fullmatch(r'\d{2,3}', t.strip()):
            y_over.append((page_no, round(ymin/CM, 2), t[:20]))

os.unlink(bbox)
print("X越栏行:", len(x_over), x_over[:8])
print("越右版心行:", len(wide_over), wide_over[:8])
print("越版心底行:", len(y_over), y_over[:8])
ok = not x_over and not y_over and len(wide_over) == 0
print("OVERFLOW AUDIT:", "PASS" if ok else "FAIL")
sys.exit(0 if ok else 1)
