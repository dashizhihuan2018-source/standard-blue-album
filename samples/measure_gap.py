# -*- coding: utf-8 -*-
"""
标准蓝画册 —— 底栏缝隙 & 行高校准测量
=====================================
渲染 PDF -> pdftotext -bbox-layout -> 对每页每栏统计：
  - 该栏最底行 yMin（文字底边）
  - 距版心底(23.55) / 距页脚线(23.7) 的缝隙(cm)
  - 栏内行高中位数(median line spacing, cm) —— 用于校准 LINE_CM
  - 满行字数中位数 —— 用于校准 BODY_CPL
  - 左右栏底栏 yMin 差 —— 栏平衡度

用法: python measure_gap.py <rendered.pdf> [page_types.json]
"""
import sys, json, subprocess, tempfile, os, re
from xml.etree import ElementTree as ET

PDF = sys.argv[1]
PTJSON = sys.argv[2] if len(sys.argv) > 2 else None
NSX = 'http://www.w3.org/1999/xhtml'
CM = 28.3465

ML = 1.8 * CM
COLW = 6.675 * CM
GAP = 0.85 * CM
C1L, C1R = ML, ML + COLW
C2L, C2R = ML + COLW + GAP, ML + COLW + GAP + COLW
DIV = C1L + COLW + 0.42 * CM       # 栏分隔线 x
TOP = 4.2 * CM
BOTTOM = 23.55 * CM                # 版心底（文本框底目标）
FOOTER = 23.7 * CM                 # 页脚线
TOL = 2.5

bbox = tempfile.mktemp(suffix='.xml')
subprocess.run(['pdftotext', '-bbox-layout', PDF, bbox], check=True)
tree = ET.parse(bbox)

PT = None
if PTJSON and os.path.exists(PTJSON):
    try:
        PT = json.load(open(PTJSON))
    except Exception:
        PT = None

def is_cjk(ch):
    return ord(ch) > 0x2E80

page_no = 0
grand = []
for page in tree.iter('{%s}page' % NSX):
    page_no += 1
    if page_no <= 2:
        continue
    _pt = PT[page_no - 1] if PT and page_no <= len(PT) else {'t': 'content'}
    ptype = _pt['t'] if isinstance(_pt, dict) else _pt
    if ptype == 'chapter':
        continue
    below_top = _pt.get('below_top') if isinstance(_pt, dict) else None

    recs = []   # (col, ymin, ymax, text)
    for line in page.iter('{%s}line' % NSX):
        xmin = float(line.get('xMin')); xmax = float(line.get('xMax'))
        ymin = float(line.get('yMin')); ymax = float(line.get('yMax'))
        h = ymax - ymin
        t = ''.join((w.text or '') for w in line.iter('{%s}word' % NSX))
        if not t.strip():
            continue
        if h > 40:                 # 大数字/水印/章名，豁免
            continue
        if ptype == 'hybrid' and below_top is not None and ymin < below_top * CM - TOL:
            continue
        if not (TOP - TOL < ymin < FOOTER + TOL):
            continue
        if xmin < DIV:
            col = 1
        elif xmin >= C2L - TOL:
            col = 2
        else:
            continue
        recs.append((col, ymin, ymax, t))

    colinfo = {}
    for c in (1, 2):
        ls = [r for r in recs if r[0] == c]
        if not ls:
            colinfo[c] = None
            continue
        maxy = max(r[1] for r in ls)
        miny = min(r[1] for r in ls)
        ys = sorted(r[1] for r in ls)
        diffs = [ys[i + 1] - ys[i] for i in range(len(ys) - 1)]
        diffs = [d for d in diffs if 0 < d < 3 * CM]
        med = sorted(diffs)[len(diffs) // 2] if diffs else 0
        adv = (maxy - miny) / (len(ys) - 1) if len(ys) > 1 else 0.0
        # 满行字数：取宽度占满本栏(>0.85*COLW)的行，数 CJK 字符
        cjk_counts = []
        for r in ls:
            if (r[2] - r[1]) < 25:          # 行高正常
                cw = r[2] - r[1]
            # 用 x 跨度估计是否满行
        # 简化：统计每行的 CJK 字符数，取中位数作为参考
        ncs = [sum(1 for ch in r[3] if is_cjk(ch)) for r in ls]
        med_cjk = sorted(ncs)[len(ncs) // 2] if ncs else 0
        max_cjk = max(ncs) if ncs else 0
        # 用 x 跨度估计满行字符数（更稳）：取占满本栏宽度>0.9 的行的 CJK 数
        full_cjk = [n for n, r in zip(ncs, ls) if (r[2] - r[1]) >= 0.6*CM]
        full_cjk_max = max(full_cjk) if full_cjk else 0
        colinfo[c] = dict(lines=len(ls), maxy=maxy, miny=miny, adv=round(adv / CM, 4),
                          gap_b=round((BOTTOM - maxy) / CM, 2),
                          gap_f=round((FOOTER - maxy) / CM, 2), lineh=round(med / CM, 3),
                          cjk_med=med_cjk, cjk_max=max_cjk, cjk_fullmax=full_cjk_max)
    b1 = colinfo[1]['maxy'] if colinfo[1] else None
    b2 = colinfo[2]['maxy'] if colinfo[2] else None
    diff = round(abs(b1 - b2) / CM, 2) if (b1 and b2) else None
    print(f"PAGE{page_no} [{ptype}] col1={colinfo[1]}  col2={colinfo[2]}  balDiff={diff}")
    for c in (1,2):
        if colinfo[c]:
            print(f"    col{c}: lines={colinfo[c]['lines']} adv={colinfo[c].get('adv')} miny={round(colinfo[c].get('miny',0)/CM,2)} maxy={round(colinfo[c].get('maxy',0)/CM,2)} lineH={colinfo[c]['lineh']}")
    grand.append((page_no, colinfo, diff))

os.unlink(bbox)
print("\n--- summary ---")
for page_no, ci, diff in grand:
    g1 = ci[1]['gap_f'] if ci[1] else None
    g2 = ci[2]['gap_f'] if ci[2] else None
    lh1 = ci[1]['lineh'] if ci[1] else None
    lh2 = ci[2]['lineh'] if ci[2] else None
    print(f"PAGE{page_no}: gap_footer col1={g1} col2={g2} | lineH col1={lh1} col2={lh2} | balDiff={diff}")
