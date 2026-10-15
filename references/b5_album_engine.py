# -*- coding: utf-8 -*-
"""
标准蓝画册排版引擎（B5 / A4，单双栏自适应）
==========================================
面向标准蓝视觉配色的通用画册排版引擎（python-pptx 实现）。
提供设计令牌、行粒度分页、九类页面渲染、自动页码域与图文表混排能力。

所有项目相关文案（书名、英文名、版次、标语、机构名、密级）均为可配置令牌，
引擎自身不含任何具体项目内容。

用法：
    from b5_album_engine import build_album
    build_album(blocks=..., frontmatter=..., toc_names=...,
                ch_en=..., out="out.pptx",
                media_dir="sources/media", cover_img="sources/media/cover.png",
                options={"book_title": "××画册", "slogan": ("匠心筑造", "品质焕新")})

blocks 结构（list[dict]）：
    {"kind":"chapter",  "text":"一、总则"}
    {"kind":"heading",  "level":2|3|4, "text":"2.1 策划原则"}
    {"kind":"para",     "text":"正文……"}
    {"kind":"image",    "path":"./x.png", "aspect":1.6, "caption":"图1 …"}
    {"kind":"table",    "rows":[[{"text":"表头","images":[]}, ...], ...]}

依赖：python-pptx（pip install python-pptx）
"""
import os, re, math, json, contextlib
from pptx import Presentation
from pptx.util import Cm, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE
from pptx.oxml.ns import qn
from pptx.oxml import parse_xml

# ============================================================
# 1. 设计令牌（Design Tokens）—— 与最终版 1:1 复刻
# ============================================================
BLUE    = RGBColor(0x00, 0x3D, 0xA5)
BLUE_D  = RGBColor(0x00, 0x2A, 0x73)
BLUE_LT = RGBColor(0xE9, 0xEF, 0xF8)
INK     = RGBColor(0x1A, 0x1A, 0x1A)
GRAY    = RGBColor(0x6E, 0x6E, 0x6E)
RULE    = RGBColor(0xC9, 0xD4, 0xE3)
WHITE   = RGBColor(0xFF, 0xFF, 0xFF)

TITLE_F = "Source Han Serif CN"      # 思源宋体（章节分隔页大标题）
HEI     = "Source Han Serif CN"                  # 黑体（标题/表头/眉脚/序号/标签）
SONG    = "Source Han Serif CN"                # 宋体：图注、表格单元等非正文角色
BODY    = "Source Han Serif CN"      # 正文默认思源宋体；单双栏一致，可由 body_font 显式覆盖
COVER_F = "Source Han Serif CN"      # 思源宋体：封面主标 / 目录章名 / 编委会 / 封尾标语（人工版封面区字体）

# ---- 品牌文案令牌 ----------------------------------------------------------
# 全部可由 build_album(options=...) 覆盖。渲染函数内严禁写死任何具体项目的
# 书名、标语、单位名——否则会把上一个项目的文案带进新交付物。
BRAND_LINE    = "STANDARD BLUE ALBUM"   # 封面品牌行
ORG_NAME      = "示例机构"                  # 页脚 / 封面 / 封尾落款
DOC_CLASS     = "公开示例"                              # 密级标注（可置空不渲染）
BOOK_TITLE    = "标准蓝画册"                              # 画册名：页眉书名行、封面主标题、编委会眉题
BOOK_SUBTITLE = ""                                      # 封面中文副标题
BOOK_TITLE_EN = ""                                      # 英文书名（封面副行，空则不渲染）
BOOK_EDITION  = ""                                      # 版次/批次标注（封面，空则不渲染）
SLOGAN        = ()                                      # 封尾标语，如 ("匠心筑造", "品质焕新")；空则回退书名

# B5 几何（cm）
PW, PH        = 17.6, 25.0
ML            = 1.8
CONTENT_W     = 14.2
COL_W         = 14.2
COL1_X        = 1.8
COL2_X        = 9.32
DIV_X         = 8.89
GAP           = 0.85             # 栏间距 (= COL2_X - COL1_X - COL_W)
TOP           = 4.2
CONTENT_H     = 19.2
CONTENT_BOTTOM= 23.62    # 正文栏底部目标 y：紧贴 23.7 页脚线（文本底约 23.6，留 0.1 安全缝）
FOOTER_LINE_Y = 23.70
FOOTER_TEXT_Y = 23.85
LINE_CM       = 0.71     # 实测 PDF 双栏正文行“首→尾”有效推进（含 CJK 行框与标题段距）≈0.71–0.74cm；取 0.71 在“贴底 27 行”与“标题不越页脚”间取得平衡
BODY_CPL      = 38        # 单栏满宽正文行物理容量约 38 CJK 字/行（COL_W=14.2cm @10.5pt）。
                           # 仅用于 _split_para_to 的“每物理行装多少字”预算（真实渲染容量）。
EST_CPL       = 37.0      # 行数估算用的悲观容量：首行缩进 + 两端对齐下真实每行仅容 ≤38 字，且段末常因
                           # 标点/justify 方差多挤 1 行。用 37 估行数 ⇒ 估算行数 ≥ 真实行数 ⇒ 栏高
                           # 估算 ≥ 真实渲染高度 ⇒ 整栏绝不越过 23.7 页脚线。密度损失仅 ~0.3cm/栏。
SAFE          = 0.35      # 栏高安全余量(cm)：吸收 CJK 行框/两端对齐方差。LINE_CM 虽按 0.71 校准，
                           # 但整栏 27 行累计真实高度仍会比模型多 ~0.2cm，留 0.35 确保末行绝不越 23.7 页脚线。
CAP_CM        = (CONTENT_BOTTOM - TOP) - SAFE   # 每栏可用高度（≈19.07cm ≈ 27 行），单/双栏同高
# 填充补偿：成本模型对每段的固定开销（段后距 +0.14cm、首行缩进补偿 +2 字、+0.04cm 安全缝）按
# 段落数累积，使估算栏高高于真实渲染。单栏每栏容纳的段落数 ≈ 双栏的 2 倍 ⇒ 单栏高估约 1.87cm，
# 双栏约 0.5cm。故 FILL_BONUS 随栏数自适应：双栏=0.5（栏底逼近 23.27cm，溢出审计 PASS）、
# 单栏=1.8（抵消更高的段数累积高估，栏底逼近 23.2cm）。值由 _apply_geometry 按栏数覆写，
# 模块级默认 1.2 对应单栏兜底。
FILL_BONUS    = 1.2
CHUNK_TARGET  = 3.6       # （保留参数）历史用途；实际填充改由 place_para 自适应按行拆分

# 混排页
BAND_GAP      = 0.40
IMG_BAND_MAX  = 0.62
TBL_BAND_MAX  = 0.80
MIN_BELOW_H   = 5.5    # 仅当表带过高、下方不足 ≥5.5cm 双栏区时才独立成表页（图始终混排）
# 正文段后距设为 0：段落间仅靠首行缩进区分（密集书籍样式），换得栏底贴齐 23.7 页脚线；
# 标题保留少量段后距以维持视觉分组。
SA_BODY = 0.0; SA_H2 = 0.10; SA_H3 = 0.08; SA_H4 = 0.06

MR     = 1.6            # 右页边（page_size 重算 CONTENT_W 用）
NCOLS  = 1              # 1=单栏（默认）, 2=双栏（可由 build_album(options=...) 覆盖）

NUM_LEAD = re.compile(r'^(\d+．[^：:]{1,30}[：:])(.*)$', re.S)
HEAD_NUM = re.compile(r'^(\d+(?:\.\d+)*)(?=\S)')

# ============================================================
# 1b. 需求可选项（build_album(options=...) 生效）
# ============================================================
def _apply_geometry(n=1, page_size='B5'):
    """按栏数与页面规格重算几何全局；返回旧值字典供还原。

    注意：BODY_CPL（每行汉字容量）是经 LibreOffice 实测校准的保守值，直接决定按行拆分的
    行数计量。双栏(COL_W=6.67)实测=18；单栏(COL_W=14.2，满宽)按宽比 14.2/6.67≈2.13 推算≈38
    （new_tb 内边距已清零，可用宽=COL_W，字数随宽线性，故可行）。EST_CPL 始终取 BODY_CPL-0.5
    维持悲观偏移，确保行数估算 ≥ 真实 ⇒ 整栏绝不越 23.7 页脚线。此处随栏数重算 BODY_CPL/EST_CPL
    与纯几何尺寸，二者一并存入 saved 字典、由 _options_scope 在 build 结束后还原。
    """
    global NCOLS, COL_W, COL1_X, COL2_X, DIV_X, PW, PH, CONTENT_W, BODY_CPL, EST_CPL, FILL_BONUS
    global CONTENT_H, CONTENT_BOTTOM, CAP_CM, FOOTER_LINE_Y, FOOTER_TEXT_Y
    saved = dict(NCOLS=NCOLS, COL_W=COL_W, COL1_X=COL1_X, COL2_X=COL2_X,
                 DIV_X=DIV_X, PW=PW, PH=PH, CONTENT_W=CONTENT_W,
                 BODY_CPL=BODY_CPL, EST_CPL=EST_CPL, FILL_BONUS=FILL_BONUS,
                 CONTENT_H=CONTENT_H, CONTENT_BOTTOM=CONTENT_BOTTOM, CAP_CM=CAP_CM,
                 FOOTER_LINE_Y=FOOTER_LINE_Y, FOOTER_TEXT_Y=FOOTER_TEXT_Y)
    if isinstance(page_size, str):
        if page_size.upper() == 'A4': PW, PH = 21.0, 29.7
        else: PW, PH = 17.6, 25.0                      # B5（默认）
    elif isinstance(page_size, (tuple, list)) and len(page_size) == 2:
        PW, PH = float(page_size[0]), float(page_size[1])
    CONTENT_W = PW - ML - MR
    # A4 不只改变纸张尺寸，同步扩展正文底界与页脚坐标。
    # 保留与 B5 一致的 1.30cm 页脚距离和 0.08cm 安全缝。
    if abs(PH - 29.7) < 0.01:
        FOOTER_LINE_Y = PH - 1.30
        FOOTER_TEXT_Y = PH - 1.15
        CONTENT_BOTTOM = FOOTER_LINE_Y - 0.08
    else:
        FOOTER_LINE_Y = 23.70
        FOOTER_TEXT_Y = 23.85
        CONTENT_BOTTOM = 23.62
    CONTENT_H = CONTENT_BOTTOM - TOP
    CAP_CM = CONTENT_H - SAFE
    n = 1 if n == 1 else 2
    NCOLS = n
    if n == 1:
        COL_W = CONTENT_W
        COL1_X = ML
        COL2_X = None
        DIV_X = None
        # 单栏满宽(14.2cm)每行约 38 汉字：双栏 6.67cm=18，按宽比 14.2/6.67≈2.13 推算
        # （new_tb 内边距已清零，可用宽=COL_W，故字数随宽线性）。EST_CPL 取 BODY_CPL-0.5
        # 维持与双栏一致的悲观偏移，确保行数估算 ≥ 真实 ⇒ 整栏绝不越 23.7 页脚线。
        BODY_CPL = 38
        EST_CPL  = 37
        FILL_BONUS = 1.2          # 单栏每栏段落数≈双栏2倍，段数累积高估~1.87cm；取1.2在"贴底"与
                                  # "末行绝不越23.7页脚线"间取得平衡（1.4时page77拉丁标题低估仍微越界0.09cm，
                                  # 1.8时5行越界0.28cm）。1.2 实测零越界、中位gap≈1.2cm。
    else:
        COL_W = (CONTENT_W - GAP) / 2
        COL1_X = ML
        COL2_X = ML + COL_W + GAP
        DIV_X = COL1_X + COL_W + 0.42
        BODY_CPL = 18
        EST_CPL  = 17.5
        FILL_BONUS = 0.5   # 双栏：段数累积高估~0.5cm，栏底逼近23.27；与 LINE_CM=0.71 组合经实测零越界
    return saved

@contextlib.contextmanager
def _options_scope(options):
    """进入 build 前应用字体/栏式/页面规格覆盖，退出时还原全局。"""
    saved = {}
    for key, gname in (('body_font', 'BODY'), ('title_font', 'TITLE_F'),
                       ('cover_font', 'COVER_F'), ('sans_font', 'HEI'), ('song_font', 'SONG'),
                       ('book_title', 'BOOK_TITLE'), ('book_subtitle', 'BOOK_SUBTITLE'),
                       ('book_title_en', 'BOOK_TITLE_EN'),
                       ('edition', 'BOOK_EDITION'), ('brand_line', 'BRAND_LINE'),
                       ('org_name', 'ORG_NAME'), ('doc_class', 'DOC_CLASS'),
                       ('slogan', 'SLOGAN')):
        if key in options:
            saved[gname] = globals()[gname]; globals()[gname] = options[key]
    if 'columns' in options or 'page_size' in options:
        saved.update(_apply_geometry(options.get('columns', NCOLS),
                                     options.get('page_size', (PW, PH))))
    try:
        yield
    finally:
        for k, v in saved.items():
            globals()[k] = v

# ============================================================
# 2. 基础工具
# ============================================================
def cn_to_int(s):
    m = {'一':1,'二':2,'三':3,'四':4,'五':5,'六':6,'七':7,'八':8,'九':9}
    if s == '十': return 10
    if s.startswith('十'): return 10 + m.get(s[1:], 0)
    if s.endswith('十'): return m.get(s[0], 0) * 10
    return m.get(s, 0)

def text_units(s):
    """粗略字数：CJK=1，西文/数字=0.5。用于行容量估算。"""
    return sum(0 if ch.isspace() else (1.0 if ord(ch) > 0x2E80 else 0.5) for ch in s)

def _is_cjk(ch):
    return ord(ch) > 0x2E80

_PN = '，。：；、！？…—）】」』》'

def _split_para_to(text, space_cm, first):
    """把正文段落按『行』自适应拆出恰好填满 space_cm 的块，消除封页时的栏内留白。

    返回 (chunk_text, rest_text, cost_cm)；若 space_cm 连一行都放不下返回 None。
    - 行数由几何空间决定（最可靠）：最多 floor(space/LINE_CM) 行。
    - 首块首行缩进约 2 字符宽，行数按 (u+2)/CPL 计（与 para_cost 一致）；续块按 u/CPL。
    - 段后距用渲染侧真实 Pt(4)=0.14cm。
    - 行首/行尾标点防护：续块不得以标点开头（并回本块末尾与前置字粘连）；若本块退化成纯标点
      （如“。”）则向后借一字，杜绝『一个字+标点独占一行 / 标点独占行首』低级 bug。"""
    if not text:
        return ("", "", 0.0)
    def _take(avail):
        lines = int(avail // LINE_CM)
        if lines < 1:
            return None
        safe_cpl = BODY_CPL
        budget = lines * safe_cpl - (2 if first else 0)
        if budget < 1:
            budget = lines * safe_cpl
        taken, u = 0, 0.0
        for ch in text:
            cu = 0.5 if not _is_cjk(ch) else 1.0
            if taken > 0 and u + cu > budget:
                break
            u += cu; taken += 1
        if taken == 0:
            taken = 1
        return taken
    def _finalize(avail):
        taken = _take(avail)
        if taken is None:
            return None
        chunk = text[:taken]; rest = text[taken:]
        # 续块若以标点开头 → 并回本块末尾（与前置字粘连，避免标点独占行首/独占一行）
        if rest and rest[0] in _PN and len(chunk) < int(avail // LINE_CM) * BODY_CPL:
            chunk += rest[0]; rest = rest[1:]
        # 本块退化成纯标点（如“。”）→ 向后借一个字
        if chunk and all(c in _PN for c in chunk) and rest:
            chunk += rest[0]; rest = rest[1:]
        u = sum(0.5 if not _is_cjk(c) else 1.0 for c in chunk)
        is_last = (rest == "")
        max_lines = max(1, math.ceil((u + 2) / EST_CPL)) if first else max(1, math.ceil(u / EST_CPL))
        cost = max_lines * LINE_CM + (0.14 if is_last else 0.0)
        return chunk, rest, cost, is_last
    res = _finalize(space_cm)
    if res is None:
        return None
    chunk, rest, cost, is_last = res
    if cost > space_cm + 1e-6:                 # 悲观估计仍溢出 → 减一行重试
        res = _finalize(space_cm - LINE_CM)
        if res is None:
            return None
        chunk, rest, cost, is_last = res
    return (chunk, rest, cost)

# 中文行首禁止出现的标点（悬挂/前置标点）。在渲染前用 U+2060 WORD JOINER
# 把这些标点前一个字“粘”在一起，避免 PowerPoint 把标点单独挤到下一行行首。
_HANG_PUNCT = re.compile(r'([\u4e00-\u9fff])([，。；：、！？；:\)\]\}》」』])')

def cjk_join_punct(s):
    """将 CJK 字符与其后的前置/悬挂标点粘在一起，减少『，维护』类行首标点 bug。"""
    wj = '\u2060'
    return _HANG_PUNCT.sub(lambda m: m.group(1) + wj + m.group(2), s)

def set_cjk(run, name):
    """关键：同时写 name + a:ea + a:cs，否则 CJK 落默认字体。"""
    run.font.name = name
    rPr = run._r.get_or_add_rPr()
    for tag in ('a:ea', 'a:cs'):
        el = rPr.find(qn(tag))
        if el is None:
            el = rPr.makeelement(qn(tag), {}); rPr.append(el)
        el.set('typeface', name)

def add_rect(slide, x, y, w, h, fill, line=None, line_w=None, rounded=False):
    sp = slide.shapes.add_shape(
        MSO_SHAPE.ROUNDED_RECTANGLE if rounded else MSO_SHAPE.RECTANGLE,
        Cm(x), Cm(y), Cm(w), Cm(h))
    sp.shadow.inherit = False
    if fill is None: sp.fill.background()
    else: sp.fill.solid(); sp.fill.fore_color.rgb = fill
    if line is None: sp.line.fill.background()
    else: sp.line.color.rgb = line; sp.line.width = line_w or Pt(0.75)
    return sp

def new_tb(slide, x, y, w, h, anchor=MSO_ANCHOR.TOP, wrap=True):
    tb = slide.shapes.add_textbox(Cm(x), Cm(y), Cm(w), Cm(h))
    tf = tb.text_frame; tf.word_wrap = wrap; tf.vertical_anchor = anchor
    tf.margin_left = 0; tf.margin_right = 0; tf.margin_top = 0; tf.margin_bottom = 0
    return tf

def para(tf, first, runs, align=None, ls=None, sb=None, sa=None, fi=None):
    p = tf.paragraphs[0] if first else tf.add_paragraph()
    if align is not None: p.alignment = align
    if ls is not None: p.line_spacing = ls
    if sb is not None: p.space_before = sb
    if sa is not None: p.space_after = sa
    # 硬规则①：关闭标点悬挂，否则行尾中文标点溢出文本框。
    pPr = p._p.get_or_add_pPr(); pPr.set('hangingPunct', '0')
    if fi is not None: pPr.set('indent', str(int(fi)))
    for r in runs:
        run = p.add_run(); run.text = r['t']
        run.font.size = r.get('sz', Pt(10.5))
        run.font.bold = r.get('b', False)
        run.font.color.rgb = r.get('c', INK)
        set_cjk(run, r.get('f', SONG))
    return p

def simple_text(slide, x, y, w, h, t, sz, b, c, f,
                align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.TOP):
    tf = new_tb(slide, x, y, w, h, anchor)
    para(tf, True, [{'t': t, 'sz': Pt(sz), 'b': b, 'c': c, 'f': f}], align=align)

def fit_size(text, max_w_cm, max_pt, min_pt=9):
    n = max(1, text_units(text))
    fs = max_w_cm / (n * 0.0353 * 1.06)
    return max(min_pt, min(max_pt, fs))

def head_disp(t):
    return HEAD_NUM.sub(lambda m: m.group(1) + ' ', t, count=1)

# ============================================================
# 3. 成本模型（cm 计量，含段后距）—— 硬规则②
# ============================================================
def heading_cost_cm(level, text):
    fs = {2: 12.5, 3: 11.5, 4: 10.5}[level]
    cpl = max(6, int(COL_W / (fs * 0.0353)))
    lines = max(1, math.ceil(text_units(text) / cpl))
    sb = {2: 0.35, 3: 0.28, 4: 0.21}.get(level, 0.21)
    sa = {2: SA_H2, 3: SA_H3, 4: SA_H4}.get(level, SA_H4)
    return lines * LINE_CM + sb + sa + 0.04        # +0.04 极小安全余量，防标题越界

def hybrid_lead_height(page):
    """在版心内为表前标题预留整行高度，返回标题和表格间距的总和(cm)。"""
    title = page.get('lead_heading', '')
    if not title:
        return 0.0
    fs = {2: 12.5, 3: 11.5, 4: 10.5}[page.get('lead_level', 2)]
    lines = sum(max(1, math.ceil(text_units(part) / max(1, CONTENT_W / (fs * 0.0353 * 1.1))))
                for part in title.split('\n'))
    return lines * fs * 0.0353 * 1.35 + 0.22

def para_cost(text):
    """精确行数估算（EST_CPL 悲观容量，行数 ≥ 真实 ⇒ 栏高估算 ≥ 渲染高度，杜绝越页脚）。
    +2 补偿首行缩进占用的约 2 字符宽度；line_spacing=1.45/1.55 已含在 LINE_CM 中，
    故行数×行高 + 段后距即为真实高度（且偏安全一侧）。"""
    return max(1, math.ceil((text_units(text) + 2) / EST_CPL))

def para_cost_cm(text):
    return para_cost(text) * LINE_CM + SA_BODY + 0.04   # +0.04cm 极小安全余量，防段落末行溢出

def _cont_para_cost(t):
    """正文续段成本：无首行缩进、无段后距（由同段下一续段或末段承担）。"""
    lines = max(1, math.ceil(text_units(t) / EST_CPL))
    return lines * LINE_CM

def split_long_para(b, max_cm=None, target_cm=None):
    """拆分块。
    - max_cm: 单块上限（混排页用 cap_cm；默认 CAP_CM）。
    - target_cm: 目标块高（content 页用 CHUNK_TARGET），即使单块能放下也按 target 切，
                 让 balance 有细粒度材料，实现左右栏底栏对齐。
    返回每个子块带 'cost'/'indent'/'continued'，渲染时据此决定是否首行缩进/段后距。
    """
    if max_cm is None: max_cm = CAP_CM
    kind = b['kind']; lvl = b.get('level')
    text = b['text']
    # 首段成本（含首行缩进+段后距）；续段成本（无缩进无段后距）；末段成本（无缩进+段后距）
    def first_cm(t):
        return heading_cost_cm(lvl, t) if kind == 'heading' else para_cost_cm(t)
    def cont_cm(t):
        if kind == 'heading': return heading_cost_cm(lvl, t)
        return _cont_para_cost(t)
    def last_cm(t):
        if kind == 'heading': return heading_cost_cm(lvl, t)
        return _cont_para_cost(t) + SA_BODY
    limit = target_cm if target_cm is not None else max_cm
    # 单块即可放下且无需按 target 细分
    if first_cm(text) <= max_cm and (target_cm is None or first_cm(text) <= target_cm):
        return [{'kind': kind, 'text': text, 'level': lvl,
                 'cost': first_cm(text), 'indent': True, 'continued': False}]
    # 按标点拆句，尽量在句读处断开，避免 mid-sentence 的硬切
    parts = re.split(r'(?<=[，,。；;！？!?\n])', text)
    raw = []; cur = ''
    for p in parts:
        if not p: continue
        cand = cur + p
        cost = first_cm(cand) if not raw else cont_cm(cand)
        if cur and cost > limit:
            raw.append(cur); cur = p
        else:
            cur = cand
    if cur: raw.append(cur)
    # 对仍超限的块按字符数硬切
    chunks = []
    for i, ch in enumerate(raw):
        cm_fn = first_cm if i == 0 else cont_cm
        if cm_fn(ch) <= limit:
            chunks.append(ch)
        else:
            step = max(1, int(limit / LINE_CM) * BODY_CPL - 6)
            for j in range(0, len(ch), step):
                chunks.append(ch[j:j + step])
    # 合并过短残留块，避免『一个字+标点独占一行』
    merged = []
    for i, ch in enumerate(chunks):
        if merged and len(ch.strip()) <= 3:
            prev = merged[-1]
            prev_text = prev['text']
            # 尝试合并到前一块；用前一块对应的成本函数判断
            cm_prev = first_cm if prev.get('idx', 0) == 0 else cont_cm
            if cm_prev(prev_text + ch) <= limit:
                prev['text'] = prev_text + ch
                continue
        merged.append({'text': ch, 'idx': i})
    # 生成带成本与渲染标记的最终块
    out = []
    n = len(merged)
    for i, m in enumerate(merged):
        is_first = (i == 0)
        is_last = (i == n - 1)
        indent = is_first
        continued = not is_last   # 除末段外均为续段（无段后距）
        if is_first and is_last:
            cost = first_cm(m['text'])
        elif is_first:
            cost = first_cm(m['text']) - SA_BODY   # 首段承担缩进，但 sa 由末段承担
        elif is_last:
            cost = last_cm(m['text'])
        else:
            cost = cont_cm(m['text'])
        out.append({'kind': kind, 'text': m['text'], 'level': lvl,
                    'cost': cost, 'indent': indent, 'continued': continued})
    return out

def image_cost(aspect, caption):
    if not aspect or aspect <= 0: aspect = 1.6
    h = min(COL_W / aspect, CONTENT_H * 0.9)
    return h + (0.7 if caption else 0) + 0.2

def cell_h(cell, w_cm):
    txt = cell.get('text', ''); imgs = cell.get('images', [])
    cpl = max(2, int(w_cm / 0.33))
    lines = max(1, math.ceil(text_units(txt) / cpl)) if txt else 0
    h = lines * 0.38 + 0.12
    for im in imgs:
        asp = im.get('aspect') or 1.6
        h += min((w_cm - 0.16) / asp, 5.0) + 0.10
    return h

def table_full_h(rows):
    ncols = max((len(r) for r in rows), default=1); cw = CONTENT_W / ncols
    return sum(max((cell_h(row[ci] if ci < len(row) else {'text': '', 'images': []}, cw)
                    for ci in range(ncols)), default=0.5) for row in rows)

def table_rows_height(rows):
    """计算一组表格行的总渲染高度（cm），用于独立表页垂直居中。"""
    ncols = max((len(r) for r in rows), default=1); cw = CONTENT_W / ncols
    return sum(max((cell_h(row[ci] if ci < len(row) else {'text': '', 'images': []}, cw)
                    for ci in range(ncols)), default=0.5) for row in rows)

# ============================================================
# 4. 分页（含混排页逻辑）
# ============================================================
def paginate(blocks):
    """按『行粒度填满双栏』排版：每个正文栏都精确填充到 CAP_CM（≈栏高），
    使左右两栏底栏都贴齐 23.7 页脚线且左右等高对齐；章节收尾不足两栏时两栏等高分栏。

    修复底栏未对齐 bug 的根因：旧 balance() 只能在『段落边界』上分割，段落粒度粗（常 ~8cm），
    两栏高度差 2–4cm；且每页内容常不足 2×CAP 而留白 4cm+。新方案用 _split_para_to 按『行』
    把段落拆到恰好填满当前栏剩余高度，再续排下一栏/下一页——两栏均达 CAP_CM ⇒ 底栏同时
    贴底且左右对齐；续排段取消首行缩进与段后距，视觉连续。
    """
    pages = []; cur_ch = ""; cur_no = 0

    # ===== 混排页（图/表通栏置顶 + 下方双栏）保持原实现 =====
    def balance(pg, cap=CAP_CM):
        """双栏底栏对齐：把本页/本区有序块按完整段落重分配为左右两栏（混排页用）。"""
        if NCOLS != 2 or pg.get('type') not in ('content', 'hybrid'): return
        items = pg.get('items', [])
        if not items:
            pg['cols'][0] = []; pg['cols'][1] = []; return
        orig = [list(pg['cols'][0]), list(pg['cols'][1])]
        costs = [it['cost'] for it in items]; n = len(costs)
        total = sum(costs)
        if total < cap * 0.5:
            pg['cols'][0] = list(items); pg['cols'][1] = []; return
        pref = [0.0] * (n + 1)
        for i in range(n): pref[i + 1] = pref[i] + costs[i]
        best_k, best_key = None, None
        for k in range(0, n + 1):
            L = pref[k]; R = total - L
            if L > cap + 1e-6 or R > cap + 1e-9: continue
            d = abs(L - R)
            if k > 0 and items[k - 1]['block']['kind'] == 'heading':
                d += 0.3
            key = (d, -k)
            if best_key is None or key < best_key:
                best_key = key; best_k = k
        if best_k is None:
            pg['cols'][0] = orig[0]; pg['cols'][1] = orig[1]; return
        L = pref[best_k]; R = total - L
        if min(L, R) < 2.0 and total <= cap + 1e-6:
            best_k = n
        pg['cols'][0] = list(items[:best_k]); pg['cols'][1] = list(items[best_k:])

    def fill_below(band_kind, band, band_h, i, lead_heading=None):
        """通栏图/表置顶 + 下方栏（单/双栏，高度自动缩短）。放不下即停、交还主循环。"""
        nonlocal cur_ch, cur_no
        ptype = 'gallery' if band_kind == 'gallery' else 'hybrid'
        pg = {'type': ptype, 'band_kind': band_kind, 'band': band, 'band_h': band_h,
              'chapter': cur_ch, 'no': cur_no, 'cols': [[] for _ in range(NCOLS)], 'items': []}
        if lead_heading:
            pg.update(lead_heading=lead_heading['text'], lead_level=lead_heading['level'])
        cap_cm = max(0.0, CONTENT_BOTTOM - TOP - hybrid_lead_height(pg) - band_h - BAND_GAP - SAFE)
        pages.append(pg); r2 = [cap_cm] * NCOLS
        N = len(blocks)
        while i < N:
            b = blocks[i]; k = b['kind']
            if k not in ('heading', 'para'): break
            subs = split_long_para(b, max_cm=cap_cm); placed = 0
            for sub in subs:
                cost = sub.get('cost',
                               (heading_cost_cm(sub['level'], sub['text']) if k == 'heading'
                                else para_cost_cm(sub['text'])))
                c = None
                for cc in range(NCOLS):
                    if r2[cc] >= cost: c = cc; break
                if c is None: break
                pg['cols'][c].append({'block': sub, 'cost': cost})
                pg['items'].append({'block': sub, 'cost': cost})
                r2[c] -= cost; placed += 1
            if placed < len(subs): break
            i += 1
        balance(pg, cap=cap_cm)
        return i

    # ===== 新：内容页按行粒度填满两栏 =====
    def _fill_col(stream, si, cap):
        """从 stream[si] 起按行粒度填满一个高度 cap 的栏，返回 (items, 新si)。
        - 标题：整体放入；本栏剩余高度放不下该标题则留到下一栏（不溢出）。
        - 段落：_split_para_to 按行拆出恰好填满剩余高度的块；未完段落原地更新剩余 text 续排。"""
        items = []; rem = cap; n = len(stream)
        while si < n:
            node = stream[si]
            if node['type'] == 'head':
                b = node['block']; lvl = b['level']; tc = b['text']
                c = heading_cost_cm(lvl, tc)
                if rem < c - 1e-6:                 # 本栏放不下该标题 → 留到下一栏
                    break
                items.append({'block': {'kind': 'heading', 'level': lvl, 'text': tc}, 'cost': c})
                rem -= c; si += 1
            else:
                full = node['text']; first = not node.get('started', False)
                res = _split_para_to(full, rem, first)
                if res is None:                    # 本栏剩余高度连一行都放不下 → 结束本栏
                    break
                chunk, rest, cost = res
                items.append({'block': {'kind': 'para', 'text': chunk, 'level': None,
                                        'indent': first, 'continued': (rest != "")}, 'cost': cost})
                rem -= cost
                if rest == "":
                    si += 1
                else:
                    node['text'] = rest; node['started'] = True
                    break
        return items, si

    def _pack_page(stream, si):
        """产出一页内容页的一栏或两栏，按『剩余正文量』分档，确保底栏尽量贴底 23.7：
        · 单栏(NCOLS==1)：整页就是一个通栏文本块，一次 _fill_col 填满到 CAP_CM+FILL_BONUS
          （剩余不足一页时 _fill_col 自然只装到 stream 末尾），始终返回 [col0]。
        · 双栏(NCOLS==2)：按剩余分三档确保底栏贴底且两栏对齐：
          - 满页（剩余 > 2×CAP）：两栏都填满到栏高 ⇒ 同时贴底且等高对齐。
          - 中长段（CAP < 剩余 ≤ 2×CAP）：两栏等高（target=剩余/2）收尾。
          - 短段（剩余 ≤ CAP）：左栏吃满全部内容（右栏空），显著缩小与 23.7 的空隙。
        返回统一为 (cols_list, si)，供 flush_body 按 NCOLS 长度消费。"""
        total_rem = 0.0
        for nd in stream[si:]:
            if nd['type'] == 'head':
                total_rem += heading_cost_cm(nd['block']['level'], nd['block']['text'])
            else:
                total_rem += para_cost_cm(nd['text'])
        if NCOLS == 1:
            target = CAP_CM + FILL_BONUS
            col0, si = _fill_col(stream, si, target)
            return [col0], si
        # 双栏三档
        if total_rem > 2.0 * CAP_CM + 1e-6:
            target = CAP_CM + FILL_BONUS          # 满页：两栏填满到栏高(含补偿) ⇒ 栏底逼近 23.27
        elif total_rem >= CAP_CM - 1e-6:
            target = total_rem / 2.0 + FILL_BONUS / 2.0   # 中长段：两栏等高收尾
        else:
            target = CAP_CM + FILL_BONUS          # 短段：左栏吃满全部内容（右栏空）
        col0, si = _fill_col(stream, si, target)
        col1, si = _fill_col(stream, si, target)
        return [col0, col1], si

    body = []
    def flush_body():
        nonlocal body
        if not body:
            return
        stream = []
        for b in body:
            if b['kind'] == 'heading':
                stream.append({'type': 'head', 'block': b})
            else:
                stream.append({'type': 'para', 'block': b, 'text': b['text'], 'started': False})
        si = 0
        while si < len(stream):
            cols, si = _pack_page(stream, si)
            if not any(cols):
                break
            pages.append({'type': 'content', 'chapter': cur_ch, 'no': cur_no,
                         'cols': cols, 'items': [it for c in cols for it in c]})
        body = []

    i, N = 0, len(blocks)
    while i < N:
        b = blocks[i]; k = b['kind']
        if k == 'chapter':
            flush_body()
            cur_ch = b['text']; cur_no = cn_to_int(b['text'].split('、')[0])
            pages.append({'type': 'chapter', 'no': cur_no, 'title': b['text']}); i += 1
        elif k in ('heading', 'para'):
            body.append(b); i += 1
        elif k == 'image':
            flush_body()
            # 收集连续 image 块，打包进同一 gallery 页（消除“一页一图”碎页）
            j = i; imgs = []
            while j < N and blocks[j]['kind'] == 'image':
                imgs.append(blocks[j]); j += 1
            n = len(imgs)
            if n == 1:
                asp = imgs[0].get('aspect') or 1.6
                band_h = min(CONTENT_W / asp, CONTENT_H * IMG_BAND_MAX)
                i = fill_below('image', imgs[0], band_h, j)
            else:
                band_h = CONTENT_H * 0.90
                i = fill_below('gallery', {'images': imgs}, band_h, j)
        elif k == 'table':
            H = table_full_h(b['rows'])
            band_h = min(H * 1.12 + 0.7, CONTENT_H * 0.80)
            lead = body[0] if len(body) == 1 and body[0]['kind'] == 'heading' else None
            lead_h = hybrid_lead_height({'lead_heading': lead['text'], 'lead_level': lead['level']}) if lead else 0
            can_mix = band_h + lead_h <= CONTENT_H - MIN_BELOW_H - BAND_GAP
            if can_mix and lead:
                body = []
            flush_body()
            if can_mix:
                i = fill_below('table', {'rows': b['rows'], 'continued': False}, band_h, i + 1, lead)
            else:
                rows = b['rows']; ncols = max((len(r) for r in rows), default=1); cw = CONTENT_W / ncols
                LIMIT = CONTENT_H * 0.95; chunks = []; cur_c = []; hh = 0.0
                for row in rows:
                    rh = max((cell_h(row[ci] if ci < len(row) else {'text': '', 'images': []}, cw)
                              for ci in range(ncols)), default=0.5)
                    if hh + rh > LIMIT and cur_c: chunks.append(cur_c); cur_c = []; hh = 0.0
                    cur_c.append(row); hh += rh
                if cur_c: chunks.append(cur_c)
                for j, c in enumerate(chunks):
                    pages.append({'type': 'table', 'rows': c, 'continued': j > 0,
                                  'chapter': cur_ch, 'no': cur_no})
                i += 1
        else:
            i += 1
    flush_body()
    return pages

def _imprint():
    """页脚 / 章节页落款：机构名 + 密级标注（密级为空时只显示机构名）。"""
    return "%s  |  %s" % (ORG_NAME, DOC_CLASS) if DOC_CLASS else ORG_NAME

def header(slide, title, no, folio=0, is_toc=False, eyebrow=None):
    simple_text(slide, ML, 0.7, CONTENT_W, 0.45, BOOK_TITLE, 8, False, GRAY, HEI)
    add_rect(slide, ML, 1.35, CONTENT_W, 0.012, None, RULE, Pt(0.75))
    tsz = fit_size(title, 13.8, 19)
    simple_text(slide, ML, 1.75, CONTENT_W, 0.95, title, tsz, True, BLUE, TITLE_F)
    if eyebrow is None:
        eyebrow = "CONTENTS" if is_toc else "CHAPTER %02d · %s" % (no, BOOK_TITLE)
    simple_text(slide, ML, 2.85, CONTENT_W, 0.4, eyebrow, 9, False, GRAY, HEI)
    add_rect(slide, ML, 3.5, 2.2, 0.04, BLUE)

def slide_number_field(slide, x, y, w, h, sz=7.5, color=GRAY, family=HEI, cached="1"):
    """右下角自动页码域：随幻灯片位置自动更新。"""
    tb = slide.shapes.add_textbox(Cm(x), Cm(y), Cm(w), Cm(h))
    tf = tb.text_frame; tf.word_wrap = False
    tf.margin_left = 0; tf.margin_right = 0; tf.margin_top = 0; tf.margin_bottom = 0
    p = tf.paragraphs[0]; p.alignment = PP_ALIGN.RIGHT
    fld = parse_xml(
        '<a:fld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" '
        'id="{4D3A1B31-561B-4472-90C7-029379FC5CB2}" type="slidenum">'
        '<a:rPr lang="zh-CN" smtClean="0"/><a:t>' + cached + '</a:t></a:fld>')
    rPr = fld.find(qn('a:rPr'))
    rPr.set('sz', str(int(sz * 100)))
    sf = rPr.makeelement(qn('a:solidFill'), {})
    clr = sf.makeelement(qn('a:srgbClr'), {'val': '%02X%02X%02X' % (color[0], color[1], color[2])})
    sf.append(clr); rPr.append(sf)
    rPr.append(rPr.makeelement(qn('a:latin'), {'typeface': family}))
    rPr.append(rPr.makeelement(qn('a:ea'), {'typeface': family}))
    p._p.append(fld)
    return tb

def footer(slide, folio):
    add_rect(slide, ML, FOOTER_LINE_Y, CONTENT_W, 0.012, None, RULE, Pt(0.75))
    simple_text(slide, ML, FOOTER_TEXT_Y, CONTENT_W*0.70, 0.5, _imprint(), 7.5, False, GRAY, HEI)
    slide_number_field(slide, ML+CONTENT_W*0.70, FOOTER_TEXT_Y, CONTENT_W*0.30, 0.5,
                       sz=7.5, color=GRAY, family=HEI, cached=str(folio))

# ============================================================
# 6. 页面渲染
# ============================================================
def render_cover(slide, cover_img, editors):
    if cover_img and os.path.exists(cover_img):
        slide.shapes.add_picture(cover_img, Cm(0), Cm(-PH*0.1212), Cm(PW*1.177), Cm(PH*1.2104))
    add_rect(slide, 0, PH*0.608, PW, PH*0.392, BLUE)
    simple_text(slide, ML, PH*0.630, CONTENT_W, 0.5, BRAND_LINE, 9, True, BLUE_LT, HEI)
    # 封面主标题：按版心宽自适应字号，避免长书名溢出（fit_size 钳制在 [14,29]）
    tsize = fit_size(BOOK_TITLE, 14.0, 29, 14)
    simple_text(slide, ML, PH*0.670, CONTENT_W, 1.5, BOOK_TITLE, tsize, True, WHITE, COVER_F)
    if BOOK_SUBTITLE:
        ssize = fit_size(BOOK_SUBTITLE, CONTENT_W, 18, 12)
        simple_text(slide, ML, PH*0.724, CONTENT_W, 0.8, BOOK_SUBTITLE, ssize, True, WHITE, COVER_F)
    if BOOK_TITLE_EN:
        esize = fit_size(BOOK_TITLE_EN, 14.0, 9.5, 7)
        simple_text(slide, ML, PH*0.758, CONTENT_W, 0.6, BOOK_TITLE_EN, esize, False, BLUE_LT, HEI)
    if BOOK_EDITION:
        simple_text(slide, ML, PH*0.790, CONTENT_W, 0.6, BOOK_EDITION, 12, False, BLUE_LT, HEI)
    add_rect(slide, ML, PH*0.824, 4.5, 0.04, WHITE)
    tf = new_tb(slide, ML, PH*0.834, CONTENT_W, 2.4)
    para(tf, True, [{'t': "主编单位：", 'sz': Pt(10.5), 'b': True, 'c': WHITE, 'f': HEI},
                    {'t': editors.get("主编单位", ""), 'sz': Pt(10.5), 'c': BLUE_LT, 'f': COVER_F}], ls=1.4)
    para(tf, False, [{'t': "参编单位：", 'sz': Pt(10), 'b': True, 'c': WHITE, 'f': HEI},
                     {'t': editors.get("参编单位", ""), 'sz': Pt(10), 'c': BLUE_LT, 'f': COVER_F}], ls=1.4)
    para(tf, False, [{'t': "发布日期：", 'sz': Pt(10.5), 'b': True, 'c': WHITE, 'f': HEI},
                     {'t': editors.get("发布日期", "2026年7月"), 'sz': Pt(10.5), 'c': BLUE_LT, 'f': COVER_F}], ls=1.4)
    simple_text(slide, ML, PH*0.952, CONTENT_W-2.8, 0.5, ORG_NAME, 9, False, BLUE_LT, HEI)
    if DOC_CLASS:
        add_rect(slide, PW-MR-2.2, PH*0.946, 2.2, 0.75, WHITE, rounded=True)
        simple_text(slide, PW-MR-2.2, PH*0.946, 2.2, 0.75, DOC_CLASS, 10.5, True, BLUE, HEI,
                    align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)

def render_toc(slide, toc_names, folios, folio=4):
    header(slide, "目录", 0, folio, is_toc=True)
    if NCOLS == 1:
        # 单栏版：整页一列（版心满宽），与正文栏式一致
        x = ML
        col_w = CONTENT_W
        y0 = 4.6
        y = y0
        for k, nm in enumerate(toc_names):
            idx = k + 1
            tf = new_tb(slide, x, y, col_w, 0.9, wrap=False)
            p = tf.paragraphs[0]
            a = p.add_run(); a.text = "%02d" % idx; a.font.size = Pt(12.5); a.font.bold = True; a.font.color.rgb = BLUE; set_cjk(a, HEI)
            t1 = p.add_run(); t1.text = "\t"; t1.font.size = Pt(12.5); set_cjk(t1, HEI)
            b = p.add_run(); b.text = nm; b.font.size = Pt(12.5); b.font.color.rgb = INK; set_cjk(b, COVER_F)
            t2 = p.add_run(); t2.text = "\t"; t2.font.size = Pt(12.5); set_cjk(t2, HEI)
            c = p.add_run(); c.text = "%02d" % folios.get(idx, 0); c.font.size = Pt(12.5); c.font.bold = True; c.font.color.rgb = BLUE; set_cjk(c, HEI)
            pPr = p._p.get_or_add_pPr()
            tabLst = pPr.makeelement(qn('a:tabLst'), {}); pPr.append(tabLst)
            for pos, algn in ((1.4, 'l'), (col_w, 'r')):
                tabLst.append(tabLst.makeelement(qn('a:tab'), {'pos': str(int(pos * 360000)), 'algn': algn}))
            y += 1.42
    else:
        # 双栏版：左右两列
        half = (len(toc_names) + 1) // 2
        col_gap = 0.5
        col_w = (14.2 - col_gap) / 2.0
        y0 = 4.6
        for ci in range(2):
            x = 1.8 + ci * (col_w + col_gap)
            items = toc_names[ci*half:(ci+1)*half] if ci == 0 else toc_names[half:]
            y = y0
            start = 1 if ci == 0 else half + 1
            for k, nm in enumerate(items):
                idx = start + k
                tf = new_tb(slide, x, y, col_w, 0.9, wrap=False)
                p = tf.paragraphs[0]
                a = p.add_run(); a.text = "%02d" % idx; a.font.size = Pt(12.5); a.font.bold = True; a.font.color.rgb = BLUE; set_cjk(a, HEI)
                t1 = p.add_run(); t1.text = "\t"; t1.font.size = Pt(12.5); set_cjk(t1, HEI)
                b = p.add_run(); b.text = nm; b.font.size = Pt(12.5); b.font.color.rgb = INK; set_cjk(b, COVER_F)
                t2 = p.add_run(); t2.text = "\t"; t2.font.size = Pt(12.5); set_cjk(t2, HEI)
                c = p.add_run(); c.text = "%02d" % folios.get(idx, 0); c.font.size = Pt(12.5); c.font.bold = True; c.font.color.rgb = BLUE; set_cjk(c, HEI)
                pPr = p._p.get_or_add_pPr()
                tabLst = pPr.makeelement(qn('a:tabLst'), {}); pPr.append(tabLst)
                for pos, algn in ((1.4, 'l'), (col_w, 'r')):
                    tabLst.append(tabLst.makeelement(qn('a:tab'), {'pos': str(int(pos * 360000)), 'algn': algn}))
                y += 1.42
    footer(slide, folio)

def render_chapter(slide, page, sections, ch_en):
    no = page['no']; title = page['title']
    add_rect(slide, 0, 0, PW, PH, BLUE)
    add_rect(slide, 0, 0, 0.6, PH, BLUE_D)
    wm = title.split('、', 1)[-1][:1]
    simple_text(slide, -1.0, 13.0, 14.0, 12.0, wm, 280, True, BLUE_D, TITLE_F)
    simple_text(slide, 2.4, 4.2, 9.0, 4.5, "%02d" % no, 140, True, WHITE, HEI)
    tf = new_tb(slide, 2.6, 11.8, 14.2, 2.2)
    ch_t = title.split('、', 1)[-1]
    para(tf, True, [{'t': ch_t, 'sz': Pt(fit_size(ch_t, 13.6, 48, min_pt=24)), 'b': True, 'c': WHITE, 'f': TITLE_F}])
    para(tf, False, [{'t': ch_en.get(no, "CHAPTER"), 'sz': Pt(14), 'c': BLUE_LT, 'f': HEI}], sb=Pt(8))
    add_rect(slide, 2.6, 15.6, 5.0, 0.04, WHITE)
    tf2 = new_tb(slide, 2.6, 16.3, 14.2, 6.0)
    para(tf2, True, [{'t': "本章内容", 'sz': Pt(10), 'b': True, 'c': BLUE_LT, 'f': HEI}], sa=Pt(6))
    for s in sections[:6]:
        para(tf2, False, [{'t': head_disp(s), 'sz': Pt(12), 'c': WHITE, 'f': SONG}], ls=1.5)
    simple_text(slide, 2.6, PH-1.5, PW-4.2, 0.6, _imprint(), 7.5, False, BLUE_LT, HEI)

def _render_table_rows(slide, rows, y_start, continued, media_dir=None):
    ncols = max((len(r) for r in rows), default=1); cw = CONTENT_W / ncols
    y = y_start
    for ri, row in enumerate(rows):
        h = max((cell_h(row[ci] if ci < len(row) else {'text': '', 'images': []}, cw)
                 for ci in range(ncols)), default=0.5)
        cx = ML
        for ci in range(ncols):
            cell = row[ci] if ci < len(row) else {'text': '', 'images': []}
            head = (ri == 0 and not continued)
            add_rect(slide, cx, y, cw, h, BLUE_LT if head else WHITE, RULE, Pt(0.5))
            if cell.get('text'):
                tf = new_tb(slide, cx + 0.1, y + 0.07, cw - 0.2, max(h - 0.14, 0.3))
                para(tf, True, [{'t': cell['text'], 'sz': Pt(8.5), 'b': head,
                                 'c': BLUE if head else INK, 'f': HEI if head else SONG}], ls=1.12)
            im_y = y + 0.07
            if cell.get('text'):
                im_y += max(1, math.ceil(text_units(cell['text']) / max(2, int(cw / 0.33)))) * 0.42 + 0.05
            for im in cell.get('images', []):
                p = im.get('path', '')
                if media_dir and not os.path.isabs(p): p = os.path.join(media_dir, os.path.basename(p))
                if p and os.path.exists(p):
                    asp = im.get('aspect') or 1.6
                    iw = cw - 0.2; ih = min(iw / asp, max(h - (im_y - y) - 0.08, 0))
                    if ih > 0.3:
                        slide.shapes.add_picture(p, Cm(cx + 0.1), Cm(im_y), Cm(iw), Cm(ih)); im_y += ih + 0.06
            cx += cw
        y += h
    return y

def _render_block(tf, b, first):
    """内联渲染单个块到给定 tf（content 双栏用）。
    对段落续段（continued=True）取消首行缩进与段后距，使被 balance 拆开的同一段
    在视觉上仍是连续的一个段落，而非多个独立段。"""
    if b['kind'] == 'heading':
        lvl = b['level']; t = cjk_join_punct(head_disp(b['text']))
        if lvl == 2:
            para(tf, first, [{'t': "▍ " + t, 'sz': Pt(12.5), 'b': True, 'c': BLUE, 'f': HEI}], ls=1.15, sb=Pt(0 if first else 10), sa=Pt(4))
        elif lvl == 3:
            para(tf, first, [{'t': t, 'sz': Pt(11.5), 'b': True, 'c': BLUE_D, 'f': HEI}], ls=1.15, sb=Pt(0 if first else 8), sa=Pt(3))
        else:
            para(tf, first, [{'t': t, 'sz': Pt(10.5), 'b': True, 'c': INK, 'f': HEI}], ls=1.15, sb=Pt(0 if first else 6), sa=Pt(2))
    else:
        m = NUM_LEAD.match(b['text']); FI = Pt(21) if b.get('indent', True) else Pt(0)
        SA = Pt(0) if b.get('continued', False) else Pt(4)
        if m:
            para(tf, first, [{'t': cjk_join_punct(m.group(1)), 'sz': Pt(10.5), 'b': True, 'c': INK, 'f': HEI},
                            {'t': cjk_join_punct(m.group(2)), 'sz': Pt(10.5), 'c': INK, 'f': BODY}],
                 align=PP_ALIGN.JUSTIFY, ls=1.45, sa=SA, fi=FI)
        else:
            para(tf, first, [{'t': cjk_join_punct(b['text']), 'sz': Pt(10.5), 'c': INK, 'f': BODY}],
                 align=PP_ALIGN.JUSTIFY, ls=1.55, sa=SA, fi=FI)

def render_content_v2(slide, page, folio, media_dir=None):
    """content 正确实现：每栏单文本框流式段落（单/双栏自适应）。"""
    header(slide, page['chapter'], page['no'], folio)
    if NCOLS == 2 and DIV_X is not None:
        add_rect(slide, DIV_X, TOP, 0.01, CONTENT_BOTTOM - TOP, None, RULE, Pt(0.5))
    for c in range(NCOLS):
        x = COL1_X if (NCOLS == 1 or c == 0) else COL2_X
        items = page['cols'][c]
        if not items: continue
        y = TOP; tf = None; first = True
        for it in items:
            b = it['block']
            if b['kind'] == 'image':
                tf = None
                asp = b.get('aspect') or 1.6
                ih = min(COL_W / asp, CONTENT_H * 0.9)
                p = b.get('path', '')
                if media_dir and not os.path.isabs(p): p = os.path.join(media_dir, os.path.basename(p))
                if p and os.path.exists(p):
                    slide.shapes.add_picture(p, Cm(x), Cm(y), Cm(COL_W), Cm(ih))
                if b.get('caption'):
                    simple_text(slide, x, y + ih + 0.08, COL_W, 0.7, b['caption'], 8, False, GRAY, SONG)
                y += it['cost']; continue
            if tf is None:
                tf = new_tb(slide, x, y, COL_W, max(CONTENT_BOTTOM - y, 0.6)); first = True
            _render_block(tf, b, first)
            first = False; y += it['cost']
    footer(slide, folio)

def render_table(slide, page, folio, media_dir=None):
    header(slide, page['chapter'], page['no'], folio)
    if page.get('continued'):
        simple_text(slide, 1.8, 3.85, 14.2, 0.4, "（续表）", 8.5, False, GRAY, HEI)
    y_start = TOP + (0.5 if page.get('continued') else 0)
    # 非续表的小表在页内垂直居中，避免页面上方挤满、下方大留白的不平衡感
    if not page.get('continued'):
        th = table_rows_height(page['rows'])
        y_start = TOP + (CONTENT_H - th) / 2
        y_start = max(TOP, y_start)
    _render_table_rows(slide, page['rows'], y_start, page.get('continued'), media_dir)
    footer(slide, folio)

def _render_below(slide, page, below_top, below_h, media_dir):
    """混排页/图库页通栏带下方的栏区（单/双栏自适应）渲染。"""
    if NCOLS == 2 and DIV_X is not None:
        add_rect(slide, DIV_X, below_top, 0.01, below_h, None, RULE, Pt(0.5))
    for c in range(NCOLS):
        x = COL1_X if (NCOLS == 1 or c == 0) else COL2_X
        items = page['cols'][c]
        if not items: continue
        tf = new_tb(slide, x, below_top, COL_W, max(below_h, 0.6)); first = True
        for it in items:
            _render_block(tf, it['block'], first); first = False

def render_hybrid(slide, page, folio, media_dir=None):
    header(slide, page['chapter'], page['no'], folio)
    band_h = page['band_h']
    lead_h = hybrid_lead_height(page)
    band_top = TOP + lead_h
    if lead_h:
        fs = {2: 12.5, 3: 11.5, 4: 10.5}[page.get('lead_level', 2)]
        simple_text(slide, ML, TOP, CONTENT_W, lead_h - 0.22,
                    head_disp(page['lead_heading']), fs, True, BLUE, HEI)
    if page['band_kind'] == 'image':
        b = page['band']; p = b.get('path', '')
        if media_dir and not os.path.isabs(p): p = os.path.join(media_dir, os.path.basename(p))
        if p and os.path.exists(p):
            slide.shapes.add_picture(p, Cm(ML), Cm(band_top), Cm(CONTENT_W), Cm(band_h))
        if b.get('caption'):
            simple_text(slide, ML, band_top + band_h + 0.06, CONTENT_W, 0.7, b['caption'], 8, False, GRAY, SONG)
    else:
        _render_table_rows(slide, page['band']['rows'], band_top, False, media_dir)
    below_top = band_top + band_h + BAND_GAP
    below_h = CONTENT_BOTTOM - below_top
    _render_below(slide, page, below_top, below_h, media_dir)
    footer(slide, folio)

def render_gallery(slide, page, folio, media_dir=None):
    """图库页：连续多张图打包进同一页（消除“一页一图”碎页）。上区为图网格（竖向堆叠、满版心宽），
    下区为可选栏内文字（与混排页相同的 below 机制）。"""
    header(slide, page['chapter'], page['no'], folio)
    band_h = page['band_h']
    imgs = page['band']['images']
    n = len(imgs)
    slot = (band_h - 0.15 * (n - 1)) / n      # 每张图占一个竖向槽位
    y = TOP
    for im in imgs:
        p = im.get('path', '')
        if media_dir and not os.path.isabs(p): p = os.path.join(media_dir, os.path.basename(p))
        asp = im.get('aspect') or 1.6
        ih = min(CONTENT_W / asp, slot - 0.45)   # 留 0.45cm 给图注
        ih = max(ih, 0.4)
        iw = ih * asp
        x = ML + (CONTENT_W - iw) / 2.0
        if p and os.path.exists(p):
            slide.shapes.add_picture(p, Cm(x), Cm(y), Cm(iw), Cm(ih))
        if im.get('caption'):
            simple_text(slide, ML, y + ih + 0.05, CONTENT_W, 0.45, im['caption'], 8, False, GRAY, SONG)
        y += slot
    below_top = TOP + band_h + BAND_GAP
    below_h = CONTENT_BOTTOM - below_top
    _render_below(slide, page, below_top, below_h, media_dir)
    footer(slide, folio)

def render_colophon(slide, frontmatter, folio):
    header(slide, "编委会", 0, folio, eyebrow="EDITORIAL BOARD · %s" % BOOK_TITLE)
    body = [t for t in frontmatter if t.strip()]
    tf = new_tb(slide, ML, TOP, CONTENT_W, CONTENT_BOTTOM-TOP)
    first = True
    for t in body:
        is_head = t in (BOOK_TITLE, "编委会") or t.startswith(("主编单位", "参编单位"))
        para(tf, first, [{'t': t, 'sz': Pt(11 if is_head else 10.5), 'b': is_head,
                         'c': BLUE if is_head else INK, 'f': HEI if is_head else COVER_F}], ls=1.5, sa=Pt(3))
        first = False
    footer(slide, folio)

def render_back_cover(slide):
    """封尾页：蓝底 + 主标语（options['slogan']，缺省回退书名）+ 机构落款。"""
    add_rect(slide, 0, 0, PW, PH, BLUE)
    add_rect(slide, 0, 0, 0.6, PH, BLUE_D)
    lines = [s for s in (SLOGAN or ()) if s] or [BOOK_TITLE]
    y0 = 0.40 if len(lines) > 1 else 0.46
    for k, ln in enumerate(lines[:2]):
        simple_text(slide, 0, PH*(y0 + 0.18*k), PW, 1.4, ln, 60, True, WHITE, COVER_F,
                    align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    simple_text(slide, 0, PH*0.80, PW, 0.6, ORG_NAME, 11, False, BLUE_LT, HEI,
                align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)

# ============================================================
# 7. 组装
# ============================================================
def build_album(blocks, frontmatter, toc_names, ch_en, out,
                media_dir=None, cover_img=None, editors=None,
                write_page_types=True, page_types_path=None, options=None):
    """options（dict）可覆盖项：
       - body_font / title_font / cover_font / sans_font / song_font：字体栈
         （cover_font 专管封面主标/目录章名/编委会/封尾标语，默认思源宋体）
       - columns：1=单栏（默认）, 2=双栏
       - page_size：'B5'(默认) / 'A4' / (宽,高)cm 元组
       - add_blank：是否插入空白扉页（默认 False，避免低密度前导页）
       - book_title / book_subtitle / book_title_en / edition：画册名、中英文副标、版次标注
       - brand_line / org_name / doc_class：封面品牌行、机构落款、密级标注
       - slogan：封尾标语二元组，如 ("匠心筑造", "品质焕新")；缺省回退书名
      注意：文案类选项若不传，将沿用引擎默认（集团通用口径），**不会**残留上一次
      调用的值（`_options_scope` 在 build 结束后还原全部全局令牌）。
    """
    options = options or {}
    editors = editors or {}
    add_blank = options.get('add_blank', False)
    with _options_scope(options):
        return _build_album(blocks, frontmatter, toc_names, ch_en, out,
                            media_dir, cover_img, editors, write_page_types, page_types_path,
                            add_blank=add_blank)


def _build_album(blocks, frontmatter, toc_names, ch_en, out,
                 media_dir=None, cover_img=None, editors=None,
                 write_page_types=True, page_types_path=None, add_blank=False):
    editors = editors or {}
    sections = {}; cur = 0
    for b in blocks:
        if b['kind'] == 'chapter':
            cur = cn_to_int(b['text'].split('、')[0]); sections[cur] = []
        elif b['kind'] == 'heading' and b['level'] == 2 and cur:
            sections[cur].append(b['text'])
    pages = paginate(blocks)
    # 前页序列：封面(1) + [空白页] + 编委会 + 目录；fp = 目录页 folio
    fp = 1 + (1 if add_blank else 0) + 1 + 1
    # 章节页 folio 按其在 pages 中的下标推导（前页 fp 张，正文从 fp+1 起）
    folios = {}
    for j, pg in enumerate(pages):
        if pg['type'] == 'chapter': folios[pg['no']] = fp + 1 + j
    prs = Presentation(); prs.slide_width = Cm(PW); prs.slide_height = Cm(PH)
    # 公开示例只使用中性署名，不继承模板作者。
    prs.core_properties.author = ORG_NAME
    prs.core_properties.last_modified_by = ''
    prs.core_properties.comments = ''
    prs.core_properties.title = BOOK_TITLE
    BLANK = prs.slide_layouts[6]
    render_cover(prs.slides.add_slide(BLANK), cover_img, editors)
    if add_blank:
        prs.slides.add_slide(BLANK)
    render_colophon(prs.slides.add_slide(BLANK), frontmatter, fp - 1)
    render_toc(prs.slides.add_slide(BLANK), toc_names, folios, folio=fp)
    for j, pg in enumerate(pages):
        s = prs.slides.add_slide(BLANK)
        folio = fp + 1 + j
        if pg['type'] == 'chapter': render_chapter(s, pg, sections.get(pg['no'], []), ch_en)
        elif pg['type'] == 'content': render_content_v2(s, pg, folio, media_dir)
        elif pg['type'] == 'hybrid': render_hybrid(s, pg, folio, media_dir)
        elif pg['type'] == 'gallery': render_gallery(s, pg, folio, media_dir)
        else: render_table(s, pg, folio, media_dir)
    render_back_cover(prs.slides.add_slide(BLANK))
    prs.save(out)
    # 页面类型清单（供 audit_overflow 按页型施检）
    if write_page_types:
        ptypes = [{'t': 'cover'}, {'t': 'colophon'}, {'t': 'toc'}]
        if add_blank:
            ptypes.insert(1, {'t': 'blank'})
        for pg in pages:
            if pg['type'] in ('hybrid', 'gallery'):
                bt = round(TOP + hybrid_lead_height(pg) + pg['band_h'] + BAND_GAP, 3)
                ptypes.append({'t': pg['type'], 'below_top': bt})
            else:
                ptypes.append({'t': pg['type']})
        ptypes.append({'t': 'back_cover'})
        pp = page_types_path or os.path.join(os.path.dirname(os.path.abspath(out)), 'page_types.json')
        json.dump(ptypes, open(pp, 'w'), ensure_ascii=False, indent=1)
    return out

if __name__ == "__main__":
    # 最小自测：生成封面/编委会/目录/章节/单栏正文/混排各一页，验证引擎可运行。
    # 示例文案均为无实际业务含义的占位内容。
    demo_blocks = [
        {"kind": "chapter", "text": "一、总则"},
        {"kind": "heading", "level": 2, "text": "1.1 适用范围"},
        {"kind": "para", "text": "本示例用于验证排版引擎的单栏流式排版、两端对齐与标点不悬挂等基础能力。" * 3},
        {"kind": "heading", "level": 3, "text": "1.1.1 基本原则"},
        {"kind": "para", "text": "示例段落，用于填充版心并检验行高模型与栏底贴齐效果，不含任何业务信息。" * 3},
        {"kind": "image", "path": "", "aspect": 1.6, "caption": "图1 示例插图"},
        {"kind": "para", "text": "混排页示例段落，用于验证通栏图带置顶后下方栏区高度自动缩短的逻辑。" * 3},
    ]
    demo_toc = ["一、总则", "二、示例章节", "三、示例章节"]
    demo_ch_en = {1: "GENERAL PRINCIPLES"}
    demo_editors = {"主编单位": "示例主编单位",
                    "参编单位": "示例参编单位",
                    "发布日期": "2026年1月"}
    build_album(demo_blocks, ["示例画册", "编委会", "主编单位：", "参编单位：", "发布日期："],
                demo_toc, demo_ch_en, "/tmp/standard_blue_album_demo.pptx",
                media_dir=None, cover_img=None, editors=demo_editors,
                options={"book_title": "示例画册",
                         "book_title_en": "SAMPLE ALBUM",
                         "slogan": ("匠心筑造", "品质焕新")})
    print("DEMO SAVED: /tmp/standard_blue_album_demo.pptx")
