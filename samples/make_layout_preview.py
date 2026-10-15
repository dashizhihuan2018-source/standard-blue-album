# -*- coding: utf-8 -*-
"""
生成「版式参考图」样例：以《示例工程技术指南》单栏 B5 排版为示例，
覆盖 cover / colophon / toc / chapter / content / table / hybrid / gallery / back_cover
全部九种版式。所有正文均为通用占位文案，不含真实项目编号、地址、人名或内部指标。

输出：
  samples/layout_preview.pptx   样例画册
  samples/page_types.json       页型清单（供审计与版式图映射）
  samples/media/placeholder*.png  图块占位示意图（不涉任何真实图纸）
"""
import os, sys
from PIL import Image, ImageDraw, ImageFont

HERE = os.path.dirname(os.path.abspath(__file__))
REF = os.path.join(os.path.dirname(HERE), "references")
sys.path.insert(0, REF)
import b5_album_engine as E   # noqa

MEDIA = os.path.join(HERE, "media")
os.makedirs(MEDIA, exist_ok=True)


def make_placeholder(name, label_cn, label_en, seed=0):
    """生成一张浅蓝灰占位示意图，仅含示意文字，不涉及任何真实图纸。"""
    W, H = 1200, 750
    img = Image.new("RGB", (W, H), (233, 239, 248))  # #E9EFF8 浅蓝
    d = ImageDraw.Draw(img)
    # 外框
    d.rectangle([20, 20, W - 20, H - 20], outline=(197, 212, 227), width=4)  # #C9D4E3
    # 示意网格（仅装饰）
    for x in range(120, W, 120):
        d.line([(x, 40), (x, H - 40)], fill=(214, 224, 236), width=1)
    for y in range(120, H, 120):
        d.line([(40, y), (W - 40, y)], fill=(214, 224, 236), width=1)
    # 中央示意块
    d.rectangle([W * 0.30, H * 0.32, W * 0.70, H * 0.68], outline=(0, 61, 165), width=6)
    try:
        family = os.environ.get('BLUE_ALBUM_FONT_PATH', 'SourceHanSerifCN-Regular.otf')
        f1 = ImageFont.truetype(family, 54)
        f2 = ImageFont.truetype(family, 30)
    except Exception:
        f1 = f2 = ImageFont.load_default()
    d.text((W / 2, H * 0.40), label_cn, font=f1, fill=(0, 42, 115), anchor="mm")
    d.text((W / 2, H * 0.58), label_en, font=f2, fill=(110, 110, 110), anchor="mm")
    out = os.path.join(MEDIA, name)
    img.save(out)
    return out


# 1) 占位示意图（不涉真实图纸）
make_placeholder("placeholder1.png", "平面示意", "SCHEMATIC PLAN")
make_placeholder("placeholder2.png", "节点示意", "DETAIL")


# 2) 通用占位正文（原拆原建领域，无真实项目 specifics）
def P(t): return {"kind": "para", "text": t}
def H2(t): return {"kind": "heading", "level": 2, "text": t}
def H3(t): return {"kind": "heading", "level": 3, "text": t}


def build_big_table():
    """生成一张较大的对照表，使其高度超过『通栏带 + 下方栏』阈值，从而成为独立的表格页（table 版式）。"""
    hdr = ["序号", "技术环节", "控制要点", "验收标准"]
    rows = [[{"text": hdr[c], "images": []} for c in range(4)]]
    items = [
        ("现状调查", "建筑尺寸与结构形式核查", "资料完整"),
        ("安全鉴定", "结构安全等级评定", "出具报告"),
        ("地质勘察", "地基承载力与沉降分析", "满足设计"),
        ("方案编制", "原规模原高度原外观", "通过评审"),
        ("专家论证", "风貌与结构双重把关", "意见闭环"),
        ("基础施工", "扩大截面或注浆补强", "控制沉降差"),
        ("上部结构", "钢筋网砂浆面层加固", "先鉴定后设计"),
        ("楼梯置换", "现浇整体置换", "保证疏散宽度"),
        ("外围护", "节能与防水同步提升", "气密水密合格"),
        ("设备更新", "管线综合与管井优化", "使用便捷"),
        ("消防系统", "疏散与报警全覆盖", "验收合格"),
        ("抗震构造", "关键节点可靠连接", "满足现行标准"),
        ("材料选用", "耐久可维护产品", "留存质保文件"),
        ("过程影像", "关键工序留痕", "归档可查"),
        ("绿色施工", "降尘降噪与资源循环", "文明施工"),
        ("质量检验", "分部分项抽检", "合格率达标"),
        ("安全监测", "施工期变形观测", "预警可控"),
        ("竣工验收", "六阶段资料汇总", "交付使用"),
        ("后期运维", "使用说明书与保修", "责任明确"),
        ("档案管理", "全过程文件归集", "长期保存"),
        ("风貌保护", "原规模原高度原外观", "专家认可"),
        ("节能环保", "围护与设备协同", "能耗达标"),
        ("智能建造", "BIM 全过程应用", "提质增效"),
        ("风险管控", "危大工程专项方案", "零事故"),
        ("造价控制", "限额设计与概预算", "投资可控"),
        ("工期管理", "关键线路动态优化", "按期竣工"),
        ("公众沟通", "公示与意见征询", "达成共识"),
        ("智慧运维", "监测数据平台化", "长效运行"),
    ]
    for idx, (a, b, c) in enumerate(items, 1):
        rows.append([{"text": str(idx), "images": []},
                     {"text": a, "images": []},
                     {"text": b, "images": []},
                     {"text": c, "images": []}])
    return rows

blocks = [
    {"kind": "chapter", "text": "一、总则"},
    H2("1.1 适用范围"),
    P("原拆原建是指在原址按原有建筑规模、高度、外观及使用功能进行拆除并重建的建设方式，"
      "适用于历史文化街区保护更新、老旧小区有机更新以及功能严重滞后的既有建筑改造等场景。"),
    P("在历史文化街区中，原拆原建应在不改变街区整体风貌与街巷尺度的前提下，改善建筑内部空间、"
      "设施条件与使用环境；老旧小区更新应充分尊重居民意愿，完善公共服务配套与无障碍设施；"
      "功能滞后建筑改造应结合城市规划要求，提升土地集约利用水平与安全韧性。"),
    P("适用范围不包括文物建筑、历史建筑以及法律法规规定不得拆除的建筑。涉及风貌保护、消防安全、"
      "结构安全、公共利益的特殊情形，应依法履行专家论证、审批公示与社会稳定风险评估等程序，"
      "确保更新工作合法合规、稳妥可控。"),
    P("各地可结合实际情况制定实施细则，明确原拆原建项目的准入条件、审批流程、技术标准、"
      "监管要求和退出机制，形成市、区、街道三级协同的管理机制；鼓励引入全过程工程咨询，"
      "统筹设计、施工、监理与运维各方责任。"),
    H3("1.1.1 基本原则"),
    P("开展原拆原建工作应当坚持尊重原有风貌、保障结构安全、提升使用功能、控制建设成本的基本原则，"
      "在保留场所记忆的同时改善人居环境。"),
    P("风貌协调原则要求新建建筑在体量、高度、色彩、材质、檐口线脚等方面与周边环境相协调；"
      "安全优先原则要求以现行国家标准为依据，开展全过程安全鉴定与监测；"
      "经济合理原则要求统筹一次投入与长期运维成本，避免过度建设。"),
    H2("1.2 工作程序"),
    P("原拆原建项目一般包括现状调查、安全鉴定、方案编制、专家论证、施工实施与竣工验收六个阶段，"
      "各阶段应形成完整的过程文件并归档留存。"),
    P("现状调查阶段应开展建筑测绘、结构检测、设备评估、使用情况调查和居民意愿征集，"
      "形成翔实的现状资料库；安全鉴定阶段应由具有相应资质的机构出具鉴定报告，"
      "明确结构安全等级、抗震构造措施和加固改造建议。"),
    P("方案编制阶段应坚持原规模、原高度、原外观的总体要求，结合结构鉴定结论和使用功能需求，"
      "形成建筑、结构、设备、景观等多专业协同的设计文件；专家论证阶段应组织建筑、结构、"
      "历史风貌、消防等领域专家进行技术审查，论证意见应逐条闭环。"),
    P("施工实施阶段应编制专项施工组织设计和安全监测方案，加强危大工程管控、文明施工与绿色建造；"
      "竣工验收阶段应组织建设、设计、施工、监理和相关主管部门进行联合验收，"
      "确保工程质量、风貌效果和使用功能全面达标。"),
    {"kind": "image", "path": "placeholder2.png", "aspect": 1.7, "caption": "图1 原建筑立面与结构体系示意"},
    P("原建筑立面与结构体系示意用于直观表达既有建筑的层数、开间、屋面形式、主要结构构件及损伤位置，"
      "为后续安全鉴定、结构复核与风貌评估提供可视化依据；图示内容应避免涉及真实门牌号、产权人姓名等敏感信息。"),

    {"kind": "chapter", "text": "二、技术标准"},
    H2("2.1 结构安全"),
    P("结构安全应重点核查原结构体系、材料实际性能与抗震能力，必要时采取整体加固或关键构件置换措施，"
      "确保重建后建筑满足现行设计标准要求。"),
    P("重建材料应符合现行标准，优先选用耐久性与可维护性良好的产品，并留存完整的质量证明文件。"),
    {"kind": "table", "rows": build_big_table()},
    H2("2.2 节点构造"),
    P("节点构造应保证新旧构件可靠连接，优先采用成熟、可检的构造做法，并留存关键工序的过程影像资料。"),
    {"kind": "image", "path": "placeholder1.png", "aspect": 1.6, "caption": "图1 典型平面布局示意"},
    {"kind": "image", "path": "placeholder2.png", "aspect": 1.6, "caption": "图2 结构加固节点示意"},
    {"kind": "image", "path": "placeholder1.png", "aspect": 1.6, "caption": "图3 外立面效果示意"},
]

frontmatter = ["示例工程技术指南", "编委会", "主编单位：", "参编单位：", "发布日期："]
toc_names = ["一、总则", "二、技术标准"]
ch_en = {1: "GENERAL PRINCIPLES", 2: "TECHNICAL STANDARDS"}
editors = {"主编单位": "示例主编单位", "参编单位": "示例参编单位", "发布日期": "2026年7月"}
options = {
    "book_title": "示例工程技术指南",
    "book_title_en": "STANDARDIZED GUIDELINE FOR DEMOLITION-REBUILD",
    "slogan": ("匠心筑造", "品质焕新"),
}

out = os.path.join(HERE, "layout_preview.pptx")
E.build_album(blocks, frontmatter, toc_names, ch_en, out,
              media_dir=MEDIA, cover_img=None, editors=editors,
              options=options, write_page_types=True,
              page_types_path=os.path.join(HERE, "page_types.json"))
print("LAYOUT PREVIEW SAVED:", out)
