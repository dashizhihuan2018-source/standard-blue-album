# -*- coding: utf-8 -*-
"""生成双栏满版小样，用于校核栏底贴齐、左右栏平衡与溢出情况。

输出：samples/sample_3pages.pptx + samples/page_types.json

说明：本文件中的全部文字均为**无业务含义的占位示例**，仅用于产生足量正文以
填满版心。复用本引擎时请替换为自己的 blocks 数据源。
"""
import os
import sys
import json

_HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, os.path.join(os.path.dirname(_HERE), "references"))

from b5_album_engine import build_album  # noqa: E402

# ---- 占位语料：组合生成长度不一的段落，模拟真实书籍的段落长度分布 ----------
_CLAUSES = [
    "本段为排版校核用占位文本，用于产生稳定的字数以填满版心",
    "文字内容不具备任何业务含义，仅供检验行高模型与栏底贴齐效果",
    "两端对齐与标点不悬挂是本引擎的两项基础约束，需在渲染后逐行核验",
    "段落首行缩进两字符，段后距为零，依靠缩进区分段落边界",
    "行粒度贪婪填栏算法会在栏容量耗尽时把剩余文字续排至下一栏",
    "标题块不可跨栏拆分，若当前栏放不下则整体下移至下一栏起始位置",
    "成本模型按厘米计量，累计段后距后再与栏容量比较，避免纯行数估算的系统性低估",
    "渲染完成后应转换为 PDF 并逐行断言 bbox，确认无越栏与压底情况",
    "图表带区置顶时下方栏区高度自动缩短，成本模型需同步按缩短后的容量拆分段落",
    "版本迭代应递增版本号并另存副本，便于回溯比对与问题定位",
]


def _para(seed, n):
    """按 seed 取 n 个子句拼成一段占位正文。"""
    parts = [_CLAUSES[(seed + k) % len(_CLAUSES)] for k in range(n)]
    return "。".join(parts) + "。"


def _build_blocks():
    blocks = [{"kind": "chapter", "text": "一、总则"}]
    h2_titles = [
        "1.1 编制说明", "1.2 适用范围", "1.3 基本约定", "1.4 版式结构",
        "1.5 分栏规则", "1.6 图表处理", "1.7 校验流程", "1.8 交付要求",
    ]
    for i, t in enumerate(h2_titles, start=1):
        blocks.append({"kind": "heading", "level": 2, "text": t})
        blocks.append({"kind": "para", "text": _para(i, 4)})
        blocks.append({"kind": "para", "text": _para(i + 3, 5)})
        if i % 2 == 0:
            blocks.append({"kind": "heading", "level": 3,
                           "text": "1.%d.1 补充说明" % i})
            blocks.append({"kind": "para", "text": _para(i + 6, 3)})
            blocks.append({"kind": "para", "text": _para(i + 1, 6)})
    return blocks


FRONTMATTER = ["示例画册", "编委会", "主编单位：", "参编单位：", "发布日期："]
EDITORS = {"主编单位": "示例主编单位",
           "参编单位": "示例参编单位",
           "发布日期": "2026年1月"}
TOC = ["一、总则", "二、示例章节", "三、示例章节", "四、示例章节"]
CH_EN = {1: "GENERAL PRINCIPLES"}

OPTIONS = {
    "book_title": "示例画册",
    "book_title_en": "SAMPLE ALBUM",
    "slogan": ("匠心筑造", "品质焕新"),
}

if __name__ == "__main__":
    out = os.path.join(_HERE, "sample_3pages.pptx")
    # 使用独立页型文件，避免覆盖版式参考图样例的 page_types.json
    pt_path = os.path.join(_HERE, "sample_3pages.page_types.json")
    build_album(_build_blocks(), FRONTMATTER, TOC, CH_EN, out,
                media_dir=None, cover_img=None, editors=EDITORS,
                write_page_types=True, page_types_path=pt_path,
                options=OPTIONS)
    pt = json.load(open(pt_path))
    print("OUT:", out, "pages:", len(pt))
    for i, p in enumerate(pt):
        print(" ", i, p)
