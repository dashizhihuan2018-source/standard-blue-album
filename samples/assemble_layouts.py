# -*- coding: utf-8 -*-
"""
从 layout_preview.pptx 渲染的逐页 PNG 中，按 page_types.json 抽取每种版式首张，
源文档应只使用虚构的中性示例，汇总为 samples/layouts/ 下「每版式一张图」。

用法：
  pdftoppm -png -r 150 layout_preview.pdf /tmp/lp_pages/page
  python3 assemble_layouts.py
"""
import os, json, sys
from PIL import Image

HERE = os.path.dirname(os.path.abspath(__file__))
RAW = "/tmp/lp_pages"                  # pdftoppm 输出目录 (page-01.png ...)
PTYPES = os.path.join(HERE, "page_types.json")
OUT = os.path.join(HERE, "layouts")
os.makedirs(OUT, exist_ok=True)

def main():
    pt = json.load(open(PTYPES, encoding="utf-8"))
    assert len(pt) == len([f for f in os.listdir(RAW) if f.endswith(".png")]), \
        "page_types 条目数(%d) 与渲染页数不符" % len(pt)
    seen = set()
    order = ["cover", "colophon", "toc", "chapter", "content",
             "hybrid", "table", "gallery", "back_cover"]
    chosen = {}  # type -> png path
    for i, p in enumerate(pt):
        t = p["t"]
        if t in seen:
            continue
        seen.add(t)
        chosen[t] = os.path.join(RAW, "page-%02d.png" % (i + 1))

    # 按既定顺序输出，未出现的版式跳过
    saved = []
    for t in order:
        src = chosen.get(t)
        if not src or not os.path.exists(src):
            print("  跳过缺失版式:", t)
            continue
        im = Image.open(src).convert("RGB")
        # 源文档只含中性示例，无需依赖图上遮挡隐藏信息。
        dst = os.path.join(OUT, "%s.png" % t)
        im.save(dst, "PNG")
        saved.append((t, os.path.getsize(dst)))
        print("  ✓ %-12s -> layouts/%s.png (%d KB)" % (t, t, os.path.getsize(dst) // 1024))
    print("共生成 %d 张版式参考图" % len(saved))


if __name__ == "__main__":
    main()
