# 标准蓝画册 · Standard Blue Album

> **© 2026 Nankong 版权所有。仅供非商业技术交流；未经书面许可，不得将 Nankong 未另行授权的原创内容用于商业用途。** 第三方内容权利归原权利人所有。详见 [版权与使用条件](COPYRIGHT.md)。

用于 B5/A4 工程技术画册、研究报告与专业手册的排版 Skill 和原生可编辑 PPTX 生成工具。支持单栏/双栏、图表混排、自动页码、成品双页镜像与 PDF 排版检查。

以清晰的阅读秩序、克制的视觉语言和原生可编辑的内容，呈现专业研究与技术成果。

项目持续更新。这是独立的通用排版工具，不代表任何企业、机构或官方品牌规范。研究报告默认采用对外分享稿，保留技术分析与正式来源，移除内部编制说明、批注和备注。发布目的、隐私保护及第三方权利说明见 [使用与权利声明](DISCLAIMER.md)。

## 安装

```bash
git clone https://github.com/dashizhihuan2018-source/standard-blue-album.git \
  ~/.codex/skills/standard-blue-album
python -m pip install -r ~/.codex/skills/standard-blue-album/requirements.txt
```

其他支持文件夹型 Skill 的工具可将本目录放入相应的 skills 目录。使用规范见 [SKILL.md](SKILL.md)。生成与检查需要 Python 3.9+；导出 PDF 还需 LibreOffice 和 Poppler。字体须自行合规安装，仓库不附带字体文件。

## 功能与验证

- 主色 `#003DA5`，B5/A4、单栏/双栏；所有品牌文案可配置。
- 文本、表格、图形保持可编辑；图片使用通用示意素材。
- 标题、表格和下方正文共享分页预算，表前小标题从版心顶起排，不侵入页眉。
- PDF 溢出检查与页眉净空检查分别执行，不将“没有越界”等同于“没有相交”。

```bash
python -m unittest discover -s tests -v
python samples/make_layout_preview.py
soffice --headless --convert-to pdf samples/layout_preview.pptx --outdir build
python references/audit_overflow_single.py build/layout_preview.pdf samples/page_types.json
python references/audit_header_clearance.py build/layout_preview.pdf samples/page_types.json
```

基本调用：

```python
from references.b5_album_engine import build_album

build_album(
    blocks=[
        {"kind": "chapter", "text": "一、研究概述"},
        {"kind": "heading", "level": 2, "text": "1.1 研究范围"},
        {"kind": "para", "text": "此处填写经过核实的正式研究内容。"},
    ],
    frontmatter=["示例研究报告"],
    toc_names=["研究概述"],
    ch_en={1: "OVERVIEW"},
    out="report.pptx",
    options={
        "book_title": "示例研究报告",
        "brand_line": "STANDARD BLUE ALBUM",
        "org_name": "示例机构",
        "doc_class": "",
        "columns": 1,
    },
)
```

## 样张

样张为重新生成的虚构示例，不含真实项目、客户资料或个人信息。无个人 Git 历史、内部路径、旧文档署名或隐藏备注。

| 封面 | 正文 | 混排 |
|---|---|---|
| ![封面](samples/layouts/cover.png) | ![正文](samples/layouts/content.png) | ![混排](samples/layouts/hybrid.png) |

## 许可与第三方资源

代码尚未指定开源许可证，仓库公开不等同于授予代码再分发许可。第三方依赖遵循各自许可证。公开样张使用思源宋体生成，字体本身不随仓库分发；使用其他字体、商标或实际项目内容前须分别核对授权。

## 持续更新

排版引擎、样张与检查工具持续维护，以当前默认分支的文档和测试为准。反馈版式问题时，请提交由虚构内容生成的最小示例，勿上传客户原稿、个人信息或内部资料。
