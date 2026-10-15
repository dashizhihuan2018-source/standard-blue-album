from __future__ import annotations

import tempfile
import unittest
import zipfile
from pathlib import Path

from lxml import etree


P = "http://schemas.openxmlformats.org/presentationml/2006/main"
A = "http://schemas.openxmlformats.org/drawingml/2006/main"
NS = {"p": P, "a": A}


def shape_xml(shape_id: int, name: str, value: str, x: int, y: int, cx: int, cy: int, align: str) -> str:
    return f"""
    <p:sp xmlns:p="{P}" xmlns:a="{A}">
      <p:nvSpPr><p:cNvPr id="{shape_id}" name="{name}"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>
      <p:spPr><a:xfrm><a:off x="{x}" y="{y}"/><a:ext cx="{cx}" cy="{cy}"/></a:xfrm></p:spPr>
      <p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:pPr algn="{align}"/><a:r><a:t>{value}</a:t></a:r></a:p></p:txBody>
    </p:sp>"""


def make_fixture(path: Path) -> None:
    presentation = f"""<p:presentation xmlns:p="{P}"><p:sldSz cx="6335713" cy="8999538"/></p:presentation>"""
    master = f"""
    <p:sldMaster xmlns:p="{P}" xmlns:a="{A}"><p:cSld><p:spTree>
      <p:nvGrpSpPr><p:cNvPr id="1" name="Group"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/>
      <p:sp><p:nvSpPr><p:cNvPr id="2" name="Slide Number"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>
      <p:spPr><a:xfrm><a:off x="4226400" y="8586000"/><a:ext cx="1533600" cy="123111"/></a:xfrm></p:spPr>
      <p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:pPr algn="r"/><a:fld id="{{11111111-1111-1111-1111-111111111111}}" type="slidenum"><a:t>‹#›</a:t></a:fld></a:p></p:txBody>
      </p:sp>
    </p:spTree></p:cSld></p:sldMaster>"""
    slide_parts = {}
    for index in (1, 2):
        slide_parts[f"ppt/slides/slide{index}.xml"] = f"""
        <p:sld xmlns:p="{P}" xmlns:a="{A}"><p:cSld><p:spTree>
          <p:nvGrpSpPr><p:cNvPr id="1" name="Group"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/>
          {shape_xml(2, 'Header', '指南', 648000, 251999, 5112000, 162000, 'l')}
          {shape_xml(3, 'Footer', '机构 | 公开示例', 648000, 8586000, 3578400, 180000, 'l')}
        </p:spTree></p:cSld></p:sld>"""
    with zipfile.ZipFile(path, "w") as archive:
        archive.writestr("ppt/presentation.xml", presentation)
        archive.writestr("ppt/slideMasters/slideMaster1.xml", master)
        for name, data in slide_parts.items():
            archive.writestr(name, data)
        archive.writestr("ppt/media/keep.bin", b"unchanged")


class MirrorBoundPptxTest(unittest.TestCase):
    def test_mirrors_even_page_and_preserves_dynamic_numbers(self) -> None:
        from references.mirror_bound_pptx import mirror_copy

        with tempfile.TemporaryDirectory() as tmp:
            source = Path(tmp) / "source.pptx"
            output = Path(tmp) / "output.pptx"
            make_fixture(source)
            source_bytes = source.read_bytes()
            stats = mirror_copy(source, output, "指南", "机构 | 公开示例")
            self.assertEqual(stats["slides"], 2)
            self.assertEqual(source.read_bytes(), source_bytes)
            with zipfile.ZipFile(output) as archive:
                self.assertEqual(archive.read("ppt/media/keep.bin"), b"unchanged")
                master = etree.fromstring(archive.read("ppt/slideMasters/slideMaster1.xml"))
                self.assertFalse(master.xpath(".//a:fld[@type='slidenum']", namespaces=NS))
                for page in (1, 2):
                    slide = etree.fromstring(archive.read(f"ppt/slides/slide{page}.xml"))
                    numbers = slide.xpath(".//p:sp[.//a:fld[@type='slidenum']]", namespaces=NS)
                    self.assertEqual(len(numbers), 1)
                    header = slide.xpath(f".//p:sp[.//a:t='指南']", namespaces=NS)[0]
                    footer = slide.xpath(f".//p:sp[.//a:t='机构 | 公开示例']", namespaces=NS)[0]
                    expected = ("l", 648000, "r", 4226400) if page == 1 else ("r", 575713, "l", 575713)
                    self.assertEqual(header.xpath("string(.//a:pPr/@algn)", namespaces=NS), expected[0])
                    self.assertEqual(int(header.xpath("string(./p:spPr/a:xfrm/a:off/@x)", namespaces=NS)), expected[1])
                    self.assertEqual(footer.xpath("string(.//a:pPr/@algn)", namespaces=NS), expected[0])
                    self.assertEqual(numbers[0].xpath("string(.//a:pPr/@algn)", namespaces=NS), expected[2])
                    self.assertEqual(int(numbers[0].xpath("string(./p:spPr/a:xfrm/a:off/@x)", namespaces=NS)), expected[3])

    def test_refuses_to_overwrite_existing_output(self) -> None:
        from references.mirror_bound_pptx import mirror_copy

        with tempfile.TemporaryDirectory() as tmp:
            source = Path(tmp) / "source.pptx"
            output = Path(tmp) / "output.pptx"
            make_fixture(source)
            output.write_bytes(b"existing")
            with self.assertRaises(FileExistsError):
                mirror_copy(source, output, "指南", "机构 | 公开示例")
            self.assertEqual(output.read_bytes(), b"existing")


if __name__ == "__main__":
    unittest.main()
