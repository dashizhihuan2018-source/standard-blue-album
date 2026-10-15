#!/usr/bin/env python3
"""Create a non-overwriting PPTX copy with mirrored odd/even folio furniture."""

from __future__ import annotations

import argparse
import copy
import os
import re
import tempfile
import zipfile
from pathlib import Path

from lxml import etree


P_NS = "http://schemas.openxmlformats.org/presentationml/2006/main"
A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main"
NS = {"p": P_NS, "a": A_NS}


def _qname(namespace: str, local: str) -> str:
    return f"{{{namespace}}}{local}"


def _text(shape: etree._Element) -> str:
    return "".join(shape.xpath(".//a:t/text()", namespaces=NS)).strip()


def _geometry(shape: etree._Element) -> tuple[int, int, int, int]:
    xfrm = shape.find("./p:spPr/a:xfrm", NS)
    if xfrm is None:
        raise ValueError("Text shape has no transform")
    off = xfrm.find("a:off", NS)
    ext = xfrm.find("a:ext", NS)
    if off is None or ext is None:
        raise ValueError("Text shape transform is incomplete")
    return tuple(int(value) for value in (off.get("x"), off.get("y"), ext.get("cx"), ext.get("cy")))


def _mirror(shape: etree._Element, slide_width: int, alignment: str) -> None:
    x, _, cx, _ = _geometry(shape)
    shape.find("./p:spPr/a:xfrm/a:off", NS).set("x", str(slide_width - x - cx))
    for paragraph in shape.xpath("./p:txBody/a:p", namespaces=NS):
        paragraph_properties = paragraph.find("a:pPr", NS)
        if paragraph_properties is None:
            paragraph_properties = etree.Element(_qname(A_NS, "pPr"))
            paragraph.insert(0, paragraph_properties)
        paragraph_properties.set("algn", alignment)


def mirror_copy(
    source: Path,
    output: Path,
    header_text: str,
    footer_text: str,
    first_folio: int = 1,
) -> dict[str, int]:
    source = Path(source).expanduser().resolve()
    output = Path(output).expanduser().resolve()
    if output.exists():
        raise FileExistsError(f"Refusing to overwrite existing output: {output}")
    if source == output:
        raise ValueError("Source and output must be different files")

    with zipfile.ZipFile(source, "r") as archive:
        entries = [(info, archive.read(info.filename)) for info in archive.infolist()]
    entry_map = {info.filename: data for info, data in entries}

    presentation = etree.fromstring(entry_map["ppt/presentation.xml"])
    slide_size = presentation.find("p:sldSz", NS)
    slide_width = int(slide_size.get("cx"))
    slide_height = int(slide_size.get("cy"))

    master_names = sorted(
        name for name in entry_map if re.fullmatch(r"ppt/slideMasters/slideMaster\d+\.xml", name)
    )
    if len(master_names) != 1:
        raise RuntimeError(
            "This helper requires exactly one slide master. Inspect multi-master decks and adapt their layouts explicitly."
        )
    master_name = master_names[0]
    master = etree.fromstring(entry_map[master_name])
    number_shapes = [
        shape for shape in master.xpath(".//p:sp", namespaces=NS)
        if shape.xpath(".//a:fld[@type='slidenum']", namespaces=NS)
    ]
    if len(number_shapes) != 1:
        raise RuntimeError(f"Expected one master slide-number field, found {len(number_shapes)}")
    number_template = copy.deepcopy(number_shapes[0])
    number_shapes[0].getparent().remove(number_shapes[0])
    entry_map[master_name] = etree.tostring(master, xml_declaration=True, encoding="UTF-8", standalone=True)

    slide_names = sorted(
        (name for name in entry_map if re.fullmatch(r"ppt/slides/slide\d+\.xml", name)),
        key=lambda name: int(re.search(r"\d+", name).group()),
    )
    stats = {"slides": 0, "headers_mirrored": 0, "footers_mirrored": 0, "numbers_added": 0}
    for sequence, slide_name in enumerate(slide_names):
        folio = first_folio + sequence
        slide = etree.fromstring(entry_map[slide_name])
        shape_tree = slide.find("./p:cSld/p:spTree", NS)
        if shape_tree is None:
            raise RuntimeError(f"Missing shape tree: {slide_name}")

        if folio % 2 == 0:
            for shape in shape_tree.xpath(".//p:sp", namespaces=NS):
                value = _text(shape)
                _, y, _, _ = _geometry(shape)
                if value == header_text and y < slide_height * 0.08:
                    _mirror(shape, slide_width, "r")
                    stats["headers_mirrored"] += 1
                elif value == footer_text and y > slide_height * 0.90:
                    _mirror(shape, slide_width, "r")
                    stats["footers_mirrored"] += 1

        number_shape = copy.deepcopy(number_template)
        ids = [int(value) for value in slide.xpath(".//p:cNvPr/@id", namespaces=NS)]
        c_nv_pr = number_shape.find("./p:nvSpPr/p:cNvPr", NS)
        c_nv_pr.set("id", str(max(ids, default=1) + 1))
        c_nv_pr.set("name", f"装订页码 {folio}")
        if folio % 2 == 0:
            _mirror(number_shape, slide_width, "l")
        shape_tree.append(number_shape)
        stats["numbers_added"] += 1
        stats["slides"] += 1
        entry_map[slide_name] = etree.tostring(slide, xml_declaration=True, encoding="UTF-8", standalone=True)

    output.parent.mkdir(parents=True, exist_ok=True)
    with tempfile.NamedTemporaryFile(dir=output.parent, prefix=output.stem + ".", suffix=".tmp", delete=False) as tmp:
        tmp_path = Path(tmp.name)
    try:
        with zipfile.ZipFile(tmp_path, "w") as archive:
            for info, original_data in entries:
                archive.writestr(info, entry_map.get(info.filename, original_data))
        tmp_path.replace(output)
        os.chmod(output, source.stat().st_mode)
    finally:
        tmp_path.unlink(missing_ok=True)
    return stats


def main() -> None:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("source", type=Path)
    parser.add_argument("output", type=Path)
    parser.add_argument("--header-text", required=True)
    parser.add_argument("--footer-text", required=True)
    parser.add_argument("--first-folio", type=int, default=1)
    args = parser.parse_args()
    stats = mirror_copy(args.source, args.output, args.header_text, args.footer_text, args.first_folio)
    print(" ".join(f"{key}={value}" for key, value in stats.items()))


if __name__ == "__main__":
    main()
