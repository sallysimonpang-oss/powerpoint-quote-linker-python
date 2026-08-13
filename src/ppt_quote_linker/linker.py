"""Add query hyperlinks to quoted SmartArt text without PowerPoint."""

from __future__ import annotations

from collections.abc import Callable, Mapping
from copy import deepcopy
from pathlib import Path
import re
import shutil
from typing import TypeAlias
from urllib.parse import quote_plus
from zipfile import ZIP_DEFLATED, ZipFile, ZipInfo

from lxml import etree

A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main"
R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
HYPERLINK_REL = f"{R_NS}/hyperlink"
NS = {"a": A_NS, "r": R_NS}
TargetSource: TypeAlias = str | Mapping[str, str] | Callable[[str, str], str]
_CITATION = re.compile(r"\s*\([^()\r\n]*\)")


def link_presentation(input_path: str | Path, output_path: str | Path, target: TargetSource) -> int:
    """Link quotes in the first quote-bearing paragraph of each SmartArt part.

    ``target`` is a URL template containing ``{query}``, a mapping from diagram
    data part to template, or a callback accepting ``(part_name, search_text)``.
    Returns the number of logical quotation hyperlinks added.
    """
    source, destination = Path(input_path), Path(output_path)
    if source.resolve() == destination.resolve():
        raise ValueError("input_path and output_path must be different")
    replacements: dict[str, bytes] = {}
    logical_count = 0
    with ZipFile(source, "r") as package:
        names = set(package.namelist())
        parts = sorted(n for n in names if re.fullmatch(r"ppt/diagrams/data\d+\.xml", n))
        for data_part in parts:
            number = re.search(r"(\d+)\.xml$", data_part).group(1)
            drawing_part = f"ppt/diagrams/drawing{number}.xml"
            result = _transform_part(package.read(data_part), data_part, target)
            if not result.changed:
                continue
            replacements[data_part] = result.xml
            replacements[_rels_name(data_part)] = _updated_relationships(package, data_part, result.relationships)
            logical_count += len(result.relationships)
            if drawing_part in names:
                drawing = _transform_part(package.read(drawing_part), data_part, target)
                if drawing.changed:
                    replacements[drawing_part] = drawing.xml
                    replacements[_rels_name(drawing_part)] = _updated_relationships(package, drawing_part, drawing.relationships)
        if not replacements:
            shutil.copyfile(source, destination)
            return 0
        destination.parent.mkdir(parents=True, exist_ok=True)
        with ZipFile(destination, "w") as output:
            for info in package.infolist():
                output.writestr(_copy_info(info), replacements.pop(info.filename, package.read(info.filename)))
            for name in sorted(replacements):
                output.writestr(name, replacements[name], compress_type=ZIP_DEFLATED)
    return logical_count


class _PartResult:
    def __init__(self, xml: bytes, relationships: list[tuple[str, str]], changed: bool):
        self.xml, self.relationships, self.changed = xml, relationships, changed


def _transform_part(xml: bytes, target_part: str, target: TargetSource) -> _PartResult:
    root = etree.fromstring(xml, etree.XMLParser(remove_blank_text=False))
    for paragraph in root.xpath(".//a:p", namespaces=NS):
        ranges = _qualifying_quotes(paragraph)
        if not ranges:
            continue
        relationships, changed = [], False
        for index, (start, end, search_text) in enumerate(ranges, 1):
            url = _resolve_target(target, target_part, search_text)
            existing_id = _range_hyperlink_id(paragraph, start, end)
            relationship_id = existing_id or f"rId{index}"
            if existing_id is None:
                _apply_hyperlink(paragraph, start, end, relationship_id)
                changed = True
            relationships.append((relationship_id, url))
        if not changed:
            return _PartResult(xml, relationships, False)
        return _PartResult(etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True), relationships, True)
    return _PartResult(xml, [], False)


def _qualifying_quotes(paragraph: etree._Element) -> list[tuple[int, int, str]]:
    runs = paragraph.xpath("./a:r", namespaces=NS)
    text = "".join("".join(r.xpath("./a:t/text()", namespaces=NS)) for r in runs)
    bold_ranges, position = [], 0
    for run in runs:
        run_text = "".join(run.xpath("./a:t/text()", namespaces=NS))
        rpr = run.find(f"{{{A_NS}}}rPr")
        if rpr is not None and rpr.get("b") in {"1", "true"} and run_text:
            bold_ranges.append((position, position + len(run_text), run_text))
        position += len(run_text)
    result, cursor = [], 0
    while True:
        opening = text.find("“", cursor)
        if opening < 0:
            break
        closing = text.find("”", opening + 1)
        if closing < 0:
            break
        first_bold = next((x for x in bold_ranges if x[0] < closing and x[1] > opening + 1), None)
        if first_bold and first_bold[2].strip():
            end = closing + 1
            citation = _CITATION.match(text, end)
            if citation:
                end = citation.end()
            result.append((opening, end, first_bold[2].strip()))
        cursor = closing + 1
    return result


def _apply_hyperlink(paragraph: etree._Element, start: int, end: int, relationship_id: str) -> None:
    position = 0
    for run in list(paragraph.xpath("./a:r", namespaces=NS)):
        node = run.find(f"{{{A_NS}}}t")
        text = "" if node is None or node.text is None else node.text
        run_start, run_end = position, position + len(text)
        position = run_end
        left, right = max(start, run_start), min(end, run_end)
        if left >= right:
            continue
        pieces = []
        if run_start < left:
            pieces.append((text[: left - run_start], False))
        pieces.append((text[left - run_start : right - run_start], True))
        if right < run_end:
            pieces.append((text[right - run_start :], False))
        parent, index = run.getparent(), run.getparent().index(run)
        parent.remove(run)
        for offset, (piece, linked) in enumerate(pieces):
            clone = deepcopy(run)
            clone.find(f"{{{A_NS}}}t").text = piece
            if linked:
                rpr = clone.find(f"{{{A_NS}}}rPr")
                if rpr is None:
                    rpr = etree.Element(f"{{{A_NS}}}rPr")
                    clone.insert(0, rpr)
                for old in rpr.findall(f"{{{A_NS}}}hlinkClick"):
                    rpr.remove(old)
                link = etree.Element(f"{{{A_NS}}}hlinkClick")
                link.set(f"{{{R_NS}}}id", relationship_id)
                rpr.append(link)
            parent.insert(index + offset, clone)


def _range_hyperlink_id(paragraph: etree._Element, start: int, end: int) -> str | None:
    ids, covered, position = set(), 0, 0
    for run in paragraph.xpath("./a:r", namespaces=NS):
        text = "".join(run.xpath("./a:t/text()", namespaces=NS))
        run_start, run_end = position, position + len(text)
        position = run_end
        overlap = max(0, min(end, run_end) - max(start, run_start))
        if not overlap:
            continue
        links = run.xpath("./a:rPr/a:hlinkClick/@r:id", namespaces=NS)
        if len(links) != 1:
            return None
        ids.add(links[0]); covered += overlap
    return ids.pop() if covered == end - start and len(ids) == 1 else None


def _resolve_target(target: TargetSource, part: str, search_text: str) -> str:
    if callable(target):
        return target(part, search_text)
    template = target[part] if isinstance(target, Mapping) else target
    return template.format(query=quote_plus(search_text, safe=""))


def _rels_name(part: str) -> str:
    path = Path(part)
    return str(path.parent / "_rels" / f"{path.name}.rels").replace("\\", "/")


def _updated_relationships(package: ZipFile, part: str, relationships: list[tuple[str, str]]) -> bytes:
    name = _rels_name(part)
    root = etree.fromstring(package.read(name)) if name in package.namelist() else etree.Element(f"{{{REL_NS}}}Relationships", nsmap={None: REL_NS})
    by_id = {element.get("Id"): element for element in root}
    for relationship_id, url in relationships:
        element = by_id.get(relationship_id)
        if element is None:
            element = etree.SubElement(root, f"{{{REL_NS}}}Relationship")
            element.set("Id", relationship_id)
        element.set("Type", HYPERLINK_REL); element.set("Target", url); element.set("TargetMode", "External")
    return etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)


def _copy_info(info: ZipInfo) -> ZipInfo:
    clone = ZipInfo(info.filename, info.date_time)
    for attr in ("compress_type", "comment", "extra", "internal_attr", "external_attr", "create_system", "create_version", "extract_version", "flag_bits", "volume"):
        setattr(clone, attr, getattr(info, attr))
    return clone
