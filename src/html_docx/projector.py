from __future__ import annotations

import html
import xml.etree.ElementTree as ET
from pathlib import Path
from typing import Any, Mapping
import re

from .utils import sha256_bytes


W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
M_NS = "http://schemas.openxmlformats.org/officeDocument/2006/math"
WP_NS = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
NS = {"w": W_NS, "m": M_NS, "wp": WP_NS}


def _tag(local: str) -> str:
    return f"{{{W_NS}}}{local}"


def _style_id(paragraph: ET.Element) -> str | None:
    p_pr = paragraph.find("w:pPr", NS)
    if p_pr is None:
        return None
    p_style = p_pr.find("w:pStyle", NS)
    if p_style is None:
        return None
    return p_style.attrib.get(_tag("val"))


def _paragraph_properties(paragraph: ET.Element) -> dict[str, str]:
    props: dict[str, str] = {}
    p_pr = paragraph.find("w:pPr", NS)
    if p_pr is None:
        return props
    jc = p_pr.find("w:jc", NS)
    if jc is not None and jc.attrib.get(_tag("val")):
        val = jc.attrib[_tag("val")]
        props["align"] = "justify" if val == "both" else val
    ind = p_pr.find("w:ind", NS)
    if ind is not None:
        first_line_chars = ind.attrib.get(_tag("firstLineChars"))
        if first_line_chars and first_line_chars.isdigit():
            props["first-line-indent"] = f"{int(first_line_chars) / 100:g}char"
        else:
            first_line = ind.attrib.get(_tag("firstLine"))
            if first_line and first_line.isdigit():
                props["first-line-indent"] = f"{int(first_line) / 20:g}pt"
    spacing = p_pr.find("w:spacing", NS)
    if spacing is not None:
        line_rule = spacing.attrib.get(_tag("lineRule"))
        line = spacing.attrib.get(_tag("line"))
        if line and line.isdigit():
            if line_rule in {None, "auto"}:
                props["line-spacing"] = f"{int(line) / 240:g}"
            elif line_rule in {"exact", "atLeast"}:
                props["line-spacing"] = f"{int(line) / 20:g}pt"
    return props


def _paragraph_numbering(paragraph: ET.Element) -> dict[str, str]:
    numbering: dict[str, str] = {}
    p_pr = paragraph.find("w:pPr", NS)
    if p_pr is None:
        return numbering
    num_pr = p_pr.find("w:numPr", NS)
    if num_pr is None:
        return numbering
    ilvl = num_pr.find("w:ilvl", NS)
    if ilvl is not None and ilvl.attrib.get(_tag("val")) is not None:
        numbering["ilvl"] = ilvl.attrib[_tag("val")]
    num_id = num_pr.find("w:numId", NS)
    if num_id is not None and num_id.attrib.get(_tag("val")) is not None:
        numbering["numId"] = num_id.attrib[_tag("val")]
    return numbering


def load_numbering_definitions(parts_dir: Path) -> dict[str, Any]:
    numbering_xml = parts_dir / "word" / "numbering.xml"
    if not numbering_xml.exists():
        return {"nums": {}, "abstractNums": {}}
    root = ET.parse(numbering_xml).getroot()
    return _numbering_definitions_from_root(root)


def load_numbering_definitions_from_bytes(numbering_xml: bytes | None) -> dict[str, Any]:
    if not numbering_xml:
        return {"nums": {}, "abstractNums": {}}
    root = ET.fromstring(numbering_xml)
    return _numbering_definitions_from_root(root)


def _numbering_definitions_from_root(root: ET.Element) -> dict[str, Any]:
    abstract_nums: dict[str, Any] = {}
    for abstract_num in root.findall("w:abstractNum", NS):
        abstract_id = abstract_num.attrib.get(_tag("abstractNumId"))
        if abstract_id is None:
            continue
        levels: dict[str, Any] = {}
        for level in abstract_num.findall("w:lvl", NS):
            ilvl = level.attrib.get(_tag("ilvl"))
            if ilvl is None:
                continue
            levels[ilvl] = _numbering_level_snapshot(level)
        abstract_nums[abstract_id] = {"abstractNumId": abstract_id, "levels": levels}

    nums: dict[str, Any] = {}
    for num in root.findall("w:num", NS):
        num_id = num.attrib.get(_tag("numId"))
        if num_id is None:
            continue
        abstract_num_id = None
        abstract_ref = num.find("w:abstractNumId", NS)
        if abstract_ref is not None:
            abstract_num_id = abstract_ref.attrib.get(_tag("val"))
        overrides: dict[str, Any] = {}
        for override in num.findall("w:lvlOverride", NS):
            ilvl = override.attrib.get(_tag("ilvl"))
            if ilvl is None:
                continue
            snapshot: dict[str, Any] = {}
            start_override = override.find("w:startOverride", NS)
            if start_override is not None and start_override.attrib.get(_tag("val")) is not None:
                snapshot["start"] = start_override.attrib[_tag("val")]
            override_level = override.find("w:lvl", NS)
            if override_level is not None:
                snapshot.update(_numbering_level_snapshot(override_level))
            overrides[ilvl] = snapshot
        nums[num_id] = {"numId": num_id, "abstractNumId": abstract_num_id, "overrides": overrides}
    return {"nums": nums, "abstractNums": abstract_nums}


def _numbering_level_snapshot(level: ET.Element) -> dict[str, Any]:
    snapshot: dict[str, Any] = {}
    for child_name, prop_name in (
        ("start", "start"),
        ("numFmt", "numFmt"),
        ("lvlText", "lvlText"),
        ("suff", "suffix"),
        ("pStyle", "pStyle"),
    ):
        child = level.find(f"w:{child_name}", NS)
        if child is not None and child.attrib.get(_tag("val")) is not None:
            snapshot[prop_name] = child.attrib[_tag("val")]
    p_pr = level.find("w:pPr", NS)
    if p_pr is not None:
        ind = p_pr.find("w:ind", NS)
        if ind is not None:
            indent: dict[str, str] = {}
            for attr_name in ("left", "hanging", "firstLine", "firstLineChars"):
                value = ind.attrib.get(_tag(attr_name))
                if value is not None:
                    indent[attr_name] = value
            if indent:
                snapshot["indent"] = indent
    return snapshot


def _resolve_numbering(
    numbering: dict[str, str],
    definitions: dict[str, Any],
) -> dict[str, str]:
    num_id = numbering.get("numId")
    ilvl = numbering.get("ilvl")
    if num_id is None or ilvl is None:
        return numbering
    resolved = dict(numbering)
    num_def = definitions.get("nums", {}).get(num_id)
    if not num_def:
        return resolved
    abstract_num_id = num_def.get("abstractNumId")
    if abstract_num_id is not None:
        resolved["abstractNumId"] = str(abstract_num_id)
    abstract_level = (
        definitions.get("abstractNums", {})
        .get(str(abstract_num_id), {})
        .get("levels", {})
        .get(ilvl, {})
    )
    merged = dict(abstract_level)
    merged.update(num_def.get("overrides", {}).get(ilvl, {}))
    for source_key, target_key in (
        ("numFmt", "numFmt"),
        ("lvlText", "lvlText"),
        ("start", "start"),
        ("suffix", "suffix"),
        ("pStyle", "pStyle"),
    ):
        if merged.get(source_key) is not None:
            resolved[target_key] = str(merged[source_key])
    indent = merged.get("indent")
    if isinstance(indent, dict):
        for key, value in indent.items():
            resolved[f"indent.{key}"] = str(value)
    return resolved


def _run_text(run: ET.Element) -> str:
    chunks: list[str] = []
    for child in list(run):
        if child.tag == _tag("t"):
            chunks.append(child.text or "")
        elif child.tag == _tag("tab"):
            chunks.append("\t")
        elif child.tag == _tag("br"):
            chunks.append("\n")
        elif child.tag == _tag("cr"):
            chunks.append("\n")
        elif child.tag == _tag("sym"):
            chunks.append("[symbol]")
        elif child.tag == _tag("drawing"):
            chunks.append("[drawing]")
        elif child.tag == _tag("object"):
            chunks.append("[object]")
        elif child.tag == _tag("fldChar"):
            chunks.append("[field]")
        elif child.tag == _tag("instrText"):
            chunks.append(child.text or "")
        elif child.tag == _tag("footnoteReference"):
            chunks.append("[footnote-ref]")
        elif child.tag == _tag("endnoteReference"):
            chunks.append("[endnote-ref]")
        elif child.tag == _tag("commentReference"):
            chunks.append("[comment-ref]")
        elif child.tag in {f"{{{M_NS}}}oMath", f"{{{M_NS}}}oMathPara"}:
            chunks.append("[equation]")
    return "".join(chunks)


def _run_lock(run: ET.Element) -> str:
    if run.find("w:drawing", NS) is not None:
        return "editable-metadata"
    protected_tags = {
        _tag("drawing"),
        _tag("object"),
        _tag("fldChar"),
        _tag("instrText"),
        _tag("pict"),
        _tag("footnoteReference"),
        _tag("endnoteReference"),
        _tag("commentReference"),
        _tag("lastRenderedPageBreak"),
        f"{{{M_NS}}}oMath",
        f"{{{M_NS}}}oMathPara",
    }
    if any(child.tag in protected_tags for child in list(run)):
        return "protected"
    return "editable"


def _drawing_properties(run: ET.Element) -> dict[str, str]:
    drawing = run.find("w:drawing", NS)
    if drawing is None:
        return {}
    props: dict[str, str] = {}
    doc_pr = drawing.find(".//wp:docPr", NS)
    if doc_pr is not None and doc_pr.attrib.get("descr"):
        props["alt"] = doc_pr.attrib["descr"]
    extent = drawing.find(".//wp:extent", NS)
    if extent is not None:
        width = extent.attrib.get("cx")
        height = extent.attrib.get("cy")
        if width and width.isdigit():
            props["width-emu"] = width
        if height and height.isdigit():
            props["height-emu"] = height
    return props


def _run_properties(run: ET.Element) -> dict[str, str]:
    return _rpr_properties(run.find("w:rPr", NS))


def _rpr_properties(r_pr: ET.Element | None) -> dict[str, str]:
    props: dict[str, str] = {}
    if r_pr is None:
        return props
    if _truthy_on_off(r_pr.find("w:b", NS)):
        props["bold"] = "true"
    if _truthy_on_off(r_pr.find("w:i", NS)):
        props["italic"] = "true"
    sz = r_pr.find("w:sz", NS)
    if sz is not None:
        raw = sz.attrib.get(_tag("val"))
        if raw and raw.isdigit():
            half_points = int(raw)
            props["font-size"] = f"{half_points / 2:g}pt"
    color = r_pr.find("w:color", NS)
    if color is not None:
        raw_color = color.attrib.get(_tag("val"))
        if raw_color and raw_color.lower() != "auto":
            props["color"] = f"#{raw_color.lower()}"
    return props


def load_style_definitions(parts_dir: Path) -> dict[str, Any]:
    styles_xml = parts_dir / "word" / "styles.xml"
    if not styles_xml.exists():
        return {}
    return load_style_definitions_from_bytes(styles_xml.read_bytes())


def load_style_definitions_from_bytes(styles_xml: bytes | None) -> dict[str, Any]:
    if not styles_xml:
        return {}
    root = ET.fromstring(styles_xml)
    styles: dict[str, Any] = {}
    for style in root.findall("w:style", NS):
        style_id = style.attrib.get(_tag("styleId"))
        if not style_id:
            continue
        snapshot: dict[str, Any] = {
            "styleId": style_id,
            "type": style.attrib.get(_tag("type")),
            "default": style.attrib.get(_tag("default")) == "1",
        }
        for child_name, prop_name in (
            ("name", "name"),
            ("basedOn", "basedOn"),
            ("next", "next"),
            ("link", "link"),
            ("uiPriority", "uiPriority"),
        ):
            child = style.find(f"w:{child_name}", NS)
            if child is not None and child.attrib.get(_tag("val")) is not None:
                snapshot[prop_name] = child.attrib[_tag("val")]
        if style.find("w:qFormat", NS) is not None:
            snapshot["qFormat"] = True
        paragraph_props = _paragraph_properties(style)
        if paragraph_props:
            snapshot["paragraphProperties"] = paragraph_props
        run_props = _rpr_properties(style.find("w:rPr", NS))
        if run_props:
            snapshot["runProperties"] = run_props
        styles[style_id] = snapshot
    return styles


def _truthy_on_off(element: ET.Element | None) -> bool:
    if element is None:
        return False
    raw = element.attrib.get(_tag("val"))
    return raw not in {"0", "false", "False", "off", "OFF"}


def project_main_document(parts_dir: Path) -> tuple[str, dict[str, Any]]:
    nodes: dict[str, Any] = {}
    context = {
        "paragraphSerial": 0,
        "runIndex": 0,
        "tableIndex": 0,
        "rowIndex": 0,
        "cellIndex": 0,
        "protectedIndex": 0,
        "numberingDefinitions": load_numbering_definitions(parts_dir),
    }
    articles: list[str] = []

    document_xml = parts_dir / "word" / "document.xml"
    if not document_xml.exists():
        return "<body></body>", nodes

    articles.append(_project_standard_part(document_xml, "/word/document.xml", "part-main", nodes, context))

    for part_path in _secondary_part_paths(parts_dir):
        rel_part = "/" + part_path.relative_to(parts_dir).as_posix()
        article_id = f"part-{_safe_class_suffix(rel_part.strip('/').replace('/', '-').replace('.', '-'))}"
        articles.append(_project_standard_part(part_path, rel_part, article_id, nodes, context))

    return "\n".join(articles) + "\n", nodes


def project_docx_entries(entry_bytes: Mapping[str, bytes]) -> tuple[str, dict[str, Any]]:
    nodes: dict[str, Any] = {}
    context = {
        "paragraphSerial": 0,
        "runIndex": 0,
        "tableIndex": 0,
        "rowIndex": 0,
        "cellIndex": 0,
        "protectedIndex": 0,
        "numberingDefinitions": load_numbering_definitions_from_bytes(entry_bytes.get("word/numbering.xml")),
    }
    document_xml = entry_bytes.get("word/document.xml")
    if document_xml is None:
        return "<body></body>", nodes
    articles = [
        _project_standard_root(ET.fromstring(document_xml), "/word/document.xml", "part-main", nodes, context)
    ]
    for entry_name in _secondary_entry_names(entry_bytes):
        rel_part = "/" + entry_name
        article_id = f"part-{_safe_class_suffix(rel_part.strip('/').replace('/', '-').replace('.', '-'))}"
        articles.append(_project_standard_root(ET.fromstring(entry_bytes[entry_name]), rel_part, article_id, nodes, context))
    return "\n".join(articles) + "\n", nodes


def _project_standard_part(
    xml_path: Path,
    part_path: str,
    article_id: str,
    nodes: dict[str, Any],
    context: dict[str, int],
) -> str:
    tree = ET.parse(xml_path)
    return _project_standard_root(tree.getroot(), part_path, article_id, nodes, context)


def _project_standard_root(
    root: ET.Element,
    part_path: str,
    article_id: str,
    nodes: dict[str, Any],
    context: dict[str, int],
) -> str:
    container = root.find("w:body", NS) or root

    lines: list[str] = [
        f'<article data-hdocx-type="part" data-hdocx-part="{html.escape(part_path, quote=True)}" data-hdocx-id="{article_id}">',
        f'  <section data-hdocx-type="section" data-hdocx-id="sec-{html.escape(article_id, quote=True)}">',
    ]
    context["partParagraphIndex"] = 0
    _project_children(list(container), lines, nodes, context, "    ", part_path)
    lines.extend(["  </section>", "</article>"])
    return "\n".join(lines)


def _secondary_entry_names(entry_bytes: Mapping[str, bytes]) -> list[str]:
    candidates: list[str] = []
    for entry_name in entry_bytes:
        if not entry_name.startswith("word/"):
            continue
        name = entry_name.rsplit("/", 1)[-1]
        if re.match(r"header\d+\.xml$", name) or re.match(r"footer\d+\.xml$", name):
            candidates.append(entry_name)
        elif name in {"footnotes.xml", "endnotes.xml", "comments.xml"}:
            candidates.append(entry_name)
    return sorted(candidates, key=lambda name: name.rsplit("/", 1)[-1])


def _secondary_part_paths(parts_dir: Path) -> list[Path]:
    word_dir = parts_dir / "word"
    if not word_dir.exists():
        return []
    candidates: list[Path] = []
    for child in word_dir.iterdir():
        name = child.name
        if not child.is_file():
            continue
        if re.match(r"header\d+\.xml$", name) or re.match(r"footer\d+\.xml$", name):
            candidates.append(child)
        elif name in {"footnotes.xml", "endnotes.xml", "comments.xml"}:
            candidates.append(child)
    return sorted(candidates, key=lambda path: path.name)


def _project_children(
    children: list[ET.Element],
    lines: list[str],
    nodes: dict[str, Any],
    context: dict[str, int],
    indent: str,
    part_path: str,
) -> None:
    for child in children:
        if child.tag == _tag("p"):
            _project_paragraph(child, lines, nodes, context, indent, part_path)
        elif child.tag == _tag("tbl"):
            _project_table(child, lines, nodes, context, indent, part_path)
        elif child.tag == _tag("comment"):
            comment_id = child.attrib.get(_tag("id"), "unknown")
            lines.append(f'{indent}<section data-hdocx-type="comment" data-hdocx-comment-id="{html.escape(comment_id, quote=True)}">')
            _project_protected_block(
                child,
                lines,
                nodes,
                context,
                indent + "  ",
                part_path,
                "comment",
                {"commentId": comment_id},
            )
            lines.append(f"{indent}</section>")
        elif child.tag in {_tag("footnote"), _tag("endnote")}:
            note_id = child.attrib.get(_tag("id"), "unknown")
            lines.append(f'{indent}<section data-hdocx-type="note" data-hdocx-note-id="{html.escape(note_id, quote=True)}">')
            _project_children(list(child), lines, nodes, context, indent + "  ", part_path)
            lines.append(f"{indent}</section>")


def _project_table(
    table: ET.Element,
    lines: list[str],
    nodes: dict[str, Any],
    context: dict[str, int],
    indent: str,
    part_path: str,
) -> None:
    context["tableIndex"] += 1
    table_id = f"tbl-{context['tableIndex']:06d}"
    table_index = context["tableIndex"]
    has_complex_merge = _table_has_complex_merge(table)
    nodes[table_id] = {
        "id": table_id,
        "kind": "table",
        "partPath": part_path,
        "lock": "editable",
        "locator": {"tableIndex": table_index},
        "simpleEditable": not has_complex_merge,
        "hasComplexMerge": has_complex_merge,
        "hash": sha256_bytes(ET.tostring(table, encoding="utf-8")),
        "children": [],
    }
    complex_attr = ' data-hdocx-complex-merge="true"' if has_complex_merge else ""
    lines.append(
        f'{indent}<table class="hdocx-table" data-hdocx-type="table" '
        f'data-hdocx-id="{table_id}" data-hdocx-lock="editable" '
        f'data-hdocx-part="{html.escape(part_path, quote=True)}"{complex_attr}>'
    )
    for row_in_table_index, row in enumerate(table.findall("w:tr", NS), start=1):
        context["rowIndex"] += 1
        row_id = f"tr-{context['rowIndex']:06d}"
        nodes[table_id]["children"].append(row_id)
        nodes[row_id] = {
            "id": row_id,
            "kind": "table-row",
            "partPath": part_path,
            "parent": table_id,
            "lock": "editable",
            "locator": {
                "tableIndex": table_index,
                "rowIndex": context["rowIndex"],
                "rowInTableIndex": row_in_table_index,
            },
            "simpleEditable": not has_complex_merge,
            "hash": sha256_bytes(ET.tostring(row, encoding="utf-8")),
            "children": [],
        }
        lines.append(
            f'{indent}  <tr class="hdocx-row" data-hdocx-type="table-row" data-hdocx-id="{row_id}" '
            f'data-hdocx-lock="editable" data-hdocx-part="{html.escape(part_path, quote=True)}">'
        )
        for cell_in_row_index, cell in enumerate(row.findall("w:tc", NS), start=1):
            context["cellIndex"] += 1
            cell_id = f"tc-{context['cellIndex']:06d}"
            nodes[row_id]["children"].append(cell_id)
            cell_props = _table_cell_properties(cell)
            nodes[cell_id] = {
                "id": cell_id,
                "kind": "table-cell",
                "partPath": part_path,
                "parent": row_id,
                "lock": "editable",
                "locator": {
                    "tableIndex": table_index,
                    "rowInTableIndex": row_in_table_index,
                    "cellIndex": context["cellIndex"],
                    "cellInRowIndex": cell_in_row_index,
                },
                "properties": cell_props,
                "simpleEditable": not cell_props.get("grid-span") and not cell_props.get("v-merge"),
                "hash": sha256_bytes(ET.tostring(cell, encoding="utf-8")),
                "children": [],
            }
            cell_attr = "".join(
                f' data-hdocx-{name}="{html.escape(value, quote=True)}"'
                for name, value in sorted(cell_props.items())
            )
            lines.append(
                f'{indent}    <td class="hdocx-cell" data-hdocx-type="table-cell" '
                f'data-hdocx-id="{cell_id}" data-hdocx-lock="editable" '
                f'data-hdocx-part="{html.escape(part_path, quote=True)}"{cell_attr}>'
            )
            for child in list(cell):
                if child.tag == _tag("p"):
                    p_id = _project_paragraph(child, lines, nodes, context, indent + "      ", part_path)
                    nodes[cell_id]["children"].append(p_id)
                elif child.tag == _tag("tbl"):
                    _project_table(child, lines, nodes, context, indent + "      ", part_path)
            lines.append(f"{indent}    </td>")
        lines.append(f"{indent}  </tr>")
    lines.append(f"{indent}</table>")


def _table_has_complex_merge(table: ET.Element) -> bool:
    return any(_table_cell_properties(cell) for cell in table.findall(".//w:tc", NS))


def _table_cell_properties(cell: ET.Element) -> dict[str, str]:
    props: dict[str, str] = {}
    tc_pr = cell.find("w:tcPr", NS)
    if tc_pr is None:
        return props
    grid_span = tc_pr.find("w:gridSpan", NS)
    if grid_span is not None and grid_span.attrib.get(_tag("val")) is not None:
        props["grid-span"] = grid_span.attrib[_tag("val")]
    v_merge = tc_pr.find("w:vMerge", NS)
    if v_merge is not None:
        props["v-merge"] = v_merge.attrib.get(_tag("val"), "continue")
    return props


def _project_paragraph(
    paragraph: ET.Element,
    lines: list[str],
    nodes: dict[str, Any],
    context: dict[str, int],
    indent: str,
    part_path: str,
) -> str:
    context["paragraphSerial"] += 1
    context["partParagraphIndex"] += 1
    paragraph_index = context["partParagraphIndex"]
    p_id = f"p-{context['paragraphSerial']:06d}"
    style_id = _style_id(paragraph)
    paragraph_properties = _paragraph_properties(paragraph)
    paragraph_numbering = _resolve_numbering(_paragraph_numbering(paragraph), context["numberingDefinitions"])
    class_names = ["hdocx-p"]
    if style_id:
        class_names.append(f"hstyle-{_safe_class_suffix(style_id)}")
    if paragraph_numbering:
        class_names.append("hdocx-list")
    nodes[p_id] = {
        "id": p_id,
        "kind": "paragraph",
        "partPath": part_path,
        "styleId": style_id,
        "lock": "editable",
        "locator": {
            "paragraphIndex": paragraph_index,
        },
        "properties": paragraph_properties,
        "numbering": paragraph_numbering,
        "hash": sha256_bytes(ET.tostring(paragraph, encoding="utf-8")),
        "children": [],
    }
    style_attr = f' data-hdocx-style-id="{html.escape(style_id, quote=True)}"' if style_id else ""
    paragraph_property_attrs = "".join(
        f' data-hdocx-{name}="{html.escape(value, quote=True)}"'
        for name, value in sorted(paragraph_properties.items())
    )
    numbering_attrs = ""
    numbering_attr_names = {
        "numId": "data-hdocx-num-id",
        "ilvl": "data-hdocx-ilvl",
        "abstractNumId": "data-hdocx-abstract-num-id",
        "numFmt": "data-hdocx-num-format",
        "lvlText": "data-hdocx-level-text",
        "start": "data-hdocx-start",
        "suffix": "data-hdocx-number-suffix",
        "pStyle": "data-hdocx-num-style-id",
        "indent.left": "data-hdocx-num-indent-left",
        "indent.hanging": "data-hdocx-num-indent-hanging",
        "indent.firstLine": "data-hdocx-num-indent-first-line",
        "indent.firstLineChars": "data-hdocx-num-indent-first-line-chars",
    }
    for key, attr_name in numbering_attr_names.items():
        value = paragraph_numbering.get(key)
        if value is not None:
            numbering_attrs += f' {attr_name}="{html.escape(value, quote=True)}"'
    lines.append(
        f'{indent}<p class="{" ".join(class_names)}" data-hdocx-type="paragraph" '
        f'data-hdocx-id="{p_id}" data-hdocx-lock="editable" '
        f'data-hdocx-part="{html.escape(part_path, quote=True)}"{style_attr}{paragraph_property_attrs}{numbering_attrs}>'
    )
    run_in_paragraph_index = 0
    protected_in_paragraph_index = 0
    for child in list(paragraph):
        if child.tag == _tag("pPr"):
            continue
        if child.tag != _tag("r"):
            if _protected_inline_kind(child) is not None:
                protected_in_paragraph_index += 1
            protected_id = _project_protected_inline(
                child,
                lines,
                nodes,
                context,
                indent + "  ",
                part_path,
                p_id,
                paragraph_index,
                protected_in_paragraph_index,
            )
            if protected_id:
                nodes[p_id]["children"].append(protected_id)
            continue
        run = child
        context["runIndex"] += 1
        run_index = context["runIndex"]
        run_in_paragraph_index += 1
        r_id = f"r-{run_index:06d}"
        lock = _run_lock(run)
        text = _run_text(run)
        properties = _run_properties(run)
        drawing_properties = _drawing_properties(run)
        text_nodes = [child for child in list(run) if child.tag == _tag("t")]
        simple_editable = lock == "editable" and len(text_nodes) == 1 and len(list(run)) in {1, 2}
        if len(list(run)) == 2 and list(run)[0].tag != _tag("rPr"):
            simple_editable = False
        nodes[p_id]["children"].append(r_id)
        nodes[r_id] = {
            "id": r_id,
            "kind": "run",
            "partPath": part_path,
            "parent": p_id,
            "locator": {
                "paragraphIndex": paragraph_index,
                "runIndex": run_index,
                "runInParagraphIndex": run_in_paragraph_index,
                "textNodeIndex": 1 if text_nodes else None,
            },
            "lock": lock,
            "objectKind": "drawing" if drawing_properties or run.find("w:drawing", NS) is not None else None,
            "simpleEditable": simple_editable,
            "text": text,
            "properties": properties,
            "objectProperties": drawing_properties,
            "textHash": sha256_bytes(text.encode("utf-8")),
            "hash": sha256_bytes(ET.tostring(run, encoding="utf-8")),
            "children": [],
        }
        property_attrs = "".join(
            f' data-hdocx-{name}="{html.escape(value, quote=True)}"'
            for name, value in sorted(properties.items())
        )
        object_property_attrs = "".join(
            f' data-hdocx-{name}="{html.escape(value, quote=True)}"'
            for name, value in sorted(drawing_properties.items())
        )
        lines.append(
            f'{indent}  <span class="hdocx-r hlock-{lock}" data-hdocx-type="run" '
            f'data-hdocx-id="{r_id}" data-hdocx-lock="{lock}" '
            f'data-hdocx-part="{html.escape(part_path, quote=True)}"{property_attrs}{object_property_attrs}>'
            f"{html.escape(text)}</span>"
        )
    lines.append(f"{indent}</p>")
    return p_id


def _project_protected_inline(
    element: ET.Element,
    lines: list[str],
    nodes: dict[str, Any],
    context: dict[str, int],
    indent: str,
    part_path: str,
    parent_id: str,
    paragraph_index: int,
    protected_in_paragraph_index: int,
) -> str | None:
    kind = _protected_inline_kind(element)
    if kind is None:
        return None
    context["protectedIndex"] += 1
    protected_id = f"prot-{context['protectedIndex']:06d}"
    text = _protected_placeholder(element, kind)
    nodes[protected_id] = {
        "id": protected_id,
        "kind": "protected",
        "protectedKind": kind,
        "partPath": part_path,
        "parent": parent_id,
        "locator": {
            "paragraphIndex": paragraph_index,
            "protectedInParagraphIndex": protected_in_paragraph_index,
            "protectedKind": kind,
            "sourceId": element.attrib.get(_tag("id")) or element.attrib.get(_tag("name")),
        },
        "lock": "protected",
        "text": text,
        "hash": sha256_bytes(ET.tostring(element, encoding="utf-8")),
        "children": [],
    }
    lines.append(
        f'{indent}<span class="hdocx-protected hlock-protected" data-hdocx-type="protected" '
        f'data-hdocx-id="{protected_id}" data-hdocx-lock="protected" '
        f'data-hdocx-part="{html.escape(part_path, quote=True)}" '
        f'data-hdocx-protected-kind="{html.escape(kind, quote=True)}">'
        f"{html.escape(text)}</span>"
    )
    return protected_id


def _protected_inline_kind(element: ET.Element) -> str | None:
    mapping = {
        _tag("commentRangeStart"): "comment-range-start",
        _tag("commentRangeEnd"): "comment-range-end",
        _tag("bookmarkStart"): "bookmark-start",
        _tag("bookmarkEnd"): "bookmark-end",
        _tag("permStart"): "permission-start",
        _tag("permEnd"): "permission-end",
        _tag("proofErr"): "proof-error",
        _tag("customXml"): "custom-xml",
        _tag("sdt"): "content-control",
        _tag("ins"): "revision-insert",
        _tag("del"): "revision-delete",
        _tag("moveFrom"): "revision-move-from",
        _tag("moveTo"): "revision-move-to",
        f"{{{M_NS}}}oMath": "equation",
        f"{{{M_NS}}}oMathPara": "equation",
    }
    return mapping.get(element.tag)


def _protected_placeholder(element: ET.Element, kind: str) -> str:
    value = element.attrib.get(_tag("id")) or element.attrib.get(_tag("name"))
    if value:
        return f"[{kind}:{value}]"
    return f"[{kind}]"


def _project_protected_block(
    element: ET.Element,
    lines: list[str],
    nodes: dict[str, Any],
    context: dict[str, int],
    indent: str,
    part_path: str,
    kind: str,
    locator: dict[str, str] | None = None,
) -> str:
    context["protectedIndex"] += 1
    protected_id = f"prot-{context['protectedIndex']:06d}"
    text = _protected_block_text(element, kind)
    nodes[protected_id] = {
        "id": protected_id,
        "kind": "protected",
        "protectedKind": kind,
        "partPath": part_path,
        "locator": locator or {},
        "lock": "protected",
        "text": text,
        "hash": sha256_bytes(ET.tostring(element, encoding="utf-8")),
        "children": [],
    }
    lines.append(
        f'{indent}<div class="hdocx-protected hlock-protected" data-hdocx-type="protected" '
        f'data-hdocx-id="{protected_id}" data-hdocx-lock="protected" '
        f'data-hdocx-part="{html.escape(part_path, quote=True)}" '
        f'data-hdocx-protected-kind="{html.escape(kind, quote=True)}">'
        f"{html.escape(text)}</div>"
    )
    return protected_id


def _protected_block_text(element: ET.Element, kind: str) -> str:
    texts: list[str] = []
    for text_node in element.findall(".//w:t", NS):
        if text_node.text:
            texts.append(text_node.text)
    joined = "".join(texts).strip()
    if joined:
        return f"[{kind}] {joined}"
    return f"[{kind}]"


def build_document_html(body: str) -> str:
    return "\n".join(
        [
            "<!doctype html>",
            '<html data-hdocx-version="0.1">',
            "  <head>",
            '    <meta charset="utf-8">',
            "    <title>H-DOCX Document</title>",
            '    <link rel="stylesheet" href="styles.generated.css" data-hdocx-role="preview">',
            '    <link rel="stylesheet" href="agent.edits.hcss" data-hdocx-role="editable-style-requests">',
            "  </head>",
            "  <body>",
            body.rstrip(),
            "  </body>",
            "</html>",
            "",
        ]
    )


def _safe_class_suffix(value: str) -> str:
    safe = []
    for ch in value:
        if ch.isalnum() or ch in ("-", "_"):
            safe.append(ch)
        else:
            safe.append("_")
    return "".join(safe) or "unnamed"
