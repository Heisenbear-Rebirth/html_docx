from __future__ import annotations

import xml.etree.ElementTree as ET
import copy
import html
import re
from typing import Any

from .errors import HDocxError
from .projector import M_NS, NS, W_NS, WP_NS, _tag
from .utils import sha256_bytes


XML_NS = "http://www.w3.org/XML/1998/namespace"
R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main"
PIC_NS = "http://schemas.openxmlformats.org/drawingml/2006/picture"
CT_NS = "http://schemas.openxmlformats.org/package/2006/content-types"
PKG_REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
IMAGE_REL_TYPE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image"


ET.register_namespace("w", W_NS)
ET.register_namespace("m", M_NS)
ET.register_namespace("wp", WP_NS)
ET.register_namespace("r", R_NS)
ET.register_namespace("a", A_NS)
ET.register_namespace("pic", PIC_NS)
ET.register_namespace("", CT_NS)


def patch_document_run_text(document_xml: bytes, edits: list[dict[str, Any]]) -> bytes:
    fragment_patched = _try_fragment_patch(document_xml, edits)
    if fragment_patched is not None:
        return fragment_patched

    root = ET.fromstring(document_xml)
    paragraphs = root.findall(".//w:p", NS)

    sorted_edits = sorted(edits, key=_edit_sort_key, reverse=True)

    for edit in sorted_edits:
        if edit["operation"] == "patch-style":
            _patch_style(root, edit)
            continue
        if edit["operation"] == "create-style":
            _create_style(root, edit)
            continue
        if edit["operation"] == "delete-style":
            _delete_style(root, edit)
            continue
        if edit["operation"] == "patch-content-types":
            _patch_content_types(root, edit)
            continue
        if edit["operation"] == "patch-relationships":
            _patch_relationships(root, edit)
            continue
        if edit["operation"] == "patch-numbering-level":
            _patch_numbering_level(root, edit)
            continue
        if edit["operation"] == "create-numbering-list":
            _create_numbering_list(root, edit)
            continue
        if edit["operation"] == "patch-comment-text":
            _patch_comment_text(root, edit)
            continue
        if edit["operation"] == "patch-revision-action":
            _patch_revision_action(root, edit)
            continue
        if edit["operation"] == "patch-equation-omml":
            locator = edit["locator"]
            paragraph_index = locator["paragraphIndex"]
            try:
                paragraph = paragraphs[paragraph_index - 1]
            except IndexError as exc:
                raise HDocxError(
                    "PATCH_PARAGRAPH_LOCATOR_NOT_FOUND",
                    "Paragraph locator no longer resolves in document.xml.",
                    {"editId": edit["id"], "paragraphIndex": paragraph_index},
                ) from exc
            _patch_equation_omml(paragraph, edit)
            continue
        if edit["operation"] == "insert-table-row-after":
            _insert_table_row_after(root, edit)
            continue
        if edit["operation"] == "delete-table-row":
            _delete_table_row(root, edit)
            continue
        if edit["operation"] == "insert-table-column-after":
            _insert_table_column_after(root, edit)
            continue
        if edit["operation"] == "delete-table-column":
            _delete_table_column(root, edit)
            continue
        locator = edit["locator"]
        paragraph_index = locator["paragraphIndex"]
        try:
            paragraph = paragraphs[paragraph_index - 1]
        except IndexError as exc:
            raise HDocxError(
                "PATCH_PARAGRAPH_LOCATOR_NOT_FOUND",
                "Paragraph locator no longer resolves in document.xml.",
                {"editId": edit["id"], "paragraphIndex": paragraph_index},
            ) from exc

        if edit["operation"] == "patch-paragraph":
            _patch_paragraph_properties(paragraph, edit.get("newProperties", {}))
            continue

        if edit["operation"] == "patch-paragraph-style":
            _patch_paragraph_style(paragraph, edit["newStyleId"])
            continue

        if edit["operation"] == "patch-paragraph-numbering":
            _patch_paragraph_numbering(paragraph, edit["numId"], edit["ilvl"])
            continue

        if edit["operation"] in {"insert-image-after-paragraph", "insert-image-before-paragraph"}:
            position = "before" if edit["operation"] == "insert-image-before-paragraph" else "after"
            _insert_image_around_paragraph(root, paragraph, edit, position)
            continue

        run_in_paragraph_index = locator["runInParagraphIndex"]
        split_segments = edit.get("splitSegments")
        new_text = edit.get("newText")
        new_properties = edit.get("newProperties", {})
        runs = paragraph.findall("w:r", NS)
        try:
            run = runs[run_in_paragraph_index - 1]
        except IndexError as exc:
            raise HDocxError(
                "PATCH_RUN_LOCATOR_NOT_FOUND",
                "Run locator no longer resolves in document.xml.",
                {"editId": edit["id"], "runInParagraphIndex": run_in_paragraph_index},
            ) from exc

        if edit["operation"] in {"patch-drawing-alt", "patch-drawing-properties"}:
            _patch_drawing_properties(run, edit)
            continue

        if split_segments:
            _replace_run_with_split_segments(paragraph, run, split_segments, edit)
            continue

        if new_text is not None:
            text_nodes = [child for child in list(run) if child.tag == _tag("t")]
            if len(text_nodes) != 1:
                raise HDocxError(
                    "PATCH_RUN_NOT_SIMPLE_TEXT",
                    "Only runs with exactly one w:t text node are supported in this phase.",
                    {"editId": edit["id"], "nodeId": edit["nodeId"]},
                )
            text_node = text_nodes[0]
            text_node.text = new_text
            if _needs_preserve_space(new_text):
                text_node.attrib[f"{{{XML_NS}}}space"] = "preserve"
            elif f"{{{XML_NS}}}space" in text_node.attrib:
                del text_node.attrib[f"{{{XML_NS}}}space"]

        if new_properties:
            _patch_run_properties(run, new_properties)

    return ET.tostring(root, encoding="utf-8", xml_declaration=True)


def _try_fragment_patch(document_xml: bytes, edits: list[dict[str, Any]]) -> bytes | None:
    if not edits or not all(_can_fragment_patch_edit(edit) for edit in edits):
        return None
    try:
        xml_text = document_xml.decode("utf-8")
    except UnicodeDecodeError:
        return None

    sorted_edits = sorted(edits, key=_edit_sort_key, reverse=True)
    current = xml_text
    for edit in sorted_edits:
        patched = _patch_single_fragment(current, edit)
        if patched is None:
            return None
        current = patched
    return current.encode("utf-8")


def _can_fragment_patch_edit(edit: dict[str, Any]) -> bool:
    operation = edit.get("operation")
    if operation == "patch-run":
        return ("newText" in edit or bool(edit.get("newProperties"))) and "locator" in edit
    if operation == "patch-paragraph":
        return "locator" in edit and bool(edit.get("newProperties"))
    if operation in {"patch-drawing-alt", "patch-drawing-properties"}:
        return "locator" in edit and bool(edit.get("newProperties"))
    return False


PARAGRAPH_FRAGMENT_RE = re.compile(r"<(?P<prefix>[A-Za-z0-9]+):p\b[^>]*>.*?</(?P=prefix):p>", re.DOTALL)
PARAGRAPH_START_RE = re.compile(r"<(?P<prefix>[A-Za-z0-9]+):p\b[^>]*>", re.DOTALL)
PPR_FRAGMENT_RE = re.compile(r"<(?P<prefix>[A-Za-z0-9]+):pPr\b[^>]*>.*?</(?P=prefix):pPr>", re.DOTALL)
RUN_FRAGMENT_RE = re.compile(r"<(?P<prefix>[A-Za-z0-9]+):r\b[^>]*>.*?</(?P=prefix):r>", re.DOTALL)
TEXT_FRAGMENT_RE = re.compile(r"<(?P<prefix>[A-Za-z0-9]+):t(?P<attrs>[^>]*)>(?P<text>.*?)</(?P=prefix):t>", re.DOTALL)
RPR_FRAGMENT_RE = re.compile(r"<(?P<prefix>[A-Za-z0-9]+):rPr\b[^>]*>.*?</(?P=prefix):rPr>", re.DOTALL)
DOCPR_FRAGMENT_RE = re.compile(r"<(?P<prefix>[A-Za-z0-9]+):docPr\b(?P<attrs>[^>]*?)(?P<close>/?>)", re.DOTALL)
EXTENT_FRAGMENT_RE = re.compile(r"<(?P<prefix>[A-Za-z0-9]+):extent\b(?P<attrs>[^>]*?)(?P<close>/?>)", re.DOTALL)
XML_ATTR_RE_TEMPLATE = r"(?P<lead>\s{attr}=)(?P<quote>['\"])(?P<value>.*?)(?P=quote)"


def _patch_single_fragment(xml_text: str, edit: dict[str, Any]) -> str | None:
    locator = edit["locator"]
    paragraph_index = locator["paragraphIndex"]
    paragraphs = list(PARAGRAPH_FRAGMENT_RE.finditer(xml_text))
    if paragraph_index < 1 or paragraph_index > len(paragraphs):
        return None
    paragraph_match = paragraphs[paragraph_index - 1]
    paragraph_xml = paragraph_match.group(0)
    if edit.get("operation") == "patch-paragraph":
        patched_paragraph_xml = _patch_paragraph_properties_fragment(paragraph_xml, edit.get("newProperties", {}))
        if patched_paragraph_xml is None:
            return None
        return xml_text[: paragraph_match.start()] + patched_paragraph_xml + xml_text[paragraph_match.end() :]

    run_index = locator["runInParagraphIndex"]
    runs = list(RUN_FRAGMENT_RE.finditer(paragraph_xml))
    if run_index < 1 or run_index > len(runs):
        return None
    run_match = runs[run_index - 1]
    run_xml = run_match.group(0)
    if edit.get("operation") in {"patch-drawing-alt", "patch-drawing-properties"}:
        patched_run_xml = _patch_drawing_fragment(run_xml, edit)
        if patched_run_xml is None:
            return None
        patched_paragraph_xml = (
            paragraph_xml[: run_match.start()]
            + patched_run_xml
            + paragraph_xml[run_match.end() :]
        )
        return xml_text[: paragraph_match.start()] + patched_paragraph_xml + xml_text[paragraph_match.end() :]

    patched_run_xml = run_xml
    if "newText" in edit:
        patched_run_xml = _patch_text_in_run_fragment(patched_run_xml, edit)
        if patched_run_xml is None:
            return None
    if edit.get("newProperties"):
        patched_run_xml = _patch_run_properties_fragment(patched_run_xml, edit["newProperties"])
        if patched_run_xml is None:
            return None
    patched_paragraph_xml = (
        paragraph_xml[: run_match.start()]
        + patched_run_xml
        + paragraph_xml[run_match.end() :]
    )
    return xml_text[: paragraph_match.start()] + patched_paragraph_xml + xml_text[paragraph_match.end() :]


def _patch_text_in_run_fragment(run_xml: str, edit: dict[str, Any]) -> str | None:
    text_matches = list(TEXT_FRAGMENT_RE.finditer(run_xml))
    if len(text_matches) != 1:
        return None
    text_match = text_matches[0]
    old_text = edit.get("oldText", "")
    existing_text = html.unescape(text_match.group("text"))
    if existing_text != old_text:
        return None
    new_text = edit.get("newText", "")
    attrs = text_match.group("attrs")
    has_preserve_space = "xml:space" in attrs and "preserve" in attrs
    if _needs_preserve_space(old_text) != _needs_preserve_space(new_text):
        return None
    if _needs_preserve_space(new_text) and not has_preserve_space:
        return None
    escaped_new_text = html.escape(new_text, quote=False)
    patched_run_xml = (
        run_xml[: text_match.start("text")]
        + escaped_new_text
        + run_xml[text_match.end("text") :]
    )
    return patched_run_xml


def _patch_run_properties_fragment(run_xml: str, properties: dict[str, str | None]) -> str | None:
    run_start = re.match(r"<(?P<prefix>[A-Za-z0-9]+):r\b[^>]*>", run_xml, flags=re.DOTALL)
    if run_start is None:
        return None
    prefix = run_start.group("prefix")
    rpr_match = RPR_FRAGMENT_RE.search(run_xml)
    if rpr_match is None:
        rpr_xml = f"<{prefix}:rPr></{prefix}:rPr>"
        run_xml = run_xml[: run_start.end()] + rpr_xml + run_xml[run_start.end() :]
        rpr_match = RPR_FRAGMENT_RE.search(run_xml)
        if rpr_match is None:
            return None
    rpr_xml = rpr_match.group(0)
    patched_rpr = rpr_xml
    for name, value in properties.items():
        if name == "bold":
            patched_rpr = _patch_rpr_empty_property(patched_rpr, prefix, "b", value)
        elif name == "italic":
            patched_rpr = _patch_rpr_empty_property(patched_rpr, prefix, "i", value)
        elif name == "font-size":
            patched_rpr = _patch_rpr_sized_property(patched_rpr, prefix, value)
        elif name == "color":
            patched_rpr = _patch_rpr_color_property(patched_rpr, prefix, value)
        else:
            return None
        if patched_rpr is None:
            return None
    return run_xml[: rpr_match.start()] + patched_rpr + run_xml[rpr_match.end() :]


def _patch_paragraph_properties_fragment(paragraph_xml: str, properties: dict[str, str | None]) -> str | None:
    paragraph_start = PARAGRAPH_START_RE.match(paragraph_xml)
    if paragraph_start is None:
        return None
    prefix = paragraph_start.group("prefix")
    ppr_match = PPR_FRAGMENT_RE.search(paragraph_xml)
    if ppr_match is None:
        ppr_xml = f"<{prefix}:pPr></{prefix}:pPr>"
        paragraph_xml = paragraph_xml[: paragraph_start.end()] + ppr_xml + paragraph_xml[paragraph_start.end() :]
        ppr_match = PPR_FRAGMENT_RE.search(paragraph_xml)
        if ppr_match is None:
            return None
    ppr_xml = ppr_match.group(0)
    patched_ppr = ppr_xml
    for name, value in properties.items():
        if name == "align":
            patched_ppr = _patch_ppr_alignment_fragment(patched_ppr, prefix, value)
        elif name == "first-line-indent":
            patched_ppr = _patch_ppr_first_line_fragment(patched_ppr, prefix, value)
        elif name == "line-spacing":
            patched_ppr = _patch_ppr_line_spacing_fragment(patched_ppr, prefix, value)
        else:
            return None
        if patched_ppr is None:
            return None
    return paragraph_xml[: ppr_match.start()] + patched_ppr + paragraph_xml[ppr_match.end() :]


def _patch_ppr_alignment_fragment(ppr_xml: str, prefix: str, value: str | None) -> str | None:
    if value is None:
        return _remove_child_element(ppr_xml, prefix, "jc")
    word_value = "both" if value == "justify" else value
    return _upsert_child_element_in(ppr_xml, prefix, "pPr", "jc", f'<{prefix}:jc {prefix}:val="{word_value}"/>')


def _patch_ppr_first_line_fragment(ppr_xml: str, prefix: str, value: str | None) -> str | None:
    attrs: dict[str, str | None] = {"firstLine": None, "firstLineChars": None}
    if value is not None:
        if value.endswith("char"):
            attrs["firstLineChars"] = str(round(float(value[:-4]) * 100))
        elif value.endswith("pt"):
            attrs["firstLine"] = str(round(float(value[:-2]) * 20))
        else:
            return None
    return _patch_child_start_attrs(ppr_xml, prefix, "pPr", "ind", attrs)


def _patch_ppr_line_spacing_fragment(ppr_xml: str, prefix: str, value: str | None) -> str | None:
    attrs: dict[str, str | None] = {"line": None, "lineRule": None}
    if value is not None:
        if value.endswith("pt"):
            attrs["line"] = str(round(float(value[:-2]) * 20))
            attrs["lineRule"] = "exact"
        else:
            attrs["line"] = str(round(float(value) * 240))
            attrs["lineRule"] = "auto"
    return _patch_child_start_attrs(ppr_xml, prefix, "pPr", "spacing", attrs)


def _patch_rpr_empty_property(rpr_xml: str, prefix: str, local_name: str, value: str | None) -> str | None:
    if value is None:
        return _remove_child_element(rpr_xml, prefix, local_name)
    if value == "true":
        element = f"<{prefix}:{local_name}/>"
    elif value == "false":
        element = f'<{prefix}:{local_name} {prefix}:val="0"/>'
    else:
        return None
    return _upsert_child_element(rpr_xml, prefix, local_name, element)


def _patch_rpr_sized_property(rpr_xml: str, prefix: str, value: str | None) -> str | None:
    if value is None:
        return _remove_child_element(rpr_xml, prefix, "sz")
    if not value.endswith("pt"):
        return None
    half_points = round(float(value[:-2]) * 2)
    if half_points <= 0:
        return None
    return _upsert_child_element(rpr_xml, prefix, "sz", f'<{prefix}:sz {prefix}:val="{half_points}"/>')


def _patch_rpr_color_property(rpr_xml: str, prefix: str, value: str | None) -> str | None:
    if value is None:
        return _remove_child_element(rpr_xml, prefix, "color")
    normalized = value[1:] if value.startswith("#") else value
    if len(normalized) != 6 or any(ch not in "0123456789abcdefABCDEF" for ch in normalized):
        return None
    return _upsert_child_element(rpr_xml, prefix, "color", f'<{prefix}:color {prefix}:val="{normalized.upper()}"/>')


def _child_element_pattern(prefix: str, local_name: str) -> re.Pattern[str]:
    return re.compile(
        rf"<{re.escape(prefix)}:{re.escape(local_name)}\b[^>]*/>|"
        rf"<{re.escape(prefix)}:{re.escape(local_name)}\b[^>]*>.*?</{re.escape(prefix)}:{re.escape(local_name)}>",
        flags=re.DOTALL,
    )


def _remove_child_element(rpr_xml: str, prefix: str, local_name: str) -> str:
    return _child_element_pattern(prefix, local_name).sub("", rpr_xml, count=1)


def _upsert_child_element(rpr_xml: str, prefix: str, local_name: str, element: str) -> str | None:
    return _upsert_child_element_in(rpr_xml, prefix, "rPr", local_name, element)


def _upsert_child_element_in(parent_xml: str, prefix: str, parent_local_name: str, child_local_name: str, element: str) -> str | None:
    pattern = _child_element_pattern(prefix, child_local_name)
    if pattern.search(parent_xml):
        return pattern.sub(element, parent_xml, count=1)
    close_tag = f"</{prefix}:{parent_local_name}>"
    index = parent_xml.rfind(close_tag)
    if index < 0:
        return None
    return parent_xml[:index] + element + parent_xml[index:]


def _patch_child_start_attrs(
    parent_xml: str,
    prefix: str,
    parent_local_name: str,
    child_local_name: str,
    attrs: dict[str, str | None],
) -> str | None:
    pattern = _child_element_pattern(prefix, child_local_name)
    match = pattern.search(parent_xml)
    if match is None:
        if not any(value is not None for value in attrs.values()):
            return parent_xml
        attr_text = " ".join(
            f'{prefix}:{name}="{html.escape(value, quote=True)}"'
            for name, value in attrs.items()
            if value is not None
        )
        element = f"<{prefix}:{child_local_name} {attr_text}/>"
        return _upsert_child_element_in(parent_xml, prefix, parent_local_name, child_local_name, element)

    child_xml = match.group(0)
    start_re = re.compile(
        rf"<{re.escape(prefix)}:{re.escape(child_local_name)}\b(?P<attrs>[^>]*?)(?P<close>/?>)",
        flags=re.DOTALL,
    )
    start_match = start_re.match(child_xml)
    if start_match is None:
        return None
    patched_start = start_match.group(0)
    for name, value in attrs.items():
        patched_start = _patch_start_tag_attr_lenient(patched_start, f"{prefix}:{name}", value)
        if patched_start is None:
            return None
    patched_child = patched_start + child_xml[start_match.end() :]
    if _is_empty_child_element(patched_child, prefix, child_local_name):
        return parent_xml[: match.start()] + parent_xml[match.end() :]
    return parent_xml[: match.start()] + patched_child + parent_xml[match.end() :]


def _patch_start_tag_attr_lenient(tag_text: str, attr_name: str, new_value: str | None) -> str | None:
    attr_re = re.compile(XML_ATTR_RE_TEMPLATE.format(attr=re.escape(attr_name)), re.DOTALL)
    if new_value is None and attr_re.search(tag_text) is None:
        return tag_text
    return _patch_start_tag_attr(tag_text, attr_name, None, new_value)


def _is_empty_child_element(child_xml: str, prefix: str, local_name: str) -> bool:
    return bool(
        re.fullmatch(rf"<{re.escape(prefix)}:{re.escape(local_name)}\s*/>", child_xml, flags=re.DOTALL)
        or re.fullmatch(
            rf"<{re.escape(prefix)}:{re.escape(local_name)}\s*></{re.escape(prefix)}:{re.escape(local_name)}>",
            child_xml,
            flags=re.DOTALL,
        )
    )


def _patch_drawing_fragment(run_xml: str, edit: dict[str, Any]) -> str | None:
    old_properties = edit.get("oldProperties", {})
    new_properties = edit.get("newProperties", {})
    patched = run_xml
    if "alt" in new_properties:
        patched = _patch_first_element_attr(
            patched,
            DOCPR_FRAGMENT_RE,
            "descr",
            old_properties.get("alt"),
            new_properties.get("alt"),
        )
        if patched is None:
            return None
    if "width-emu" in new_properties:
        patched = _patch_first_element_attr(
            patched,
            EXTENT_FRAGMENT_RE,
            "cx",
            old_properties.get("width-emu"),
            new_properties.get("width-emu"),
        )
        if patched is None:
            return None
    if "height-emu" in new_properties:
        patched = _patch_first_element_attr(
            patched,
            EXTENT_FRAGMENT_RE,
            "cy",
            old_properties.get("height-emu"),
            new_properties.get("height-emu"),
        )
        if patched is None:
            return None
    return patched


def _patch_first_element_attr(
    xml_text: str,
    element_re: re.Pattern[str],
    attr_name: str,
    old_value: str | None,
    new_value: str | None,
) -> str | None:
    match = element_re.search(xml_text)
    if match is None:
        return None
    tag_text = match.group(0)
    patched_tag = _patch_start_tag_attr(tag_text, attr_name, old_value, new_value)
    if patched_tag is None:
        return None
    return xml_text[: match.start()] + patched_tag + xml_text[match.end() :]


def _patch_start_tag_attr(
    tag_text: str,
    attr_name: str,
    old_value: str | None,
    new_value: str | None,
) -> str | None:
    attr_re = re.compile(XML_ATTR_RE_TEMPLATE.format(attr=re.escape(attr_name)), re.DOTALL)
    match = attr_re.search(tag_text)
    if match is None:
        if old_value is not None or new_value is None:
            return None
        insert_at = _start_tag_attr_insert_position(tag_text)
        if insert_at is None:
            return None
        return tag_text[:insert_at] + f' {attr_name}="{html.escape(new_value, quote=True)}"' + tag_text[insert_at:]

    actual_old = html.unescape(match.group("value"))
    if old_value is not None and actual_old != old_value:
        return None
    if new_value is None:
        remove_start = match.start("lead")
        while remove_start > 0 and tag_text[remove_start - 1].isspace():
            remove_start -= 1
        return tag_text[:remove_start] + tag_text[match.end() :]
    escaped_new = html.escape(new_value, quote=True)
    return tag_text[: match.start("value")] + escaped_new + tag_text[match.end("value") :]


def _start_tag_attr_insert_position(tag_text: str) -> int | None:
    stripped = tag_text.rstrip()
    if stripped.endswith("/>"):
        return tag_text.rfind("/>")
    if stripped.endswith(">"):
        return tag_text.rfind(">")
    return None


def _needs_preserve_space(text: str) -> bool:
    return bool(text) and (text[0].isspace() or text[-1].isspace() or "  " in text)


def _edit_sort_key(edit: dict[str, Any]) -> tuple[int, int]:
    if "locator" not in edit:
        return (0, 0)
    locator = edit["locator"]
    if "paragraphIndex" in locator:
        return (locator["paragraphIndex"], locator.get("runInParagraphIndex", 0))
    if "tableIndex" in locator:
        return (locator["tableIndex"], locator.get("rowInTableIndex", 0))
    return (0, 0)


def _patch_paragraph_properties(paragraph: ET.Element, properties: dict[str, str | None]) -> None:
    p_pr = paragraph.find("w:pPr", NS)
    if p_pr is None:
        p_pr = ET.Element(_tag("pPr"))
        paragraph.insert(0, p_pr)
    for name, value in properties.items():
        if name == "align":
            _set_alignment(p_pr, value)
        elif name == "first-line-indent":
            _set_first_line_indent(p_pr, value)
        elif name == "line-spacing":
            _set_line_spacing(p_pr, value)
        else:
            raise HDocxError(
                "PATCH_UNSUPPORTED_PARAGRAPH_PROPERTY",
                "Unsupported paragraph property reached patcher.",
                {"property": name},
            )


def _patch_paragraph_style(paragraph: ET.Element, style_id: str) -> None:
    p_pr = paragraph.find("w:pPr", NS)
    if p_pr is None:
        p_pr = ET.Element(_tag("pPr"))
        paragraph.insert(0, p_pr)
    p_style = p_pr.find("w:pStyle", NS)
    if p_style is None:
        p_style = ET.Element(_tag("pStyle"))
        p_pr.insert(0, p_style)
    p_style.attrib[_tag("val")] = style_id


def _patch_paragraph_numbering(paragraph: ET.Element, num_id: str, ilvl: str) -> None:
    p_pr = paragraph.find("w:pPr", NS)
    if p_pr is None:
        p_pr = ET.Element(_tag("pPr"))
        paragraph.insert(0, p_pr)
    num_pr = p_pr.find("w:numPr", NS)
    if num_pr is None:
        num_pr = ET.SubElement(p_pr, _tag("numPr"))
    ilvl_el = num_pr.find("w:ilvl", NS)
    if ilvl_el is None:
        ilvl_el = ET.Element(_tag("ilvl"))
        num_pr.insert(0, ilvl_el)
    ilvl_el.attrib[_tag("val")] = ilvl
    num_id_el = num_pr.find("w:numId", NS)
    if num_id_el is None:
        num_id_el = ET.SubElement(num_pr, _tag("numId"))
    num_id_el.attrib[_tag("val")] = num_id


def _set_alignment(p_pr: ET.Element, value: str | None) -> None:
    existing = p_pr.find("w:jc", NS)
    if value is None:
        if existing is not None:
            p_pr.remove(existing)
        return
    if existing is None:
        existing = ET.SubElement(p_pr, _tag("jc"))
    existing.attrib[_tag("val")] = "both" if value == "justify" else value


def _set_first_line_indent(p_pr: ET.Element, value: str | None) -> None:
    existing = p_pr.find("w:ind", NS)
    if value is None:
        if existing is not None:
            existing.attrib.pop(_tag("firstLine"), None)
            existing.attrib.pop(_tag("firstLineChars"), None)
            if not existing.attrib:
                p_pr.remove(existing)
        return
    if existing is None:
        existing = ET.SubElement(p_pr, _tag("ind"))
    existing.attrib.pop(_tag("firstLine"), None)
    existing.attrib.pop(_tag("firstLineChars"), None)
    if value.endswith("char"):
        chars = float(value[:-4])
        existing.attrib[_tag("firstLineChars")] = str(round(chars * 100))
    elif value.endswith("pt"):
        points = float(value[:-2])
        existing.attrib[_tag("firstLine")] = str(round(points * 20))
    else:
        raise HDocxError("PATCH_INVALID_FIRST_LINE_INDENT", "Indent must use char or pt units.", {"value": value})


def _set_line_spacing(p_pr: ET.Element, value: str | None) -> None:
    existing = p_pr.find("w:spacing", NS)
    if value is None:
        if existing is not None:
            existing.attrib.pop(_tag("line"), None)
            existing.attrib.pop(_tag("lineRule"), None)
            if not existing.attrib:
                p_pr.remove(existing)
        return
    if existing is None:
        existing = ET.SubElement(p_pr, _tag("spacing"))
    if value.endswith("pt"):
        points = float(value[:-2])
        existing.attrib[_tag("line")] = str(round(points * 20))
        existing.attrib[_tag("lineRule")] = "exact"
    else:
        multiple = float(value)
        existing.attrib[_tag("line")] = str(round(multiple * 240))
        existing.attrib[_tag("lineRule")] = "auto"


def _patch_run_properties(run: ET.Element, properties: dict[str, str | None]) -> None:
    r_pr = run.find("w:rPr", NS)
    if r_pr is None:
        r_pr = ET.Element(_tag("rPr"))
        run.insert(0, r_pr)
    _patch_rpr_properties(r_pr, properties)


def _patch_rpr_properties(r_pr: ET.Element, properties: dict[str, str | None]) -> None:

    for name, value in properties.items():
        if name == "bold":
            _set_on_off(r_pr, "b", value)
        elif name == "italic":
            _set_on_off(r_pr, "i", value)
        elif name == "font-size":
            _set_font_size(r_pr, value)
        elif name == "color":
            _set_color(r_pr, value)
        else:
            raise HDocxError(
                "PATCH_UNSUPPORTED_RUN_PROPERTY",
                "Unsupported run property reached patcher.",
                {"property": name},
            )


def _set_on_off(r_pr: ET.Element, local_name: str, value: str | None) -> None:
    existing = r_pr.find(f"w:{local_name}", NS)
    if value is None:
        if existing is not None:
            r_pr.remove(existing)
        return
    if existing is None:
        existing = ET.SubElement(r_pr, _tag(local_name))
    if value == "true":
        existing.attrib.pop(_tag("val"), None)
    elif value == "false":
        existing.attrib[_tag("val")] = "0"
    else:
        raise HDocxError(
            "PATCH_INVALID_ON_OFF_VALUE",
            "On/off run property must be true or false.",
            {"property": local_name, "value": value},
        )


def _set_font_size(r_pr: ET.Element, value: str | None) -> None:
    existing = r_pr.find("w:sz", NS)
    if value is None:
        if existing is not None:
            r_pr.remove(existing)
        return
    if not value.endswith("pt"):
        raise HDocxError("PATCH_INVALID_FONT_SIZE", "Font size must use pt units.", {"value": value})
    try:
        points = float(value[:-2])
    except ValueError as exc:
        raise HDocxError("PATCH_INVALID_FONT_SIZE", "Font size must be numeric pt.", {"value": value}) from exc
    half_points = round(points * 2)
    if half_points <= 0:
        raise HDocxError("PATCH_INVALID_FONT_SIZE", "Font size must be positive.", {"value": value})
    if existing is None:
        existing = ET.SubElement(r_pr, _tag("sz"))
    existing.attrib[_tag("val")] = str(half_points)


def _set_color(r_pr: ET.Element, value: str | None) -> None:
    existing = r_pr.find("w:color", NS)
    if value is None:
        if existing is not None:
            r_pr.remove(existing)
        return
    normalized = value[1:] if value.startswith("#") else value
    if len(normalized) != 6 or any(ch not in "0123456789abcdefABCDEF" for ch in normalized):
        raise HDocxError("PATCH_INVALID_COLOR", "Color must be #RRGGBB.", {"value": value})
    if existing is None:
        existing = ET.SubElement(r_pr, _tag("color"))
    existing.attrib[_tag("val")] = normalized.upper()


def _replace_run_with_split_segments(
    paragraph: ET.Element,
    run: ET.Element,
    split_segments: list[dict[str, Any]],
    edit: dict[str, Any],
) -> None:
    text_nodes = [child for child in list(run) if child.tag == _tag("t")]
    if len(text_nodes) != 1:
        raise HDocxError(
            "PATCH_RUN_NOT_SIMPLE_TEXT",
            "Only runs with exactly one w:t text node can be split in this phase.",
            {"editId": edit["id"], "nodeId": edit["nodeId"]},
        )
    children = list(paragraph)
    try:
        run_position = children.index(run)
    except ValueError as exc:
        raise HDocxError(
            "PATCH_RUN_LOCATOR_NOT_FOUND",
            "Run no longer belongs to the located paragraph.",
            {"editId": edit["id"], "nodeId": edit["nodeId"]},
        ) from exc

    new_runs: list[ET.Element] = []
    for segment in split_segments:
        new_run = copy.deepcopy(run)
        new_text = segment["text"]
        new_text_nodes = [child for child in list(new_run) if child.tag == _tag("t")]
        new_text_nodes[0].text = new_text
        if _needs_preserve_space(new_text):
            new_text_nodes[0].attrib[f"{{{XML_NS}}}space"] = "preserve"
        elif f"{{{XML_NS}}}space" in new_text_nodes[0].attrib:
            del new_text_nodes[0].attrib[f"{{{XML_NS}}}space"]
        if segment.get("properties"):
            _patch_run_properties(new_run, segment["properties"])
        new_runs.append(new_run)

    paragraph.remove(run)
    for offset, new_run in enumerate(new_runs):
        paragraph.insert(run_position + offset, new_run)


def _patch_style(root: ET.Element, edit: dict[str, Any]) -> None:
    style_id = edit["styleId"]
    target_style = None
    for style in root.findall("w:style", NS):
        if style.attrib.get(_tag("styleId")) == style_id:
            target_style = style
            break
    if target_style is None:
        raise HDocxError(
            "PATCH_STYLE_NOT_FOUND",
            "Target styleId was not found in styles.xml.",
            {"styleId": style_id},
        )
    new_properties = edit.get("newProperties", {})
    paragraph_props = new_properties.get("paragraph") or {}
    run_props = new_properties.get("run") or {}
    if paragraph_props:
        p_pr = target_style.find("w:pPr", NS)
        if p_pr is None:
            p_pr = ET.Element(_tag("pPr"))
            target_style.append(p_pr)
        _patch_ppr_element(p_pr, paragraph_props)
    if run_props:
        r_pr = target_style.find("w:rPr", NS)
        if r_pr is None:
            r_pr = ET.Element(_tag("rPr"))
            target_style.append(r_pr)
        _patch_rpr_properties(r_pr, run_props)


def _create_style(root: ET.Element, edit: dict[str, Any]) -> None:
    style_id = edit["styleId"]
    for style in root.findall("w:style", NS):
        if style.attrib.get(_tag("styleId")) == style_id:
            raise HDocxError(
                "PATCH_STYLE_ALREADY_EXISTS",
                "Target styleId already exists in styles.xml.",
                {"styleId": style_id},
            )
    style = ET.SubElement(root, _tag("style"), {_tag("type"): edit.get("styleType", "paragraph"), _tag("styleId"): style_id})
    ET.SubElement(style, _tag("name"), {_tag("val"): edit.get("name") or style_id})
    if edit.get("basedOn"):
        ET.SubElement(style, _tag("basedOn"), {_tag("val"): edit["basedOn"]})
    if edit.get("next"):
        ET.SubElement(style, _tag("next"), {_tag("val"): edit["next"]})
    if edit.get("qFormat", True):
        ET.SubElement(style, _tag("qFormat"))
    new_properties = edit.get("newProperties", {})
    paragraph_props = new_properties.get("paragraph") or {}
    run_props = new_properties.get("run") or {}
    if paragraph_props:
        p_pr = ET.SubElement(style, _tag("pPr"))
        _patch_ppr_element(p_pr, paragraph_props)
    if run_props:
        r_pr = ET.SubElement(style, _tag("rPr"))
        _patch_rpr_properties(r_pr, run_props)


def _delete_style(root: ET.Element, edit: dict[str, Any]) -> None:
    style_id = edit["styleId"]
    for style in root.findall("w:style", NS):
        if style.attrib.get(_tag("styleId")) == style_id:
            root.remove(style)
            return
    raise HDocxError(
        "PATCH_STYLE_NOT_FOUND",
        "Target styleId was not found in styles.xml.",
        {"styleId": style_id},
    )


def _patch_drawing_properties(run: ET.Element, edit: dict[str, Any]) -> None:
    new_properties = edit.get("newProperties", {})
    if "alt" in new_properties:
        doc_pr = run.find(".//wp:docPr", NS)
        if doc_pr is None:
            raise HDocxError(
                "PATCH_DRAWING_DOCPR_NOT_FOUND",
                "Drawing docPr was not found for alt text patch.",
                {"editId": edit["id"], "nodeId": edit["nodeId"]},
            )
        new_alt = new_properties.get("alt")
        if new_alt is None:
            doc_pr.attrib.pop("descr", None)
        else:
            doc_pr.attrib["descr"] = new_alt

    size_properties = {key: value for key, value in new_properties.items() if key in {"width-emu", "height-emu"}}
    if not size_properties:
        return
    extent = run.find(".//wp:extent", NS)
    if extent is None:
        raise HDocxError(
            "PATCH_DRAWING_EXTENT_NOT_FOUND",
            "Drawing wp:extent was not found for image size patch.",
            {"editId": edit["id"], "nodeId": edit["nodeId"]},
        )
    if "width-emu" in size_properties:
        extent.attrib["cx"] = _validate_emu_size("width-emu", size_properties["width-emu"], edit)
    if "height-emu" in size_properties:
        extent.attrib["cy"] = _validate_emu_size("height-emu", size_properties["height-emu"], edit)


def _validate_emu_size(prop_name: str, value: str | None, edit: dict[str, Any]) -> str:
    if value is None or not value.isdigit() or int(value) <= 0:
        raise HDocxError(
            "PATCH_INVALID_DRAWING_SIZE",
            "Drawing size must be a positive EMU integer.",
            {"editId": edit["id"], "nodeId": edit["nodeId"], "property": prop_name, "value": value},
        )
    return str(int(value))


def _patch_ppr_element(p_pr: ET.Element, properties: dict[str, str | None]) -> None:
    for name, value in properties.items():
        if name == "align":
            _set_alignment(p_pr, value)
        elif name == "first-line-indent":
            _set_first_line_indent(p_pr, value)
        elif name == "line-spacing":
            _set_line_spacing(p_pr, value)
        else:
            raise HDocxError(
                "PATCH_UNSUPPORTED_PARAGRAPH_PROPERTY",
                "Unsupported paragraph property reached patcher.",
                {"property": name},
            )


def _patch_content_types(root: ET.Element, edit: dict[str, Any]) -> None:
    part_name = edit.get("partName")
    content_type = edit["contentType"]
    if part_name:
        for override in root.findall(f"{{{CT_NS}}}Override"):
            if override.attrib.get("PartName") == part_name:
                if override.attrib.get("ContentType") != content_type:
                    raise HDocxError(
                        "PATCH_CONTENT_TYPE_CONFLICT",
                        "Existing content type override conflicts with requested part type.",
                        {"partName": part_name, "existing": override.attrib.get("ContentType"), "requested": content_type},
                    )
                return
        ET.SubElement(root, f"{{{CT_NS}}}Override", {"PartName": part_name, "ContentType": content_type})
        return

    extension = edit["extension"]
    for default in root.findall(f"{{{CT_NS}}}Default"):
        if default.attrib.get("Extension", "").lower() == extension.lower():
            if default.attrib.get("ContentType") != content_type:
                raise HDocxError(
                    "PATCH_CONTENT_TYPE_CONFLICT",
                    "Existing content type default conflicts with requested image type.",
                    {"extension": extension, "existing": default.attrib.get("ContentType"), "requested": content_type},
                )
            return
    ET.SubElement(root, f"{{{CT_NS}}}Default", {"Extension": extension, "ContentType": content_type})


def _patch_relationships(root: ET.Element, edit: dict[str, Any]) -> None:
    rel_id = edit["relationshipId"]
    relationship_type = edit.get("relationshipType", IMAGE_REL_TYPE)
    for rel in root.findall(f"{{{PKG_REL_NS}}}Relationship"):
        if rel.attrib.get("Id") == rel_id:
            raise HDocxError(
                "PATCH_RELATIONSHIP_ID_CONFLICT",
                "Relationship id already exists.",
                {"relationshipId": rel_id},
            )
    ET.SubElement(
        root,
        f"{{{PKG_REL_NS}}}Relationship",
        {"Id": rel_id, "Type": relationship_type, "Target": edit["target"]},
    )


def _insert_image_around_paragraph(root: ET.Element, paragraph: ET.Element, edit: dict[str, Any], position: str) -> None:
    parent = _find_parent(root, paragraph)
    if parent is None:
        raise HDocxError(
            "PATCH_PARAGRAPH_PARENT_NOT_FOUND",
            "Target paragraph parent was not found for image insertion.",
            {"editId": edit["id"], "nodeId": edit["nodeId"]},
        )
    children = list(parent)
    try:
        index = children.index(paragraph)
    except ValueError as exc:
        raise HDocxError(
            "PATCH_PARAGRAPH_PARENT_NOT_FOUND",
            "Target paragraph no longer belongs to its parent.",
            {"editId": edit["id"], "nodeId": edit["nodeId"]},
        ) from exc
    insert_index = index if position == "before" else index + 1
    parent.insert(insert_index, _image_paragraph(edit))


def _find_parent(root: ET.Element, target: ET.Element) -> ET.Element | None:
    for parent in root.iter():
        if target in list(parent):
            return parent
    return None


def _image_paragraph(edit: dict[str, Any]) -> ET.Element:
    paragraph = ET.Element(_tag("p"))
    run = ET.SubElement(paragraph, _tag("r"))
    drawing = ET.SubElement(run, _tag("drawing"))
    inline = ET.SubElement(drawing, f"{{{WP_NS}}}inline")
    ET.SubElement(inline, f"{{{WP_NS}}}extent", {"cx": edit["widthEmu"], "cy": edit["heightEmu"]})
    doc_pr_attrs = {"id": str(edit["docPrId"]), "name": f"Picture {edit['docPrId']}"}
    if edit.get("alt"):
        doc_pr_attrs["descr"] = edit["alt"]
    ET.SubElement(inline, f"{{{WP_NS}}}docPr", doc_pr_attrs)
    c_nv = ET.SubElement(inline, f"{{{WP_NS}}}cNvGraphicFramePr")
    ET.SubElement(c_nv, f"{{{A_NS}}}graphicFrameLocks", {"noChangeAspect": "1"})
    graphic = ET.SubElement(inline, f"{{{A_NS}}}graphic")
    graphic_data = ET.SubElement(
        graphic,
        f"{{{A_NS}}}graphicData",
        {"uri": "http://schemas.openxmlformats.org/drawingml/2006/picture"},
    )
    pic = ET.SubElement(graphic_data, f"{{{PIC_NS}}}pic")
    nv_pic_pr = ET.SubElement(pic, f"{{{PIC_NS}}}nvPicPr")
    ET.SubElement(nv_pic_pr, f"{{{PIC_NS}}}cNvPr", {"id": "0", "name": edit["mediaFileName"]})
    ET.SubElement(nv_pic_pr, f"{{{PIC_NS}}}cNvPicPr")
    blip_fill = ET.SubElement(pic, f"{{{PIC_NS}}}blipFill")
    ET.SubElement(blip_fill, f"{{{A_NS}}}blip", {f"{{{R_NS}}}embed": edit["relationshipId"]})
    stretch = ET.SubElement(blip_fill, f"{{{A_NS}}}stretch")
    ET.SubElement(stretch, f"{{{A_NS}}}fillRect")
    sp_pr = ET.SubElement(pic, f"{{{PIC_NS}}}spPr")
    xfrm = ET.SubElement(sp_pr, f"{{{A_NS}}}xfrm")
    ET.SubElement(xfrm, f"{{{A_NS}}}off", {"x": "0", "y": "0"})
    ET.SubElement(xfrm, f"{{{A_NS}}}ext", {"cx": edit["widthEmu"], "cy": edit["heightEmu"]})
    prst_geom = ET.SubElement(sp_pr, f"{{{A_NS}}}prstGeom", {"prst": "rect"})
    ET.SubElement(prst_geom, f"{{{A_NS}}}avLst")
    return paragraph


def _patch_numbering_level(root: ET.Element, edit: dict[str, Any]) -> None:
    abstract_num_id = edit["abstractNumId"]
    ilvl = edit["ilvl"]
    abstract_num = None
    for candidate in root.findall("w:abstractNum", NS):
        if candidate.attrib.get(_tag("abstractNumId")) == abstract_num_id:
            abstract_num = candidate
            break
    if abstract_num is None:
        raise HDocxError(
            "PATCH_NUMBERING_ABSTRACT_NOT_FOUND",
            "Target abstract numbering definition was not found.",
            {"abstractNumId": abstract_num_id},
        )
    level = None
    for candidate in abstract_num.findall("w:lvl", NS):
        if candidate.attrib.get(_tag("ilvl")) == ilvl:
            level = candidate
            break
    if level is None:
        level = ET.SubElement(abstract_num, _tag("lvl"), {_tag("ilvl"): ilvl})
    for name, value in edit.get("newProperties", {}).items():
        if name == "num-format":
            _set_level_value(level, "numFmt", value)
        elif name == "level-text":
            _set_level_value(level, "lvlText", value)
        elif name == "start":
            _set_level_value(level, "start", value)
        elif name == "number-suffix":
            _set_level_value(level, "suff", value)
        elif name in {"num-indent-left", "num-indent-hanging", "num-indent-first-line"}:
            _set_level_indent(level, name, value)
        else:
            raise HDocxError(
                "PATCH_UNSUPPORTED_NUMBERING_PROPERTY",
                "Unsupported numbering property reached patcher.",
                {"property": name},
            )


def _set_level_value(level: ET.Element, child_name: str, value: str) -> None:
    child = level.find(f"w:{child_name}", NS)
    if child is None:
        child = ET.SubElement(level, _tag(child_name))
    child.attrib[_tag("val")] = value


def _set_level_indent(level: ET.Element, prop_name: str, value: str) -> None:
    p_pr = level.find("w:pPr", NS)
    if p_pr is None:
        p_pr = ET.SubElement(level, _tag("pPr"))
    ind = p_pr.find("w:ind", NS)
    if ind is None:
        ind = ET.SubElement(p_pr, _tag("ind"))
    attr_name = {
        "num-indent-left": "left",
        "num-indent-hanging": "hanging",
        "num-indent-first-line": "firstLine",
    }[prop_name]
    ind.attrib[_tag(attr_name)] = value


def _create_numbering_list(root: ET.Element, edit: dict[str, Any]) -> None:
    abstract_num_id = edit["abstractNumId"]
    num_id = edit["numId"]
    for abstract_num in root.findall("w:abstractNum", NS):
        if abstract_num.attrib.get(_tag("abstractNumId")) == abstract_num_id:
            raise HDocxError(
                "PATCH_NUMBERING_ABSTRACT_ALREADY_EXISTS",
                "abstractNumId already exists.",
                {"abstractNumId": abstract_num_id},
            )
    for num in root.findall("w:num", NS):
        if num.attrib.get(_tag("numId")) == num_id:
            raise HDocxError(
                "PATCH_NUMBERING_NUM_ALREADY_EXISTS",
                "numId already exists.",
                {"numId": num_id},
            )
    abstract_num = ET.SubElement(root, _tag("abstractNum"), {_tag("abstractNumId"): abstract_num_id})
    levels = edit.get("levels") or {edit.get("ilvl", "0"): edit.get("newProperties", {})}
    for ilvl, props in sorted(levels.items(), key=lambda item: int(item[0])):
        level = ET.SubElement(abstract_num, _tag("lvl"), {_tag("ilvl"): str(ilvl)})
        display_level = int(ilvl) + 1
        _set_level_value(level, "start", props.get("start", "1"))
        _set_level_value(level, "numFmt", props.get("num-format", "decimal"))
        _set_level_value(level, "lvlText", props.get("level-text", f"%{display_level}."))
        if props.get("number-suffix"):
            _set_level_value(level, "suff", props["number-suffix"])
        for prop_name in ("num-indent-left", "num-indent-hanging", "num-indent-first-line"):
            if props.get(prop_name) is not None:
                _set_level_indent(level, prop_name, props[prop_name])
    num = ET.SubElement(root, _tag("num"), {_tag("numId"): num_id})
    ET.SubElement(num, _tag("abstractNumId"), {_tag("val"): abstract_num_id})


def _patch_comment_text(root: ET.Element, edit: dict[str, Any]) -> None:
    comment_id = edit.get("locator", {}).get("commentId")
    if comment_id is None:
        raise HDocxError(
            "PATCH_COMMENT_LOCATOR_MISSING",
            "Comment locator is missing.",
            {"editId": edit["id"], "nodeId": edit["nodeId"]},
        )
    target = None
    for comment in root.findall("w:comment", NS):
        if comment.attrib.get(_tag("id")) == str(comment_id):
            target = comment
            break
    if target is None:
        raise HDocxError(
            "PATCH_COMMENT_NOT_FOUND",
            "Target comment was not found.",
            {"editId": edit["id"], "commentId": comment_id},
        )
    text_nodes = target.findall(".//w:t", NS)
    if len(text_nodes) != 1:
        raise HDocxError(
            "PATCH_COMMENT_TEXT_NOT_SIMPLE",
            "Only comments with exactly one w:t text node are supported in this phase.",
            {"editId": edit["id"], "commentId": comment_id},
        )
    text_nodes[0].text = edit["newText"]
    if _needs_preserve_space(edit["newText"]):
        text_nodes[0].attrib[f"{{{XML_NS}}}space"] = "preserve"
    else:
        text_nodes[0].attrib.pop(f"{{{XML_NS}}}space", None)


def _patch_revision_action(root: ET.Element, edit: dict[str, Any]) -> None:
    revision_kind = edit["revisionKind"]
    action = edit["action"]
    source_id = edit.get("locator", {}).get("sourceId")
    tag_name = {
        "revision-insert": "ins",
        "revision-delete": "del",
    }.get(revision_kind)
    if tag_name is None:
        raise HDocxError(
            "PATCH_REVISION_KIND_UNSUPPORTED",
            "Unsupported revision kind.",
            {"editId": edit["id"], "revisionKind": revision_kind},
        )
    target = None
    parent = None
    for candidate_parent in root.iter():
        for child in list(candidate_parent):
            if child.tag != _tag(tag_name):
                continue
            if source_id is None or child.attrib.get(_tag("id")) == str(source_id):
                target = child
                parent = candidate_parent
                break
        if target is not None:
            break
    if target is None or parent is None:
        raise HDocxError(
            "PATCH_REVISION_NOT_FOUND",
            "Target revision wrapper was not found.",
            {"editId": edit["id"], "revisionKind": revision_kind, "sourceId": source_id},
        )
    if (revision_kind == "revision-insert" and action == "accept") or (
        revision_kind == "revision-delete" and action == "reject"
    ):
        _unwrap_child(parent, target)
    else:
        parent.remove(target)


def _patch_equation_omml(paragraph: ET.Element, edit: dict[str, Any]) -> None:
    target_index = edit.get("locator", {}).get("protectedInParagraphIndex")
    if target_index is None:
        raise HDocxError(
            "PATCH_EQUATION_LOCATOR_MISSING",
            "Equation locator is missing.",
            {"editId": edit["id"], "nodeId": edit["nodeId"]},
        )
    current_count = 0
    target = None
    target_position = None
    children = list(paragraph)
    for position, child in enumerate(children):
        kind = _protected_inline_kind(child)
        if kind is None:
            continue
        current_count += 1
        if current_count == int(target_index):
            target = child
            target_position = position
            if kind != "equation":
                raise HDocxError(
                    "PATCH_EQUATION_KIND_MISMATCH",
                    "Equation locator resolved to a non-equation protected node.",
                    {"editId": edit["id"], "nodeId": edit["nodeId"], "actualKind": kind},
                )
            break
    if target is None or target_position is None:
        raise HDocxError(
            "PATCH_EQUATION_NOT_FOUND",
            "Target equation was not found.",
            {"editId": edit["id"], "nodeId": edit["nodeId"], "protectedInParagraphIndex": target_index},
        )
    expected_hash = edit.get("expectedOldHash")
    if expected_hash and sha256_bytes(ET.tostring(target, encoding="utf-8")) != expected_hash:
        raise HDocxError(
            "PATCH_EQUATION_HASH_MISMATCH",
            "Target equation no longer matches the manifest hash.",
            {"editId": edit["id"], "nodeId": edit["nodeId"]},
        )
    try:
        replacement = ET.fromstring(edit["newOmml"])
    except ET.ParseError as exc:
        raise HDocxError(
            "PATCH_EQUATION_OMML_PARSE_ERROR",
            "Replacement OMML could not be parsed.",
            {"editId": edit["id"], "nodeId": edit["nodeId"], "message": str(exc)},
        ) from exc
    if replacement.tag not in {f"{{{M_NS}}}oMath", f"{{{M_NS}}}oMathPara"}:
        raise HDocxError(
            "PATCH_EQUATION_OMML_INVALID_ROOT",
            "Replacement OMML root must be m:oMath or m:oMathPara.",
            {"editId": edit["id"], "nodeId": edit["nodeId"], "rootTag": replacement.tag},
        )
    paragraph.remove(target)
    paragraph.insert(target_position, replacement)


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


def _unwrap_child(parent: ET.Element, child: ET.Element) -> None:
    children = list(parent)
    index = children.index(child)
    parent.remove(child)
    for offset, grandchild in enumerate(list(child)):
        parent.insert(index + offset, grandchild)


def _insert_table_row_after(root: ET.Element, edit: dict[str, Any]) -> None:
    locator = edit["locator"]
    table_index = locator["tableIndex"]
    row_in_table_index = locator["rowInTableIndex"]
    tables = root.findall(".//w:tbl", NS)
    if table_index < 1 or table_index > len(tables):
        raise HDocxError(
            "PATCH_TABLE_LOCATOR_NOT_FOUND",
            "Table locator no longer resolves.",
            {"editId": edit["id"], "tableIndex": table_index},
        )
    table = tables[table_index - 1]
    rows = table.findall("w:tr", NS)
    if row_in_table_index < 1 or row_in_table_index > len(rows):
        raise HDocxError(
            "PATCH_TABLE_ROW_LOCATOR_NOT_FOUND",
            "Table row locator no longer resolves.",
            {"editId": edit["id"], "rowInTableIndex": row_in_table_index},
        )
    source_row = rows[row_in_table_index - 1]
    new_row = copy.deepcopy(source_row)
    cells = new_row.findall("w:tc", NS)
    cell_texts = edit.get("cellTexts", [])
    if len(cells) != len(cell_texts):
        raise HDocxError(
            "PATCH_TABLE_ROW_CELL_COUNT",
            "Inserted row cell count does not match cloned row.",
            {"editId": edit["id"], "expected": len(cells), "actual": len(cell_texts)},
        )
    for cell, text in zip(cells, cell_texts):
        _reset_cell_content(cell, text)
    table_children = list(table)
    try:
        insert_index = table_children.index(source_row) + 1
    except ValueError as exc:
        raise HDocxError(
            "PATCH_TABLE_ROW_LOCATOR_NOT_FOUND",
            "Target row no longer belongs to located table.",
            {"editId": edit["id"]},
        ) from exc
    table.insert(insert_index, new_row)


def _delete_table_row(root: ET.Element, edit: dict[str, Any]) -> None:
    locator = edit["locator"]
    table_index = locator["tableIndex"]
    row_in_table_index = locator["rowInTableIndex"]
    tables = root.findall(".//w:tbl", NS)
    if table_index < 1 or table_index > len(tables):
        raise HDocxError(
            "PATCH_TABLE_LOCATOR_NOT_FOUND",
            "Table locator no longer resolves.",
            {"editId": edit["id"], "tableIndex": table_index},
        )
    table = tables[table_index - 1]
    rows = table.findall("w:tr", NS)
    if row_in_table_index < 1 or row_in_table_index > len(rows):
        raise HDocxError(
            "PATCH_TABLE_ROW_LOCATOR_NOT_FOUND",
            "Table row locator no longer resolves.",
            {"editId": edit["id"], "rowInTableIndex": row_in_table_index},
        )
    if len(rows) <= 1:
        raise HDocxError(
            "PATCH_TABLE_ROW_DELETE_LAST_ROW",
            "Deleting the only row in a table is not supported.",
            {"editId": edit["id"]},
        )
    table.remove(rows[row_in_table_index - 1])


def _insert_table_column_after(root: ET.Element, edit: dict[str, Any]) -> None:
    table, _, cell_index = _locate_table_column(root, edit)
    rows = table.findall("w:tr", NS)
    cell_texts = edit.get("cellTexts", [])
    if len(rows) != len(cell_texts):
        raise HDocxError(
            "PATCH_TABLE_COLUMN_CELL_COUNT",
            "Inserted column cell count does not match table row count.",
            {"editId": edit["id"], "expected": len(rows), "actual": len(cell_texts)},
        )
    for row, text in zip(rows, cell_texts):
        cells = row.findall("w:tc", NS)
        if cell_index < 1 or cell_index > len(cells):
            raise HDocxError(
                "PATCH_TABLE_COLUMN_LOCATOR_NOT_FOUND",
                "Column locator no longer resolves in every row.",
                {"editId": edit["id"], "cellInRowIndex": cell_index},
            )
        new_cell = copy.deepcopy(cells[cell_index - 1])
        _reset_cell_content(new_cell, text)
        row_children = list(row)
        insert_index = row_children.index(cells[cell_index - 1]) + 1
        row.insert(insert_index, new_cell)


def _delete_table_column(root: ET.Element, edit: dict[str, Any]) -> None:
    table, _, cell_index = _locate_table_column(root, edit)
    rows = table.findall("w:tr", NS)
    for row in rows:
        cells = row.findall("w:tc", NS)
        if len(cells) <= 1:
            raise HDocxError(
                "PATCH_TABLE_COLUMN_DELETE_LAST_COLUMN",
                "Deleting the only column in a table is not supported.",
                {"editId": edit["id"]},
            )
        if cell_index < 1 or cell_index > len(cells):
            raise HDocxError(
                "PATCH_TABLE_COLUMN_LOCATOR_NOT_FOUND",
                "Column locator no longer resolves in every row.",
                {"editId": edit["id"], "cellInRowIndex": cell_index},
            )
        row.remove(cells[cell_index - 1])


def _locate_table_column(root: ET.Element, edit: dict[str, Any]) -> tuple[ET.Element, ET.Element, int]:
    locator = edit["locator"]
    table_index = locator["tableIndex"]
    row_in_table_index = locator["rowInTableIndex"]
    cell_index = locator["cellInRowIndex"]
    tables = root.findall(".//w:tbl", NS)
    if table_index < 1 or table_index > len(tables):
        raise HDocxError(
            "PATCH_TABLE_LOCATOR_NOT_FOUND",
            "Table locator no longer resolves.",
            {"editId": edit["id"], "tableIndex": table_index},
        )
    table = tables[table_index - 1]
    rows = table.findall("w:tr", NS)
    if row_in_table_index < 1 or row_in_table_index > len(rows):
        raise HDocxError(
            "PATCH_TABLE_ROW_LOCATOR_NOT_FOUND",
            "Table row locator no longer resolves.",
            {"editId": edit["id"], "rowInTableIndex": row_in_table_index},
        )
    row = rows[row_in_table_index - 1]
    return table, row, cell_index


def _reset_cell_content(cell: ET.Element, text: str) -> None:
    tc_pr = cell.find("w:tcPr", NS)
    for child in list(cell):
        if child is not tc_pr:
            cell.remove(child)
    paragraph = ET.Element(_tag("p"))
    run = ET.SubElement(paragraph, _tag("r"))
    text_node = ET.SubElement(run, _tag("t"))
    text_node.text = text
    if _needs_preserve_space(text):
        text_node.attrib[f"{{{XML_NS}}}space"] = "preserve"
    cell.append(paragraph)
