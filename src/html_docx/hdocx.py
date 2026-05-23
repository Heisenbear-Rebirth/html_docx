from __future__ import annotations

from datetime import datetime, timezone
from pathlib import Path
from typing import Any
import copy
import importlib.metadata
import platform
import zipfile
import re
import sys
import xml.etree.ElementTree as ET

from . import __version__
from .creator import create_canonical_docx
from .errors import HDocxError
from .html_scan import collect_hdocx_ids, collect_hdocx_nodes
from .package import (
    compare_docx_entries,
    extract_entries,
    read_docx_entries,
    repack_docx_with_modified_entries,
)
from .patcher import patch_document_run_text
from .projector import (
    M_NS,
    NS,
    W_NS,
    _tag,
    build_document_html,
    load_numbering_definitions,
    load_style_definitions,
    project_docx_entries,
    project_main_document,
)
from .utils import copy_file, ensure_new_dir, read_json, sha256_bytes, sha256_file, write_json


MANIFEST_NAME = "manifest.json"
HTML_NAME = "document.html"
PREVIEW_CSS_NAME = "styles.generated.css"
AGENT_HCSS_NAME = "agent.edits.hcss"
RUN_PROPERTY_ATTRS = {
    "bold": "data-hdocx-bold",
    "italic": "data-hdocx-italic",
    "font-size": "data-hdocx-font-size",
    "font-family": "data-hdocx-font-family",
    "ascii-font": "data-hdocx-ascii-font",
    "hansi-font": "data-hdocx-hansi-font",
    "east-asia-font": "data-hdocx-east-asia-font",
    "cs-font": "data-hdocx-cs-font",
    "color": "data-hdocx-color",
}
RUN_PROPERTY_ALIASES = {
    "font": "font-family",
    "eastAsia-font": "east-asia-font",
    "eastasia-font": "east-asia-font",
    "east-asian-font": "east-asia-font",
    "latin-font": "font-family",
}
DRAWING_PROPERTY_ATTRS = {
    "alt": "data-hdocx-alt",
    "width-emu": "data-hdocx-width-emu",
    "height-emu": "data-hdocx-height-emu",
}
IMAGE_FORMATTING_PROPERTIES = {
    "alt",
    "width-emu",
    "height-emu",
}
IMAGE_PARAGRAPH_PROPERTY_PREFIX = "paragraph-"
PARAGRAPH_PROPERTY_ATTRS = {
    "align": "data-hdocx-align",
    "first-line-indent": "data-hdocx-first-line-indent",
    "line-spacing": "data-hdocx-line-spacing",
    "space-before": "data-hdocx-space-before",
    "space-after": "data-hdocx-space-after",
}
PARAGRAPH_PROPERTY_ALIASES = {
    "text-align": "align",
    "alignment": "align",
    "line-spacing-exact": "line-spacing",
    "line-spacingExact": "line-spacing",
}
PARAGRAPH_STRUCTURAL_PROPERTIES = {
    "manual-page-break-before",
    "insert-empty-paragraph-before",
    "insert-empty-paragraph-after",
}
EMPTY_PARAGRAPH_PROPERTY_PREFIX = "empty-paragraph-"
EMPTY_PARAGRAPH_PROPERTY_ALIASES = {
    "line-spacing-exact": "line-spacing",
    "line-spacingExact": "line-spacing",
}
EMPTY_PARAGRAPH_SUPPORTED_PROPERTIES = {
    "style-id",
    "line-spacing",
    "space-before",
    "space-after",
}
OOXML_PROPERTY_MAP = {
    "bold": "w:rPr/w:b",
    "italic": "w:rPr/w:i",
    "font-size": "w:rPr/w:sz half-points",
    "font-family": "w:rPr/w:rFonts @w:ascii and @w:hAnsi",
    "ascii-font": "w:rPr/w:rFonts @w:ascii",
    "hansi-font": "w:rPr/w:rFonts @w:hAnsi",
    "east-asia-font": "w:rPr/w:rFonts @w:eastAsia",
    "cs-font": "w:rPr/w:rFonts @w:cs",
    "color": "w:rPr/w:color @w:val",
    "align": "w:pPr/w:jc @w:val",
    "first-line-indent": "w:pPr/w:ind @w:firstLine or @w:firstLineChars",
    "line-spacing": "w:pPr/w:spacing @w:line and @w:lineRule",
    "space-before": "w:pPr/w:spacing @w:before or @w:beforeLines",
    "space-after": "w:pPr/w:spacing @w:after or @w:afterLines",
    "manual-page-break-before": "insert <w:p><w:r><w:br w:type=\"page\"/></w:r></w:p> before target paragraph",
    "insert-empty-paragraph-before": "insert empty <w:p> before target paragraph",
    "insert-empty-paragraph-after": "insert empty <w:p> after target paragraph",
    "empty-paragraph-style-id": "inserted empty paragraph w:pPr/w:pStyle @w:val",
    "empty-paragraph-line-spacing": "inserted empty paragraph w:pPr/w:spacing @w:line and @w:lineRule",
    "empty-paragraph-space-before": "inserted empty paragraph w:pPr/w:spacing @w:before or @w:beforeLines",
    "empty-paragraph-space-after": "inserted empty paragraph w:pPr/w:spacing @w:after or @w:afterLines",
}
SUPPORTED_EDIT_MODES = [
    "all-runs",
    "paragraph-formatting",
    "image-formatting",
    "paragraph-structure",
    "direct-formatting",
    "style-definition",
    "paragraph-style",
    "numbering-definition",
    "paragraph-numbering",
    "comment-text",
    "revision-action",
    "equation-omml",
]
SUPPORTED_HCSS_PROPERTIES = {
    "all-runs": [
        "hdocx-font-family",
        "hdocx-eastAsia-font",
        "hdocx-east-asia-font",
        "hdocx-ascii-font",
        "hdocx-hansi-font",
        "hdocx-cs-font",
        "hdocx-font-size",
        "hdocx-bold",
        "hdocx-italic",
        "hdocx-color",
    ],
    "paragraph-formatting": [
        "hdocx-text-align",
        "hdocx-align",
        "hdocx-first-line-indent",
        "hdocx-line-spacing",
        "hdocx-line-spacing-exact",
        "hdocx-space-before",
        "hdocx-space-after",
        "hdocx-manual-page-break-before",
    ],
    "image-formatting": [
        "hdocx-alt",
        "hdocx-width-emu",
        "hdocx-height-emu",
        "hdocx-paragraph-text-align",
        "hdocx-paragraph-align",
        "hdocx-paragraph-line-spacing",
        "hdocx-paragraph-line-spacing-exact",
        "hdocx-paragraph-space-before",
        "hdocx-paragraph-space-after",
    ],
    "paragraph-structure": [
        "hdocx-manual-page-break-before",
        "hdocx-insert-empty-paragraph-before",
        "hdocx-insert-empty-paragraph-after",
        "hdocx-empty-paragraph-style-id",
        "hdocx-empty-paragraph-line-spacing",
        "hdocx-empty-paragraph-line-spacing-exact",
        "hdocx-empty-paragraph-space-before",
        "hdocx-empty-paragraph-space-after",
    ],
    "direct-formatting": [
        "hdocx-text-align",
        "hdocx-align",
        "hdocx-first-line-indent",
        "hdocx-line-spacing",
        "hdocx-line-spacing-exact",
        "hdocx-space-before",
        "hdocx-space-after",
        "hdocx-font-family",
        "hdocx-eastAsia-font",
        "hdocx-east-asia-font",
        "hdocx-ascii-font",
        "hdocx-hansi-font",
        "hdocx-cs-font",
        "hdocx-font-size",
        "hdocx-bold",
        "hdocx-italic",
        "hdocx-color",
    ],
    "style-definition": [
        "hdocx-text-align",
        "hdocx-align",
        "hdocx-first-line-indent",
        "hdocx-line-spacing",
        "hdocx-line-spacing-exact",
        "hdocx-space-before",
        "hdocx-space-after",
        "hdocx-font-family",
        "hdocx-eastAsia-font",
        "hdocx-east-asia-font",
        "hdocx-ascii-font",
        "hdocx-hansi-font",
        "hdocx-cs-font",
        "hdocx-font-size",
        "hdocx-bold",
        "hdocx-italic",
        "hdocx-color",
    ],
    "paragraph-style": ["hdocx-style-id"],
    "numbering-definition": [
        "hdocx-num-format",
        "hdocx-level-text",
        "hdocx-start",
    ],
    "paragraph-numbering": [
        "hdocx-list-id",
        "hdocx-num-id",
        "hdocx-ilvl",
    ],
    "comment-text": ["hdocx-comment-text"],
    "revision-action": ["hdocx-revision-action"],
    "equation-omml": ["hdocx-omml-source"],
}
SUPPORTED_HCSS_AT_RULES = [
    "@hdocx-token",
    "@hdocx-set",
    "@hdocx-format",
    "@hdocx-include",
    "@hdocx-style",
    "@hdocx-delete-style",
    "@hdocx-list",
    "@hdocx-insert-image",
    "@hdocx-insert-table-row",
    "@hdocx-delete-table-row",
    "@hdocx-insert-table-column",
    "@hdocx-delete-table-column",
]
PARAGRAPH_NUMBERING_ATTRS = {
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
NUMBERING_PROPERTY_NAMES = {
    "num-format",
    "level-text",
    "start",
    "number-suffix",
    "num-indent-left",
    "num-indent-hanging",
    "num-indent-first-line",
}
IMAGE_CONTENT_TYPES = {
    ".png": "image/png",
    ".jpg": "image/jpeg",
    ".jpeg": "image/jpeg",
    ".gif": "image/gif",
    ".bmp": "image/bmp",
    ".tif": "image/tiff",
    ".tiff": "image/tiff",
}
PKG_REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
WORD_DOCUMENT_RELS_ENTRY = "word/_rels/document.xml.rels"
STYLES_CONTENT_TYPE = "application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"
NUMBERING_CONTENT_TYPE = "application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"
STYLES_REL_TYPE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles"
NUMBERING_REL_TYPE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering"


def create_docx(
    output_docx: Path,
    *,
    title: str | None = None,
    paragraphs: list[str] | None = None,
    template: str = "blank",
    force: bool = False,
    export_dir: Path | None = None,
) -> dict[str, Any]:
    create_report = create_canonical_docx(
        output_docx,
        title=title,
        paragraphs=paragraphs,
        template=template,
        force=force,
    )
    if export_dir is None:
        return create_report
    export_report = export_docx(output_docx, export_dir, force=force)
    return {
        **create_report,
        "export": export_report,
        "hdocx": str(export_dir.resolve()),
    }


def export_docx(input_docx: Path, output_dir: Path, *, force: bool = False) -> dict[str, Any]:
    input_docx = input_docx.resolve()
    output_dir = output_dir.resolve()
    if not input_docx.exists():
        raise HDocxError("PACKAGE_INPUT_NOT_FOUND", "Input DOCX does not exist.", {"path": str(input_docx)})

    entries = read_docx_entries(input_docx)
    ensure_new_dir(output_dir, force=force)

    original_dir = output_dir / "original"
    parts_dir = output_dir / "parts"
    original_docx = original_dir / "original.docx"
    copy_file(input_docx, original_docx)
    write_json(original_dir / "entries.json", entries)
    extract_entries(input_docx, parts_dir)

    body, nodes = project_main_document(parts_dir)
    numbering = load_numbering_definitions(parts_dir)
    styles = load_style_definitions(parts_dir)
    document_html = build_document_html(body)
    html_path = output_dir / HTML_NAME
    html_path.write_text(document_html, encoding="utf-8", newline="\n")

    preview_css = _default_preview_css()
    (output_dir / PREVIEW_CSS_NAME).write_text(preview_css, encoding="utf-8", newline="\n")

    agent_hcss = _default_agent_hcss()
    (output_dir / AGENT_HCSS_NAME).write_text(agent_hcss, encoding="utf-8", newline="\n")

    manifest = {
        "hdocxVersion": "0.1",
        "createdAt": datetime.now(timezone.utc).isoformat(),
        "sourceDocx": {
            "fileName": input_docx.name,
            "sha256": sha256_file(input_docx),
            "size": input_docx.stat().st_size,
        },
        "package": {"entries": entries},
        "parts": _part_manifest(entries),
        "numbering": numbering,
        "styles": styles,
        "nodes": nodes,
        "files": {
            HTML_NAME: {"sha256": sha256_file(html_path)},
            PREVIEW_CSS_NAME: {"sha256": sha256_file(output_dir / PREVIEW_CSS_NAME)},
            AGENT_HCSS_NAME: {"sha256": sha256_file(output_dir / AGENT_HCSS_NAME)},
            "original/original.docx": {"sha256": sha256_file(original_docx)},
        },
        "roundtripPolicy": "copy-original-if-unmodified",
    }
    write_json(output_dir / MANIFEST_NAME, manifest)
    report = {
        "ok": True,
        "command": "export",
        "input": str(input_docx),
        "output": str(output_dir),
        "entryCount": len(entries),
        "nodeCount": len(nodes),
    }
    _append_audit(output_dir, report)
    return report


def validate_hdocx(bundle_dir: Path) -> dict[str, Any]:
    bundle_dir = bundle_dir.resolve()
    manifest = _load_manifest(bundle_dir)
    errors: list[dict[str, Any]] = []
    warnings: list[dict[str, Any]] = []

    for rel_path in [HTML_NAME, PREVIEW_CSS_NAME, AGENT_HCSS_NAME, "original/original.docx"]:
        path = bundle_dir / rel_path
        if not path.exists():
            errors.append({"code": "BUNDLE_FILE_MISSING", "path": rel_path})

    if errors:
        return _validation_report(bundle_dir, False, errors, warnings)

    original = bundle_dir / "original" / "original.docx"
    expected_source_hash = manifest.get("sourceDocx", {}).get("sha256")
    actual_source_hash = sha256_file(original)
    if expected_source_hash != actual_source_hash:
        errors.append(
            {
                "code": "MANIFEST_ORIGINAL_HASH_MISMATCH",
                "expected": expected_source_hash,
                "actual": actual_source_hash,
            }
        )

    for rel_path, recorded in manifest.get("files", {}).items():
        path = bundle_dir / rel_path
        if not path.exists():
            errors.append({"code": "BUNDLE_FILE_MISSING", "path": rel_path})
            continue
        actual = sha256_file(path)
        if recorded.get("sha256") != actual and rel_path in {"original/original.docx"}:
            errors.append(
                {
                    "code": "MANIFEST_FILE_HASH_MISMATCH",
                    "path": rel_path,
                    "expected": recorded.get("sha256"),
                    "actual": actual,
                }
            )

    html_text = (bundle_dir / HTML_NAME).read_text(encoding="utf-8")
    html_ids = collect_hdocx_ids(html_text)
    duplicate_ids = sorted({node_id for node_id in html_ids if html_ids.count(node_id) > 1})
    if duplicate_ids:
        errors.append({"code": "HTML_DUPLICATE_HDOCX_ID", "ids": duplicate_ids})

    manifest_nodes = set(manifest.get("nodes", {}))
    unknown_html_ids = sorted(
        node_id
        for node_id in set(html_ids) - manifest_nodes
        if not (node_id.startswith("part-") or node_id.startswith("sec-"))
    )
    if unknown_html_ids:
        errors.append({"code": "HTML_UNKNOWN_HDOCX_ID", "ids": unknown_html_ids})

    missing_html_ids = sorted(node_id for node_id in manifest_nodes if node_id not in set(html_ids))
    if missing_html_ids:
        errors.append({"code": "HTML_MISSING_MANIFEST_NODE_ID", "ids": missing_html_ids[:50]})

    report = _validation_report(bundle_dir, not errors, errors, warnings)
    _append_audit(bundle_dir, {"command": "validate", **report})
    return report


def plan_hdocx(bundle_dir: Path) -> dict[str, Any]:
    bundle_dir = bundle_dir.resolve()
    validation = validate_hdocx(bundle_dir)
    if not validation["ok"]:
        return {"ok": False, "command": "plan", "validation": validation, "patches": []}

    analysis = _analyze_edits(bundle_dir)
    if not analysis["modified"]:
        report = {
            "ok": True,
            "command": "plan",
            "modified": False,
            "patches": [],
            "summary": "No HTML or H-CSS edits detected. apply will copy original DOCX bytes.",
        }
    elif analysis["errors"]:
        report = {
            "ok": False,
            "command": "plan",
            "modified": True,
            "patches": analysis["patches"],
            "errors": analysis["errors"],
        }
    else:
        report = {
            "ok": True,
            "command": "plan",
            "modified": True,
            "patches": analysis["patches"],
            "summary": f"Planned {len(analysis['patches'])} document patch(es).",
        }
    if analysis.get("hcss") is not None:
        report["hcss"] = analysis["hcss"]
    _append_audit(bundle_dir, report)
    return report


def apply_hdocx(bundle_dir: Path, output_docx: Path) -> dict[str, Any]:
    bundle_dir = bundle_dir.resolve()
    _validate_path_encodable(output_docx, "output")
    output_docx = output_docx.resolve()
    validation = validate_hdocx(bundle_dir)
    if not validation["ok"]:
        raise HDocxError("VALIDATION_FAILED", "H-DOCX bundle validation failed.", validation)

    analysis = _analyze_edits(bundle_dir)
    if analysis["errors"]:
        raise HDocxError(
            "PATCH_PLAN_FAILED",
            "H-DOCX edits cannot be safely applied.",
            {"errors": analysis["errors"]},
        )

    original = bundle_dir / "original" / "original.docx"
    if not analysis["patches"]:
        copy_file(original, output_docx)
        mode = "copy-original-if-unmodified"
        patch_count = 0
    else:
        xml_patches = [
            patch for patch in analysis["patches"] if patch["operation"] not in {"replace-media", "add-media"}
        ]
        media_patches = [
            patch for patch in analysis["patches"] if patch["operation"] in {"replace-media", "add-media"}
        ]
        grouped: dict[str, list[dict[str, Any]]] = {}
        for patch in xml_patches:
            grouped.setdefault(patch["entryName"], []).append(patch)
        modified_entries: dict[str, bytes] = {}
        with zipfile.ZipFile(original, "r") as zf:
            original_names = set(zf.namelist())
            for entry_name, entry_patches in grouped.items():
                if entry_name in original_names:
                    xml_bytes = zf.read(entry_name)
                elif entry_name.startswith("word/_rels/") and entry_name.endswith(".rels"):
                    xml_bytes = _empty_relationships_xml()
                elif entry_name == "word/styles.xml":
                    xml_bytes = _empty_styles_xml()
                elif entry_name == "word/numbering.xml":
                    xml_bytes = _empty_numbering_xml()
                else:
                    raise HDocxError(
                        "PATCH_ENTRY_NOT_FOUND",
                        "Patch target entry is not present in the original DOCX.",
                        {"entryName": entry_name},
                    )
                modified_entries[entry_name] = patch_document_run_text(xml_bytes, entry_patches)
        for patch in media_patches:
            source_path = bundle_dir / patch["sourcePath"]
            modified_entries[patch["entryName"]] = source_path.read_bytes()
        repack_docx_with_modified_entries(original, output_docx, modified_entries)
        mode = _apply_mode(xml_patches, media_patches)
        patch_count = len(analysis["patches"])
    report = {
        "ok": True,
        "command": "apply",
        "output": str(output_docx),
        "mode": mode,
        "patchCount": patch_count,
        "patchSummary": _summarize_patches(analysis["patches"]),
        "packageDiff": diff_docx(original, output_docx),
        "outputSha256": sha256_file(output_docx),
    }
    _append_audit(bundle_dir, report)
    return report


def roundtrip_docx(input_docx: Path, work_dir: Path, output_docx: Path, *, force: bool = False) -> dict[str, Any]:
    export_report = export_docx(input_docx, work_dir, force=force)
    apply_report = apply_hdocx(work_dir, output_docx)
    identical = sha256_file(input_docx.resolve()) == sha256_file(output_docx.resolve())
    return {
        "ok": identical,
        "command": "roundtrip",
        "export": export_report,
        "apply": apply_report,
        "byteIdentical": identical,
    }


def check_docx(input_docx: Path, work_dir: Path, output_docx: Path, *, force: bool = False) -> dict[str, Any]:
    roundtrip_report = roundtrip_docx(input_docx, work_dir, output_docx, force=force)
    diff_report = diff_docx(input_docx, output_docx)
    acceptance = {
        "byteIdentical": roundtrip_report.get("byteIdentical") is True and diff_report.get("byteIdentical") is True,
        "contentIdentical": diff_report.get("summary", {}).get("contentIdentical") is True,
        "semanticIdentical": diff_report.get("semanticDiff", {}).get("identical") is True,
        "changedEntries": diff_report.get("summary", {}).get("changedEntries", 0),
        "leftOnlyEntries": diff_report.get("summary", {}).get("leftOnlyEntries", 0),
        "rightOnlyEntries": diff_report.get("summary", {}).get("rightOnlyEntries", 0),
    }
    ok = all(
        (
            acceptance["byteIdentical"],
            acceptance["contentIdentical"],
            acceptance["semanticIdentical"],
            acceptance["changedEntries"] == 0,
            acceptance["leftOnlyEntries"] == 0,
            acceptance["rightOnlyEntries"] == 0,
        )
    )
    return {
        "ok": ok,
        "command": "check",
        "input": str(input_docx.resolve()),
        "work": str(work_dir.resolve()),
        "output": str(output_docx.resolve()),
        "acceptance": acceptance,
        "roundtrip": roundtrip_report,
        "diff": diff_report,
    }


def batch_check_docx(input_path: Path, work_dir: Path, output_dir: Path, *, force: bool = False) -> dict[str, Any]:
    input_path = input_path.resolve()
    work_dir = work_dir.resolve()
    output_dir = output_dir.resolve()
    _validate_batch_output_paths(input_path, work_dir, output_dir)
    ensure_new_dir(work_dir, force=force)
    ensure_new_dir(output_dir, force=force)
    inputs = _collect_docx_inputs(input_path)

    results: list[dict[str, Any]] = []
    for index, docx_path in enumerate(inputs, start=1):
        relative_path = docx_path.name if input_path.is_file() else docx_path.relative_to(input_path).as_posix()
        output_stem = _safe_output_stem(relative_path, index)
        item_work = work_dir / f"{output_stem}.hdocx"
        item_output = output_dir / f"{output_stem}.docx"
        try:
            item_report = check_docx(docx_path, item_work, item_output, force=True)
            results.append(
                {
                    "ok": item_report["ok"],
                    "input": str(docx_path),
                    "relativePath": relative_path,
                    "work": str(item_work),
                    "output": str(item_output),
                    "acceptance": item_report["acceptance"],
                    "report": item_report,
                }
            )
        except HDocxError as exc:
            results.append(
                {
                    "ok": False,
                    "input": str(docx_path),
                    "relativePath": relative_path,
                    "work": str(item_work),
                    "output": str(item_output),
                    "error": exc.to_dict(),
                }
            )
        except Exception as exc:
            results.append(
                {
                    "ok": False,
                    "input": str(docx_path),
                    "relativePath": relative_path,
                    "work": str(item_work),
                    "output": str(item_output),
                    "error": {
                        "code": "UNEXPECTED_ERROR",
                        "severity": "error",
                        "message": str(exc),
                    },
                }
            )
    passed = sum(1 for item in results if item["ok"])
    failed = len(results) - passed
    return {
        "ok": failed == 0,
        "command": "batch-check",
        "input": str(input_path),
        "work": str(work_dir),
        "output": str(output_dir),
        "summary": {
            "total": len(results),
            "passed": passed,
            "failed": failed,
        },
        "results": results,
    }


def diff_docx(left_docx: Path, right_docx: Path) -> dict[str, Any]:
    left_docx = left_docx.resolve()
    right_docx = right_docx.resolve()
    diff = compare_docx_entries(left_docx, right_docx)
    left_sha = sha256_file(left_docx)
    right_sha = sha256_file(right_docx)
    byte_identical = left_sha == right_sha
    semantic_diff = _semantic_diff_docx(left_docx, right_docx)
    fragment_diff = _fragment_diff_docx(left_docx, right_docx, diff, semantic_diff)
    manual_page_break_diff = _manual_page_break_diff_docx(left_docx, right_docx)
    empty_paragraph_diff = _empty_paragraph_diff_docx(left_docx, right_docx)
    return {
        "ok": True,
        "command": "diff",
        "left": {
            "path": str(left_docx),
            "sha256": left_sha,
            "size": left_docx.stat().st_size,
        },
        "right": {
            "path": str(right_docx),
            "sha256": right_sha,
            "size": right_docx.stat().st_size,
        },
        "byteIdentical": byte_identical,
        "summary": _diff_summary(diff, byte_identical),
        "semanticDiff": semantic_diff,
        "fragmentDiff": fragment_diff,
        "manualPageBreakDiff": manual_page_break_diff,
        "emptyParagraphDiff": empty_paragraph_diff,
        **diff,
    }


def audit_docx(input_docx: Path) -> dict[str, Any]:
    input_docx = input_docx.resolve()
    if not input_docx.exists():
        raise HDocxError("AUDIT_INPUT_NOT_FOUND", "Input DOCX does not exist.", {"path": str(input_docx)})
    entry_bytes = _read_docx_entry_bytes(input_docx)
    entries = sorted(entry_bytes)
    xml_entries = {name: data for name, data in entry_bytes.items() if name.lower().endswith(".xml")}
    rel_entries = {name: data for name, data in entry_bytes.items() if name.lower().endswith(".rels")}
    features = {
        "customXml": _entry_feature(entries, "customXml/"),
        "chart": _entry_or_relationship_feature(entries, rel_entries, ["word/charts/", "charts/"], ["chart"]),
        "smartArt": _entry_or_relationship_feature(entries, rel_entries, ["word/diagrams/", "diagrams/"], ["diagram"]),
        "ole": _entry_or_relationship_feature(entries, rel_entries, ["word/embeddings/", "embeddings/"], ["oleObject", "relationships/package"]),
        "alternateContent": _xml_local_feature(xml_entries, "AlternateContent"),
        "vml": _xml_any_local_feature(xml_entries, ["pict", "shape", "imagedata"]),
        "textBox": _xml_any_local_feature(xml_entries, ["txbxContent", "textbox"]),
        "field": _xml_any_local_feature(xml_entries, ["fldChar", "instrText", "fldSimple"]),
        "equation": _xml_any_local_feature(xml_entries, ["oMath", "oMathPara"]),
        "revision": _xml_any_local_feature(xml_entries, ["ins", "del", "moveFrom", "moveTo"]),
        "comment": _entry_or_xml_feature(entries, xml_entries, ["word/comments.xml"], ["commentRangeStart", "commentRangeEnd", "commentReference"]),
        "footnote": _entry_feature(entries, "word/footnotes.xml"),
        "endnote": _entry_feature(entries, "word/endnotes.xml"),
        "headerFooter": _entry_any_prefix_feature(entries, ["word/header", "word/footer"]),
        "image": _entry_or_relationship_feature(entries, rel_entries, ["word/media/"], ["image"]),
    }
    for name, feature in features.items():
        feature["policy"] = _audit_policy(name)
    high_risk = [
        name
        for name in ("customXml", "chart", "smartArt", "ole", "alternateContent", "vml", "textBox")
        if features[name]["present"]
    ]
    protected = [
        name
        for name in ("field", "equation", "revision", "comment")
        if features[name]["present"]
    ]
    return {
        "ok": True,
        "command": "audit",
        "input": str(input_docx),
        "sha256": sha256_file(input_docx),
        "entryCount": len(entries),
        "xmlEntryCount": len(xml_entries),
        "features": features,
        "summary": {
            "highRiskPresent": high_risk,
            "protectedStructurePresent": protected,
            "hasNonEditableAdvancedObjects": bool(high_risk or protected),
        },
    }


def inspect_node(bundle_dir: Path, node_id: str) -> dict[str, Any]:
    manifest = _load_manifest(bundle_dir.resolve())
    node = manifest.get("nodes", {}).get(node_id)
    if node is None:
        raise HDocxError("INSPECT_NODE_NOT_FOUND", "Node id not found in manifest.", {"nodeId": node_id})
    return {"ok": True, "command": "inspect", "kind": "node", "node": node}


def inspect_hdocx(bundle_dir: Path, target_id: str, *, kind: str = "node") -> dict[str, Any]:
    manifest = _load_manifest(bundle_dir.resolve())
    normalized_kind = kind.lower().strip()
    if normalized_kind == "node":
        return inspect_node(bundle_dir, target_id)
    if normalized_kind == "style":
        return _inspect_style(manifest, target_id)
    if normalized_kind == "list":
        return _inspect_list(manifest, target_id)
    if normalized_kind == "table":
        return _inspect_table(manifest, target_id)
    if normalized_kind == "image":
        return _inspect_image(manifest, target_id)
    raise HDocxError(
        "INSPECT_KIND_UNSUPPORTED",
        "Unsupported inspect kind.",
        {"kind": kind, "supportedKinds": ["node", "style", "list", "table", "image"]},
    )


def query_hdocx(
    bundle_dir: Path,
    *,
    kind: str = "paragraph",
    text: str | None = None,
    text_regex: str | None = None,
    style_id: str | None = None,
    font_family: str | None = None,
    east_asia_font: str | None = None,
    font_size: str | None = None,
    bold: bool | None = None,
    italic: bool | None = None,
    color: str | None = None,
    align: str | None = None,
    line_spacing: str | None = None,
    space_before: str | None = None,
    space_after: str | None = None,
    has_image: bool | None = None,
    suspected_heading_level1: bool | None = None,
    include_runs: bool = True,
    limit: int = 100,
) -> dict[str, Any]:
    bundle_dir = bundle_dir.resolve()
    validation = validate_hdocx(bundle_dir)
    if not validation["ok"]:
        return {"ok": False, "command": "query", "validation": validation, "matches": []}
    manifest = _load_manifest(bundle_dir)
    criteria = _query_criteria(
        kind=kind,
        text=text,
        text_regex=text_regex,
        style_id=style_id,
        font_family=font_family,
        east_asia_font=east_asia_font,
        font_size=font_size,
        bold=bold,
        italic=italic,
        color=color,
        align=align,
        line_spacing=line_spacing,
        space_before=space_before,
        space_after=space_after,
        has_image=has_image,
        suspected_heading_level1=suspected_heading_level1,
        include_runs=include_runs,
        limit=limit,
    )
    nodes = manifest.get("nodes", {})
    normalized_kind = _normalize_query_kind(kind)
    candidates = _query_candidates(nodes, normalized_kind)
    matches: list[dict[str, Any]] = []
    for node_id in candidates:
        node = nodes[node_id]
        match = _query_match(bundle_dir, manifest, node_id, node, criteria)
        if match is not None:
            matches.append(match)
    limit = max(1, int(limit or 100))
    report = {
        "ok": True,
        "command": "query",
        "kind": normalized_kind,
        "criteria": criteria,
        "matchCount": len(matches),
        "returnedCount": min(len(matches), limit),
        "truncated": len(matches) > limit,
        "matches": matches[:limit],
    }
    _append_audit(bundle_dir, {key: value for key, value in report.items() if key != "matches"})
    return report


def find_hdocx(bundle_dir: Path, **kwargs: Any) -> dict[str, Any]:
    report = query_hdocx(bundle_dir, **kwargs)
    report["command"] = "find"
    return report


def assert_hdocx(bundle_dir: Path, assertions: list[Any] | None = None) -> dict[str, Any]:
    bundle_dir = bundle_dir.resolve()
    validation = validate_hdocx(bundle_dir)
    if not validation["ok"]:
        return {"ok": False, "command": "assert", "validation": validation, "assertions": []}
    manifest = _load_manifest(bundle_dir)
    normalized_assertions = _normalize_assertions(assertions)
    planned_context: dict[str, Any] | None = None
    results: list[dict[str, Any]] = []
    for assertion in normalized_assertions:
        assertion_type = _normalize_assertion_type(str(assertion.get("type", "")))
        if _assertion_after_apply(assertion):
            if planned_context is None:
                planned_context = _planned_assertion_context(bundle_dir, manifest)
            if assertion_type == "text-payload-unchanged":
                result = _assert_text_payload_unchanged_after_apply(
                    manifest,
                    planned_context["manifest"],
                    assertion_type,
                )
            else:
                translated_assertion = _translate_assertion_for_planned_output(assertion, planned_context["nodeIdMap"])
                result = _run_assertion(planned_context["bundle"], planned_context["manifest"], translated_assertion)
                if translated_assertion.get("nodeIdTranslations"):
                    result["nodeIdTranslations"] = translated_assertion["nodeIdTranslations"]
            result["scope"] = "planned-output"
            result["afterApply"] = True
            result["plannedOutput"] = {
                "docx": str(planned_context["docx"]),
                "bundle": str(planned_context["bundle"]),
            }
        else:
            result = _run_assertion(bundle_dir, manifest, assertion)
            result["scope"] = "planned-patches" if assertion_type == "text-payload-unchanged" else "current-bundle"
            result["afterApply"] = False
        results.append(result)
    failed = [result for result in results if not result.get("passed", False)]
    report = {
        "ok": not failed,
        "command": "assert",
        "assertionCount": len(results),
        "passedCount": len(results) - len(failed),
        "failedCount": len(failed),
        "assertions": results,
    }
    _append_audit(bundle_dir, report)
    return report


def doctor_report() -> dict[str, Any]:
    loaded_module_path = Path(__file__).resolve()
    distribution_version = _package_version()
    package_version = __version__
    git_commit = _git_commit_for_module(loaded_module_path)
    module_mtime = _mtime_iso(loaded_module_path)
    return {
        "ok": True,
        "command": "doctor",
        "python": {
            "version": platform.python_version(),
            "implementation": platform.python_implementation(),
            "executable": sys.executable,
        },
        "package": {
            "name": "html_docx",
            "distributionName": "html-docx",
            "version": package_version,
            "sourceVersion": __version__,
            "distributionVersion": distribution_version,
            "dependencies": [],
            "loadedModulePath": str(loaded_module_path),
            "moduleMtimeUtc": module_mtime,
            "gitCommit": git_commit,
        },
        "runtime": {
            "loaded_module_path": str(loaded_module_path),
            "package_version": package_version,
            "git_commit_or_build_time": git_commit or module_mtime,
            "supported_edit_modes": SUPPORTED_EDIT_MODES,
            "supported_hcss_properties": SUPPORTED_HCSS_PROPERTIES,
            "supported_query_kinds": list(QUERY_KINDS),
            "supported_assertions": list(ASSERTION_TYPES),
        },
        "capabilities": {
            "requiresExternalRuntime": False,
            "requiresNetwork": False,
            "bundleFormat": "H-DOCX",
            "creation": {"available": True, "templates": ["blank"]},
            "query": {"available": True, "kinds": list(QUERY_KINDS)},
            "assertions": {"available": True, "types": list(ASSERTION_TYPES)},
            "hcss": {
                "supportedEditModes": SUPPORTED_EDIT_MODES,
                "supportedProperties": SUPPORTED_HCSS_PROPERTIES,
                "supportedAtRules": SUPPORTED_HCSS_AT_RULES,
            },
        },
    }


def _package_version() -> str | None:
    try:
        return importlib.metadata.version("html-docx")
    except importlib.metadata.PackageNotFoundError:
        return None


def _mtime_iso(path: Path) -> str | None:
    try:
        return datetime.fromtimestamp(path.stat().st_mtime, tz=timezone.utc).isoformat()
    except OSError:
        return None


def _git_commit_for_module(module_path: Path) -> str | None:
    for parent in [module_path.parent, *module_path.parents]:
        git_dir = parent / ".git"
        if not git_dir.exists():
            continue
        head = git_dir / "HEAD"
        try:
            head_text = head.read_text(encoding="utf-8").strip()
        except OSError:
            return None
        if head_text.startswith("ref: "):
            ref_path = git_dir / head_text.removeprefix("ref: ").strip()
            try:
                return ref_path.read_text(encoding="utf-8").strip() or None
            except OSError:
                packed_refs = git_dir / "packed-refs"
                try:
                    ref_name = head_text.removeprefix("ref: ").strip()
                    for line in packed_refs.read_text(encoding="utf-8").splitlines():
                        if line.startswith("#") or not line.strip():
                            continue
                        commit, _, name = line.partition(" ")
                        if name == ref_name:
                            return commit
                except OSError:
                    return None
                return None
        return head_text or None
    return None


def _inspect_style(manifest: dict[str, Any], style_id: str) -> dict[str, Any]:
    style = manifest.get("styles", {}).get(style_id)
    if style is None:
        raise HDocxError("INSPECT_STYLE_NOT_FOUND", "Style id not found in manifest.", {"styleId": style_id})
    usages = [
        _node_summary(node_id, node)
        for node_id, node in sorted(manifest.get("nodes", {}).items(), key=lambda item: _node_sort_key(manifest.get("nodes", {}), item[0]))
        if node.get("kind") == "paragraph" and node.get("styleId") == style_id
    ]
    return {
        "ok": True,
        "command": "inspect",
        "kind": "style",
        "style": style,
        "usageCount": len(usages),
        "paragraphs": usages[:50],
        "truncated": len(usages) > 50,
    }


def _inspect_list(manifest: dict[str, Any], list_id: str) -> dict[str, Any]:
    numbering = manifest.get("numbering", {})
    num = numbering.get("nums", {}).get(list_id)
    if num is None:
        created = manifest.get("createdLists", {}).get(list_id)
        if created is not None:
            num = {"numId": created.get("numId"), "abstractNumId": created.get("abstractNumId"), "createdListId": list_id}
        else:
            raise HDocxError("INSPECT_LIST_NOT_FOUND", "List numId not found in manifest.", {"listId": list_id})
    abstract_num_id = str(num.get("abstractNumId")) if num.get("abstractNumId") is not None else None
    abstract = numbering.get("abstractNums", {}).get(abstract_num_id) if abstract_num_id is not None else None
    usages = [
        _node_summary(node_id, node)
        for node_id, node in sorted(manifest.get("nodes", {}).items(), key=lambda item: _node_sort_key(manifest.get("nodes", {}), item[0]))
        if node.get("kind") == "paragraph" and str(node.get("numbering", {}).get("numId")) == str(num.get("numId", list_id))
    ]
    return {
        "ok": True,
        "command": "inspect",
        "kind": "list",
        "list": {
            "num": num,
            "abstractNum": abstract,
        },
        "usageCount": len(usages),
        "paragraphs": usages[:50],
        "truncated": len(usages) > 50,
    }


def _inspect_table(manifest: dict[str, Any], table_id: str) -> dict[str, Any]:
    nodes = manifest.get("nodes", {})
    table = nodes.get(table_id)
    if table is None:
        raise HDocxError("INSPECT_TABLE_NOT_FOUND", "Table id not found in manifest.", {"tableId": table_id})
    if table.get("kind") != "table":
        raise HDocxError(
            "INSPECT_TABLE_KIND_MISMATCH",
            "Inspect kind table requires a table node id.",
            {"tableId": table_id, "actualKind": table.get("kind")},
        )
    rows: list[dict[str, Any]] = []
    for row_id in table.get("children", []):
        row = nodes.get(row_id, {})
        row_summary = _node_summary(row_id, row)
        cells: list[dict[str, Any]] = []
        for cell_id in row.get("children", []):
            cell = nodes.get(cell_id, {})
            cell_summary = _node_summary(cell_id, cell)
            paragraphs: list[dict[str, Any]] = []
            for paragraph_id in cell.get("children", []):
                paragraph = nodes.get(paragraph_id, {})
                paragraph_summary = _node_summary(paragraph_id, paragraph)
                paragraph_summary["runs"] = [
                    _node_summary(run_id, nodes[run_id])
                    for run_id in paragraph.get("children", [])
                    if nodes.get(run_id, {}).get("kind") == "run"
                ]
                paragraphs.append(paragraph_summary)
            cell_summary["paragraphs"] = paragraphs
            cells.append(cell_summary)
        row_summary["cells"] = cells
        rows.append(row_summary)
    return {
        "ok": True,
        "command": "inspect",
        "kind": "table",
        "table": _node_summary(table_id, table),
        "rowCount": len(rows),
        "rows": rows,
    }


def _inspect_image(manifest: dict[str, Any], image_id: str) -> dict[str, Any]:
    nodes = manifest.get("nodes", {})
    node = nodes.get(image_id)
    if node is None:
        raise HDocxError("INSPECT_IMAGE_NOT_FOUND", "Image run node id not found in manifest.", {"imageId": image_id})
    if node.get("objectKind") != "drawing":
        raise HDocxError(
            "INSPECT_IMAGE_KIND_MISMATCH",
            "Inspect kind image requires a drawing run node id.",
            {"imageId": image_id, "actualKind": node.get("kind"), "objectKind": node.get("objectKind")},
        )
    media_parts = {
        part_path: part
        for part_path, part in sorted(manifest.get("parts", {}).items())
        if part_path.startswith("/word/media/")
    }
    return {
        "ok": True,
        "command": "inspect",
        "kind": "image",
        "image": _node_summary(image_id, node),
        "mediaParts": media_parts,
    }


QUERY_KINDS = ("paragraph", "run", "image", "all")
ASSERTION_TYPES = (
    "text-payload-unchanged",
    "paragraphs-have-empty-before",
    "paragraphs-have-empty-after",
    "images-host-paragraph-not-exact-line-spacing",
    "level1-headings-have-empty-paragraph-before",
)
DEFAULT_LEVEL1_HEADING_EXCLUDE_REGEX = (
    r"^\s*(?:"
    r"\u6458\s*\u8981|\u6458\u8981|abstract|"
    r"\u76ee\s*\u5f55|\u76ee\u5f55|contents?|table\s+of\s+contents|"
    r"\u5173\u952e\u8bcd|key\s*words?"
    r")\s*[:\uff1a]?\s*$"
)


def _normalize_query_kind(kind: str) -> str:
    normalized = (kind or "paragraph").strip().lower()
    aliases = {
        "p": "paragraph",
        "paragraphs": "paragraph",
        "r": "run",
        "runs": "run",
        "images": "image",
        "drawing": "image",
        "drawings": "image",
    }
    normalized = aliases.get(normalized, normalized)
    if normalized not in QUERY_KINDS:
        raise HDocxError("QUERY_KIND_UNSUPPORTED", "Unsupported query kind.", {"kind": kind, "supportedKinds": list(QUERY_KINDS)})
    return normalized


def _query_criteria(**kwargs: Any) -> dict[str, Any]:
    criteria = {key: value for key, value in kwargs.items() if value is not None}
    if "font_size" in criteria:
        criteria["font_size"] = _normalize_query_pt(criteria["font_size"])
    for key in ("font_family", "east_asia_font", "color", "align", "line_spacing", "space_before", "space_after", "style_id"):
        if key in criteria and isinstance(criteria[key], str):
            criteria[key] = criteria[key].strip()
    if "text_regex" in criteria:
        try:
            re.compile(str(criteria["text_regex"]))
        except re.error as exc:
            raise HDocxError("QUERY_INVALID_TEXT_REGEX", "textRegex is not a valid regular expression.", {"textRegex": criteria["text_regex"], "reason": str(exc)}) from exc
    return criteria


def _query_candidates(nodes: dict[str, Any], kind: str) -> list[str]:
    candidates: list[str] = []
    for node_id, node in sorted(nodes.items(), key=lambda item: _node_sort_key(nodes, item[0])):
        node_kind = node.get("kind")
        if kind == "all":
            if node_kind in {"paragraph", "run"}:
                candidates.append(node_id)
        elif kind == "image":
            if node_kind == "run" and node.get("objectKind") == "drawing":
                candidates.append(node_id)
        elif node_kind == kind:
            candidates.append(node_id)
    return candidates


def _query_match(bundle_dir: Path, manifest: dict[str, Any], node_id: str, node: dict[str, Any], criteria: dict[str, Any]) -> dict[str, Any] | None:
    nodes = manifest.get("nodes", {})
    styles = manifest.get("styles", {})
    if not _query_text_matches(nodes, node_id, node, criteria):
        return None
    if not _query_paragraph_matches(manifest, node_id, node, criteria):
        return None
    if not _query_run_properties_match(manifest, node_id, node, criteria):
        return None
    if "has_image" in criteria and _node_has_image(nodes, node_id, node) != bool(criteria["has_image"]):
        return None
    if "suspected_heading_level1" in criteria:
        paragraph_id = node_id if node.get("kind") == "paragraph" else node.get("parent")
        paragraph = nodes.get(paragraph_id or "", {})
        guess = _heading_level1_guess(paragraph_id or "", paragraph, nodes, styles)
        if bool(guess["suspected"]) != bool(criteria["suspected_heading_level1"]):
            return None
    if node.get("kind") == "paragraph":
        return _paragraph_query_summary(bundle_dir, manifest, node_id, node, bool(criteria.get("include_runs", True)))
    return _run_query_summary(bundle_dir, manifest, node_id, node)


def _query_text_matches(nodes: dict[str, Any], node_id: str, node: dict[str, Any], criteria: dict[str, Any]) -> bool:
    text = _node_visible_text(nodes, node_id, node)
    expected = criteria.get("text")
    if expected is not None and str(expected) not in text:
        return False
    pattern = criteria.get("text_regex")
    if pattern is not None and re.search(str(pattern), text) is None:
        return False
    return True


def _query_paragraph_matches(manifest: dict[str, Any], node_id: str, node: dict[str, Any], criteria: dict[str, Any]) -> bool:
    nodes = manifest.get("nodes", {})
    paragraph_id = node_id if node.get("kind") == "paragraph" else node.get("parent")
    paragraph = nodes.get(paragraph_id or "")
    if paragraph is None:
        return not any(key in criteria for key in ("style_id", "align", "line_spacing", "space_before", "space_after"))
    if "style_id" in criteria and str(paragraph.get("styleId") or "") != str(criteria["style_id"]):
        return False
    props = _effective_paragraph_properties(manifest.get("styles", {}), paragraph)
    for criteria_key, prop_key in (
        ("align", "align"),
        ("line_spacing", "line-spacing"),
        ("space_before", "space-before"),
        ("space_after", "space-after"),
    ):
        if criteria_key in criteria and not _query_value_equal(props.get(prop_key), criteria[criteria_key]):
            return False
    return True


def _query_run_properties_match(manifest: dict[str, Any], node_id: str, node: dict[str, Any], criteria: dict[str, Any]) -> bool:
    if not any(key in criteria for key in ("font_family", "east_asia_font", "font_size", "bold", "italic", "color")):
        return True
    nodes = manifest.get("nodes", {})
    styles = manifest.get("styles", {})
    if node.get("kind") == "run":
        parent = nodes.get(node.get("parent", ""))
        props = _effective_run_properties(styles, node, parent)
        return _run_properties_match_criteria(props, criteria)
    if node.get("kind") != "paragraph":
        return False
    return bool(_matching_runs_for_paragraph(manifest, node_id, node, criteria))


def _matching_runs_for_paragraph(manifest: dict[str, Any], node_id: str, paragraph: dict[str, Any], criteria: dict[str, Any]) -> list[dict[str, Any]]:
    nodes = manifest.get("nodes", {})
    styles = manifest.get("styles", {})
    matched: list[dict[str, Any]] = []
    for run_id in paragraph.get("children", []):
        run = nodes.get(run_id)
        if run is None or run.get("kind") != "run":
            continue
        props = _effective_run_properties(styles, run, paragraph)
        if _run_properties_match_criteria(props, criteria):
            matched.append(_node_summary(run_id, run))
    if matched:
        return matched
    style_props = _style_effective_properties(styles, paragraph.get("styleId"), "runProperties")
    if style_props and _run_properties_match_criteria(style_props, criteria):
        return [{"nodeId": node_id, "kind": "paragraph-style-runProperties", "properties": style_props}]
    return []


def _run_properties_match_criteria(props: dict[str, Any], criteria: dict[str, Any]) -> bool:
    if "font_family" in criteria and not _font_value_matches(props, str(criteria["font_family"])):
        return False
    if "east_asia_font" in criteria and not _query_value_equal(props.get("east-asia-font"), criteria["east_asia_font"]):
        return False
    if "font_size" in criteria and not _query_value_equal(_normalize_query_pt(props.get("font-size")), criteria["font_size"]):
        return False
    if "bold" in criteria and _prop_bool(props.get("bold")) != bool(criteria["bold"]):
        return False
    if "italic" in criteria and _prop_bool(props.get("italic")) != bool(criteria["italic"]):
        return False
    if "color" in criteria and not _query_value_equal(props.get("color"), criteria["color"]):
        return False
    return True


def _font_value_matches(props: dict[str, Any], expected: str) -> bool:
    expected_folded = expected.casefold()
    values = [
        props.get("font-family"),
        props.get("ascii-font"),
        props.get("hansi-font"),
        props.get("east-asia-font"),
        props.get("cs-font"),
    ]
    return any(isinstance(value, str) and value.casefold() == expected_folded for value in values)


def _query_value_equal(left: Any, right: Any) -> bool:
    if left is None:
        return False
    return str(left).strip().casefold() == str(right).strip().casefold()


def _normalize_query_pt(value: Any) -> str | None:
    if value is None:
        return None
    if isinstance(value, (int, float)):
        return f"{float(value):g}pt"
    text = str(value).strip()
    if not text:
        return text
    if re.fullmatch(r"\d+(?:\.\d+)?", text):
        return f"{float(text):g}pt"
    return text


def _prop_bool(value: Any) -> bool:
    if isinstance(value, bool):
        return value
    if value is None:
        return False
    return str(value).strip().lower() in {"1", "true", "yes", "on"}


def _node_visible_text(nodes: dict[str, Any], node_id: str, node: dict[str, Any]) -> str:
    if node.get("kind") == "paragraph":
        return _paragraph_visible_text(nodes, node_id, node)
    return str(node.get("text") or "")


def _paragraph_visible_text(nodes: dict[str, Any], node_id: str, paragraph: dict[str, Any]) -> str:
    chunks: list[str] = []
    for child_id in paragraph.get("children", []):
        child = nodes.get(child_id, {})
        if child.get("kind") in {"run", "protected"}:
            chunks.append(str(child.get("text") or ""))
    return "".join(chunks)


def _node_has_image(nodes: dict[str, Any], node_id: str, node: dict[str, Any]) -> bool:
    if node.get("kind") == "run":
        return node.get("objectKind") == "drawing"
    if node.get("kind") != "paragraph":
        return False
    return any(nodes.get(child_id, {}).get("objectKind") == "drawing" for child_id in node.get("children", []))


def _effective_paragraph_properties(styles: dict[str, Any], paragraph: dict[str, Any]) -> dict[str, Any]:
    props = _style_effective_properties(styles, paragraph.get("styleId"), "paragraphProperties")
    props.update(paragraph.get("properties", {}) or {})
    return props


def _effective_run_properties(styles: dict[str, Any], run: dict[str, Any], paragraph: dict[str, Any] | None) -> dict[str, Any]:
    props = _style_effective_properties(styles, paragraph.get("styleId") if paragraph else None, "runProperties")
    props.update(run.get("properties", {}) or {})
    return props


def _style_effective_properties(styles: dict[str, Any], style_id: str | None, property_key: str) -> dict[str, Any]:
    effective: dict[str, Any] = {}
    seen: set[str] = set()

    def visit(current_id: str | None) -> None:
        if not current_id or current_id in seen:
            return
        seen.add(current_id)
        style = styles.get(current_id)
        if not style:
            return
        visit(style.get("basedOn"))
        effective.update(style.get(property_key, {}) or {})

    visit(style_id)
    return effective


def _paragraph_query_summary(bundle_dir: Path, manifest: dict[str, Any], node_id: str, paragraph: dict[str, Any], include_runs: bool) -> dict[str, Any]:
    nodes = manifest.get("nodes", {})
    styles = manifest.get("styles", {})
    summary = _node_summary(node_id, paragraph)
    summary["text"] = _paragraph_visible_text(nodes, node_id, paragraph)
    summary["textPreview"] = _preview_text(summary["text"])
    summary["effectiveProperties"] = _effective_paragraph_properties(styles, paragraph)
    summary["hasImage"] = _node_has_image(nodes, node_id, paragraph)
    summary["imageRuns"] = [
        _run_query_summary(bundle_dir, manifest, child_id, nodes[child_id])
        for child_id in paragraph.get("children", [])
        if nodes.get(child_id, {}).get("objectKind") == "drawing"
    ]
    summary["suspectedHeadingLevel1"] = _heading_level1_guess(node_id, paragraph, nodes, styles)
    if include_runs:
        summary["runs"] = [
            _run_query_summary(bundle_dir, manifest, child_id, nodes[child_id])
            for child_id in paragraph.get("children", [])
            if nodes.get(child_id, {}).get("kind") == "run"
        ]
    return summary


def _run_query_summary(bundle_dir: Path, manifest: dict[str, Any], node_id: str, run: dict[str, Any]) -> dict[str, Any]:
    nodes = manifest.get("nodes", {})
    styles = manifest.get("styles", {})
    paragraph = nodes.get(run.get("parent", ""))
    summary = _node_summary(node_id, run)
    summary["textPreview"] = _preview_text(str(run.get("text") or ""))
    summary["effectiveProperties"] = _effective_run_properties(styles, run, paragraph)
    if paragraph:
        host = _node_summary(run["parent"], paragraph)
        host["text"] = _paragraph_visible_text(nodes, run["parent"], paragraph)
        host["textPreview"] = _preview_text(host["text"])
        host["effectiveProperties"] = _effective_paragraph_properties(styles, paragraph)
        host["lineSpacingRule"] = _paragraph_line_spacing_rule(bundle_dir, paragraph)
        summary["hostParagraph"] = host
    return summary


def _preview_text(text: str, limit: int = 120) -> str:
    collapsed = re.sub(r"\s+", " ", text).strip()
    if len(collapsed) <= limit:
        return collapsed
    return collapsed[: limit - 1] + "..."


def _heading_level1_guess(node_id: str, paragraph: dict[str, Any], nodes: dict[str, Any], styles: dict[str, Any]) -> dict[str, Any]:
    if paragraph.get("kind") != "paragraph":
        return {"suspected": False, "score": 0, "reasons": []}
    reasons: list[str] = []
    score = 0
    style_id = str(paragraph.get("styleId") or "")
    style = styles.get(style_id, {})
    style_name = str(style.get("name") or "")
    style_text = f"{style_id} {style_name}".casefold().replace(" ", "")
    if any(marker in style_text for marker in ("heading1", "标题1", "title1")):
        reasons.append("style-level-1")
        score += 4
    paragraph_props = _effective_paragraph_properties(styles, paragraph)
    if paragraph_props.get("align") == "center":
        reasons.append("centered")
        score += 1
    run_props = _paragraph_effective_run_props_for_heading(paragraph, nodes, styles)
    max_size = max((_pt_number(props.get("font-size")) or 0 for props in run_props), default=0)
    if max_size >= 14:
        reasons.append(f"large-font-{max_size:g}pt")
        score += 1
    if any(_prop_bool(props.get("bold")) for props in run_props):
        reasons.append("bold")
        score += 1
    if any(_looks_like_heading_font(props) for props in run_props):
        reasons.append("heading-font")
        score += 1
    numbering = paragraph.get("numbering", {}) or {}
    if str(numbering.get("ilvl", "")) == "0":
        reasons.append("top-level-numbering")
        score += 1
    text = _paragraph_visible_text(nodes, node_id, paragraph)
    if 0 < len(text.strip()) <= 80:
        reasons.append("short-text")
        score += 1
    suspected = score >= 3 or "style-level-1" in reasons
    return {"suspected": suspected, "score": score, "reasons": reasons}


def _paragraph_effective_run_props_for_heading(paragraph: dict[str, Any], nodes: dict[str, Any], styles: dict[str, Any]) -> list[dict[str, Any]]:
    props: list[dict[str, Any]] = []
    for run_id in paragraph.get("children", []):
        run = nodes.get(run_id, {})
        if run.get("kind") == "run":
            props.append(_effective_run_properties(styles, run, paragraph))
    if props:
        return props
    style_props = _style_effective_properties(styles, paragraph.get("styleId"), "runProperties")
    return [style_props] if style_props else []


def _looks_like_heading_font(props: dict[str, Any]) -> bool:
    values = " ".join(str(props.get(key, "")) for key in ("font-family", "ascii-font", "hansi-font", "east-asia-font", "cs-font")).casefold()
    return any(marker in values for marker in ("黑体", "simhei", "hei", "heiti", "bold"))


def _pt_number(value: Any) -> float | None:
    if value is None:
        return None
    match = re.fullmatch(r"\s*(\d+(?:\.\d+)?)\s*pt\s*", str(value), flags=re.IGNORECASE)
    if match is None:
        return None
    return float(match.group(1))


def _normalize_assertions(assertions: list[Any] | None) -> list[dict[str, Any]]:
    raw_assertions = assertions if assertions else [{"type": "text-payload-unchanged"}]
    normalized: list[dict[str, Any]] = []
    for assertion in raw_assertions:
        if isinstance(assertion, str):
            normalized.append({"type": assertion})
        elif isinstance(assertion, dict):
            normalized.append(assertion)
        else:
            normalized.append({"type": "invalid", "value": assertion})
    return normalized


def _assertion_after_apply(assertion: dict[str, Any]) -> bool:
    return bool(
        assertion.get("afterApply")
        or assertion.get("after_apply")
        or assertion.get("plannedOutput")
        or assertion.get("planned_output")
    )


def _planned_assertion_context(bundle_dir: Path, original_manifest: dict[str, Any]) -> dict[str, Any]:
    scratch = bundle_dir / ".hdocx-assert-after-apply"
    ensure_new_dir(scratch, force=True)
    planned_docx = scratch / "planned-output.docx"
    planned_bundle = scratch / "planned-output.hdocx"
    apply_hdocx(bundle_dir, planned_docx)
    export_docx(planned_docx, planned_bundle, force=True)
    planned_manifest = _load_manifest(planned_bundle)
    return {
        "docx": planned_docx,
        "bundle": planned_bundle,
        "manifest": planned_manifest,
        "nodeIdMap": _map_nodes_between_manifests(original_manifest, planned_manifest),
    }


def _translate_assertion_for_planned_output(assertion: dict[str, Any], node_id_map: dict[str, str]) -> dict[str, Any]:
    translated = copy.deepcopy(assertion)
    translations: dict[str, str] = {}
    for key in ("paragraphIds", "paragraph_ids"):
        value = translated.get(key)
        if not isinstance(value, list):
            continue
        translated_ids: list[Any] = []
        for item in value:
            original_id = str(item)
            planned_id = node_id_map.get(original_id, original_id)
            translated_ids.append(planned_id)
            if planned_id != original_id:
                translations[original_id] = planned_id
        translated[key] = translated_ids
    if translations:
        translated["nodeIdTranslations"] = translations
    return translated


def _map_nodes_between_manifests(left_manifest: dict[str, Any], right_manifest: dict[str, Any]) -> dict[str, str]:
    left_nodes = left_manifest.get("nodes", {})
    right_nodes = right_manifest.get("nodes", {})
    right_buckets: dict[tuple[Any, ...], list[str]] = {}
    for node_id, node in sorted(right_nodes.items(), key=lambda item: _node_sort_key(right_nodes, item[0])):
        signature = _node_mapping_signature(right_nodes, node_id, node)
        if signature is not None:
            right_buckets.setdefault(signature, []).append(node_id)
    mapping: dict[str, str] = {}
    for node_id, node in sorted(left_nodes.items(), key=lambda item: _node_sort_key(left_nodes, item[0])):
        signature = _node_mapping_signature(left_nodes, node_id, node)
        if signature is None:
            continue
        candidates = right_buckets.get(signature, [])
        if candidates:
            mapping[node_id] = candidates.pop(0)
    return mapping


def _node_mapping_signature(nodes: dict[str, Any], node_id: str, node: dict[str, Any]) -> tuple[Any, ...] | None:
    kind = node.get("kind")
    if kind == "paragraph":
        return (
            "paragraph",
            node.get("partPath"),
            _paragraph_visible_text(nodes, node_id, node),
            _node_has_image(nodes, node_id, node),
        )
    if kind == "run":
        return (
            "run",
            node.get("partPath"),
            str(node.get("text") or ""),
            node.get("objectKind"),
        )
    return None


def _run_assertion(bundle_dir: Path, manifest: dict[str, Any], assertion: dict[str, Any]) -> dict[str, Any]:
    assertion_type = _normalize_assertion_type(str(assertion.get("type", "")))
    if assertion_type not in ASSERTION_TYPES:
        return {
            "type": assertion.get("type"),
            "passed": False,
            "code": "ASSERTION_UNSUPPORTED",
            "message": "Unsupported assertion type.",
            "supportedTypes": list(ASSERTION_TYPES),
        }
    if assertion_type == "text-payload-unchanged":
        return _assert_text_payload_unchanged(bundle_dir, assertion_type)
    if assertion_type == "paragraphs-have-empty-before":
        return _assert_paragraphs_have_empty_neighbor(manifest, assertion, assertion_type, "before")
    if assertion_type == "paragraphs-have-empty-after":
        return _assert_paragraphs_have_empty_neighbor(manifest, assertion, assertion_type, "after")
    if assertion_type == "images-host-paragraph-not-exact-line-spacing":
        return _assert_images_host_paragraph_not_exact_line_spacing(bundle_dir, manifest, assertion_type)
    if assertion_type == "level1-headings-have-empty-paragraph-before":
        return _assert_level1_headings_have_empty_before(manifest, assertion, assertion_type)
    raise AssertionError(assertion_type)


def _normalize_assertion_type(assertion_type: str) -> str:
    normalized = assertion_type.strip().lower().replace("_", "-")
    aliases = {
        "text-unchanged": "text-payload-unchanged",
        "no-text-change": "text-payload-unchanged",
        "paragraph-empty-before": "paragraphs-have-empty-before",
        "paragraph-empty-after": "paragraphs-have-empty-after",
        "images-no-exact-line-spacing": "images-host-paragraph-not-exact-line-spacing",
        "image-host-not-exact": "images-host-paragraph-not-exact-line-spacing",
        "level1-empty-before": "level1-headings-have-empty-paragraph-before",
    }
    return aliases.get(normalized, normalized)


def _assert_text_payload_unchanged(bundle_dir: Path, assertion_type: str) -> dict[str, Any]:
    analysis = _analyze_edits(bundle_dir)
    text_patches = [
        patch
        for patch in analysis.get("patches", [])
        if ("newText" in patch and patch.get("newText") != patch.get("oldText")) or patch.get("operation") == "patch-comment-text"
    ]
    failures = [
        {
            "patchId": patch.get("id"),
            "operation": patch.get("operation"),
            "nodeId": patch.get("nodeId"),
            "oldTextPreview": _preview_text(str(patch.get("oldText", ""))),
            "newTextPreview": _preview_text(str(patch.get("newText", ""))),
        }
        for patch in text_patches
    ]
    errors = analysis.get("errors", [])
    return {
        "type": assertion_type,
        "passed": not failures and not errors,
        "checkedPatchCount": len(analysis.get("patches", [])),
        "failures": failures,
        "errors": errors,
    }


def _assert_text_payload_unchanged_after_apply(
    original_manifest: dict[str, Any],
    planned_manifest: dict[str, Any],
    assertion_type: str,
) -> dict[str, Any]:
    original_payload = _manifest_text_payload(original_manifest)
    planned_payload = _manifest_text_payload(planned_manifest)
    failures: list[dict[str, Any]] = []
    if original_payload != planned_payload:
        first_difference = _first_sequence_difference(original_payload, planned_payload)
        failures.append(
            {
                "reason": "text-payload-differs-after-apply",
                "originalTextItemCount": len(original_payload),
                "plannedTextItemCount": len(planned_payload),
                "firstDifference": first_difference,
            }
        )
    return {
        "type": assertion_type,
        "passed": not failures,
        "checkedOriginalTextItemCount": len(original_payload),
        "checkedPlannedTextItemCount": len(planned_payload),
        "failures": failures,
    }


def _manifest_text_payload(manifest: dict[str, Any]) -> list[str]:
    nodes = manifest.get("nodes", {})
    payload: list[str] = []
    for _node_id, node in sorted(nodes.items(), key=lambda item: _node_sort_key(nodes, item[0])):
        if node.get("kind") in {"run", "protected"}:
            text = str(node.get("text") or "")
            if text:
                payload.append(text)
    return payload


def _first_sequence_difference(left: list[str], right: list[str]) -> dict[str, Any] | None:
    for index, (left_item, right_item) in enumerate(zip(left, right)):
        if left_item != right_item:
            return {
                "index": index,
                "original": _preview_text(left_item),
                "planned": _preview_text(right_item),
            }
    if len(left) != len(right):
        index = min(len(left), len(right))
        return {
            "index": index,
            "original": _preview_text(left[index]) if index < len(left) else None,
            "planned": _preview_text(right[index]) if index < len(right) else None,
        }
    return None


def _assert_paragraphs_have_empty_neighbor(manifest: dict[str, Any], assertion: dict[str, Any], assertion_type: str, direction: str) -> dict[str, Any]:
    nodes = manifest.get("nodes", {})
    paragraph_ids = [str(value) for value in assertion.get("paragraphIds", assertion.get("paragraph_ids", []))]
    failures: list[dict[str, Any]] = []
    if not paragraph_ids:
        return {
            "type": assertion_type,
            "passed": False,
            "code": "ASSERTION_MISSING_PARAGRAPH_IDS",
            "message": "paragraphIds is required for this assertion.",
            "failures": [],
        }
    for paragraph_id in paragraph_ids:
        paragraph = nodes.get(paragraph_id)
        if paragraph is None or paragraph.get("kind") != "paragraph":
            failures.append({"nodeId": paragraph_id, "reason": "paragraph-not-found"})
            continue
        neighbor_id = _neighbor_paragraph_id(nodes, paragraph_id, direction)
        neighbor = nodes.get(neighbor_id or "")
        if neighbor is None or not _is_empty_projected_paragraph(nodes, neighbor):
            failures.append({"nodeId": paragraph_id, "neighborId": neighbor_id, "reason": f"empty-paragraph-{direction}-missing"})
    return {
        "type": assertion_type,
        "passed": not failures,
        "checkedParagraphIds": paragraph_ids,
        "failures": failures,
    }


def _assert_images_host_paragraph_not_exact_line_spacing(bundle_dir: Path, manifest: dict[str, Any], assertion_type: str) -> dict[str, Any]:
    nodes = manifest.get("nodes", {})
    failures: list[dict[str, Any]] = []
    checked: list[dict[str, Any]] = []
    for run_id, run in sorted(nodes.items(), key=lambda item: _node_sort_key(nodes, item[0])):
        if run.get("kind") != "run" or run.get("objectKind") != "drawing":
            continue
        paragraph_id = run.get("parent")
        paragraph = nodes.get(paragraph_id or "")
        rule = _paragraph_line_spacing_rule(bundle_dir, paragraph or {})
        checked.append({"imageRunId": run_id, "hostParagraphId": paragraph_id, "lineSpacingRule": rule})
        if rule.get("lineRule") == "exact":
            failures.append({"imageRunId": run_id, "hostParagraphId": paragraph_id, "lineSpacingRule": rule})
    return {
        "type": assertion_type,
        "passed": not failures,
        "checkedImageCount": len(checked),
        "checkedImages": checked,
        "failures": failures,
    }


def _assert_level1_headings_have_empty_before(manifest: dict[str, Any], assertion: dict[str, Any], assertion_type: str) -> dict[str, Any]:
    nodes = manifest.get("nodes", {})
    styles = manifest.get("styles", {})
    min_score = int(assertion.get("minScore", assertion.get("min_score", 3)))
    allow_first = bool(assertion.get("allowFirstInPart", assertion.get("allow_first_in_part", True)))
    explicit_paragraph_ids = [str(value) for value in assertion.get("paragraphIds", assertion.get("paragraph_ids", []))]
    regexes = _level1_assertion_regexes(assertion)
    if regexes.get("error"):
        return {
            "type": assertion_type,
            "passed": False,
            "code": "ASSERTION_INVALID_REGEX",
            "message": "includeRegex or excludeRegex is invalid.",
            "error": regexes["error"],
            "failures": [],
        }
    headings: list[dict[str, Any]] = []
    skipped: list[dict[str, Any]] = []
    failures: list[dict[str, Any]] = []
    candidates: list[tuple[str, dict[str, Any]]]
    if explicit_paragraph_ids:
        candidates = [(paragraph_id, nodes.get(paragraph_id, {})) for paragraph_id in explicit_paragraph_ids]
    else:
        candidates = [
            (paragraph_id, paragraph)
            for paragraph_id, paragraph in sorted(nodes.items(), key=lambda item: _node_sort_key(nodes, item[0]))
            if paragraph.get("kind") == "paragraph"
        ]
    for paragraph_id, paragraph in candidates:
        if paragraph.get("kind") != "paragraph":
            failures.append({"nodeId": paragraph_id, "reason": "paragraph-not-found"})
            continue
        guess = _heading_level1_guess(paragraph_id, paragraph, nodes, styles)
        text = _paragraph_visible_text(nodes, paragraph_id, paragraph)
        if not explicit_paragraph_ids and (not guess["suspected"] or int(guess["score"]) < min_score):
            continue
        skip_reason = None if explicit_paragraph_ids else _level1_assertion_skip_reason(text, regexes)
        if skip_reason:
            skipped.append({"nodeId": paragraph_id, "textPreview": _preview_text(text), "guess": guess, "reason": skip_reason})
            continue
        neighbor_id = _neighbor_paragraph_id(nodes, paragraph_id, "before")
        if neighbor_id is None and allow_first:
            skipped.append({"nodeId": paragraph_id, "textPreview": _preview_text(text), "guess": guess, "reason": "first-in-part-allowed"})
            continue
        headings.append({"nodeId": paragraph_id, "textPreview": _preview_text(text), "guess": guess, "previousParagraphId": neighbor_id})
        neighbor = nodes.get(neighbor_id or "")
        if neighbor is None or not _is_empty_projected_paragraph(nodes, neighbor):
            failures.append({"nodeId": paragraph_id, "previousParagraphId": neighbor_id, "reason": "empty-paragraph-before-missing", "guess": guess})
    return {
        "type": assertion_type,
        "passed": not failures,
        "checkedHeadingCount": len(headings),
        "explicitParagraphIds": explicit_paragraph_ids,
        "criteria": {
            "minScore": min_score,
            "allowFirstInPart": allow_first,
            "includeRegex": regexes.get("includePattern"),
            "excludeRegex": regexes.get("excludePattern"),
            "useDefaultExcludes": regexes.get("useDefaultExcludes"),
            "defaultExcludeRegex": regexes.get("defaultExcludePattern"),
        },
        "headings": headings,
        "skippedHeadings": skipped,
        "failures": failures,
    }


def _level1_assertion_regexes(assertion: dict[str, Any]) -> dict[str, Any]:
    include_pattern = assertion.get("includeRegex", assertion.get("include_regex"))
    exclude_pattern = assertion.get("excludeRegex", assertion.get("exclude_regex"))
    use_default_excludes = bool(assertion.get("useDefaultExcludes", assertion.get("use_default_excludes", True)))
    result: dict[str, Any] = {
        "includePattern": include_pattern,
        "excludePattern": exclude_pattern,
        "defaultExcludePattern": DEFAULT_LEVEL1_HEADING_EXCLUDE_REGEX if use_default_excludes else None,
        "useDefaultExcludes": use_default_excludes,
    }
    try:
        if include_pattern:
            result["include"] = re.compile(str(include_pattern), flags=re.IGNORECASE)
        if exclude_pattern:
            result["exclude"] = re.compile(str(exclude_pattern), flags=re.IGNORECASE)
        if use_default_excludes:
            result["defaultExclude"] = re.compile(DEFAULT_LEVEL1_HEADING_EXCLUDE_REGEX, flags=re.IGNORECASE)
    except re.error as exc:
        result["error"] = {"pattern": include_pattern or exclude_pattern, "reason": str(exc)}
    return result


def _level1_assertion_skip_reason(text: str, regexes: dict[str, Any]) -> str | None:
    include_regex = regexes.get("include")
    if include_regex is not None and include_regex.search(text) is None:
        return "include-regex-mismatch"
    exclude_regex = regexes.get("exclude")
    if exclude_regex is not None and exclude_regex.search(text) is not None:
        return "exclude-regex-match"
    default_exclude_regex = regexes.get("defaultExclude")
    if default_exclude_regex is not None and default_exclude_regex.search(text) is not None:
        return "default-exclude-regex-match"
    return None


def _neighbor_paragraph_id(nodes: dict[str, Any], paragraph_id: str, direction: str) -> str | None:
    paragraph = nodes.get(paragraph_id)
    if paragraph is None:
        return None
    locator = paragraph.get("locator") or {}
    paragraph_index = locator.get("paragraphIndex")
    part_path = paragraph.get("partPath")
    if paragraph_index is None or part_path is None:
        return None
    target_index = paragraph_index - 1 if direction == "before" else paragraph_index + 1
    for candidate_id, candidate in nodes.items():
        candidate_locator = candidate.get("locator") or {}
        if candidate.get("kind") == "paragraph" and candidate.get("partPath") == part_path and candidate_locator.get("paragraphIndex") == target_index:
            return candidate_id
    return None


def _is_empty_projected_paragraph(nodes: dict[str, Any], paragraph: dict[str, Any]) -> bool:
    if paragraph.get("kind") != "paragraph":
        return False
    for child_id in paragraph.get("children", []):
        child = nodes.get(child_id, {})
        if child.get("objectKind") == "drawing":
            return False
        if child.get("protectedKind"):
            return False
        if str(child.get("text") or ""):
            return False
    return True


def _paragraph_line_spacing_rule(bundle_dir: Path, paragraph: dict[str, Any]) -> dict[str, Any]:
    locator = paragraph.get("locator") or {}
    paragraph_index = locator.get("paragraphIndex")
    part_path = paragraph.get("partPath")
    if paragraph_index is None or not part_path:
        return {"available": False, "reason": "paragraph-locator-missing"}
    xml_path = bundle_dir / "parts" / part_path.lstrip("/")
    if not xml_path.exists():
        return {"available": False, "reason": "part-missing", "partPath": part_path}
    try:
        root = ET.parse(xml_path).getroot()
    except ET.ParseError as exc:
        return {"available": False, "reason": "xml-parse-error", "partPath": part_path, "message": str(exc)}
    paragraphs = root.findall(".//w:p", NS)
    if paragraph_index < 1 or paragraph_index > len(paragraphs):
        return {"available": False, "reason": "paragraph-index-out-of-range", "partPath": part_path, "paragraphIndex": paragraph_index}
    spacing = paragraphs[paragraph_index - 1].find("w:pPr/w:spacing", NS)
    if spacing is None:
        return {"available": True, "lineRule": None, "line": None}
    return {
        "available": True,
        "lineRule": spacing.attrib.get(_tag("lineRule")) or "auto",
        "line": spacing.attrib.get(_tag("line")),
        "before": spacing.attrib.get(_tag("before")),
        "after": spacing.attrib.get(_tag("after")),
    }


def _diff_summary(diff: dict[str, Any], byte_identical: bool) -> dict[str, Any]:
    counts = diff.get("entryCounts", {})
    return {
        "byteIdentical": byte_identical,
        "contentIdentical": diff.get("identical", False),
        "zipMetadataIdentical": diff.get("zipMetadataIdentical", False),
        "changedEntries": counts.get("changed", 0),
        "metadataChangedEntries": counts.get("metadataChanged", 0),
        "leftOnlyEntries": counts.get("leftOnly", 0),
        "rightOnlyEntries": counts.get("rightOnly", 0),
        "unchangedEntries": counts.get("unchanged", 0),
    }


def _summarize_patches(patches: list[dict[str, Any]]) -> dict[str, Any]:
    by_entry: dict[str, int] = {}
    by_operation: dict[str, int] = {}
    by_risk: dict[str, int] = {}
    targets: list[dict[str, Any]] = []
    for patch in patches:
        entry_name = patch.get("entryName", "")
        operation = patch.get("operation", "")
        risk = _patch_risk_class(operation)
        by_entry[entry_name] = by_entry.get(entry_name, 0) + 1
        by_operation[operation] = by_operation.get(operation, 0) + 1
        by_risk[risk] = by_risk.get(risk, 0) + 1
        target: dict[str, Any] = {
            "patchId": patch.get("id"),
            "operation": operation,
            "entryName": entry_name,
            "riskClass": risk,
        }
        for key in ("nodeId", "styleId", "partPath"):
            if key in patch:
                target[key] = patch[key]
        targets.append(target)
    return {
        "count": len(patches),
        "byEntry": dict(sorted(by_entry.items())),
        "byOperation": dict(sorted(by_operation.items())),
        "byRiskClass": dict(sorted(by_risk.items())),
        "targets": targets,
    }


def _patch_risk_class(operation: str) -> str:
    if operation in {"patch-run", "patch-paragraph", "patch-drawing-alt", "patch-drawing-properties"}:
        return "fragment-preserving-eligible"
    if operation in {"replace-media", "add-media"}:
        return "binary-package-entry"
    if operation in {"patch-relationships", "patch-content-types"}:
        return "package-metadata"
    if operation.startswith("insert-image-"):
        return "structural-insert"
    if operation in {
        "split-run",
        "patch-style",
        "create-style",
        "delete-style",
        "patch-numbering-level",
        "create-numbering-list",
        "patch-paragraph-style",
        "patch-paragraph-numbering",
        "insert-table-row-after",
        "delete-table-row",
        "insert-table-column-after",
        "delete-table-column",
        "insert-manual-page-break-before",
        "insert-empty-paragraph-before",
        "insert-empty-paragraph-after",
        "patch-comment-text",
        "patch-revision-action",
        "patch-equation-omml",
    }:
        return "xml-entry-reserialize"
    return "unknown"


def _apply_mode(xml_patches: list[dict[str, Any]], media_patches: list[dict[str, Any]]) -> str:
    if xml_patches and media_patches:
        return "document-and-media-patch"
    if media_patches:
        return "media-patch"
    return "document-patch"


def _empty_relationships_xml() -> bytes:
    return (
        b'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        b'<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"/>'
    )


def _empty_styles_xml() -> bytes:
    return (
        b'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        b'<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"/>'
    )


def _empty_numbering_xml() -> bytes:
    return (
        b'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        b'<w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"/>'
    )


def _semantic_diff_docx(left_docx: Path, right_docx: Path) -> dict[str, Any]:
    try:
        _, left_nodes = project_docx_entries(_read_docx_entry_bytes(left_docx))
        _, right_nodes = project_docx_entries(_read_docx_entry_bytes(right_docx))
        alignment = _manual_page_break_semantic_alignment(left_docx, right_docx, left_nodes, right_nodes)
        if alignment.get("applied"):
            left_nodes = alignment["leftNodes"]
            right_nodes = alignment["rightNodes"]
    except Exception as exc:
        return {
            "available": False,
            "error": {
                "code": "SEMANTIC_DIFF_FAILED",
                "message": str(exc),
            },
        }
    left_ids = set(left_nodes)
    right_ids = set(right_nodes)
    common_ids = left_ids & right_ids
    changed = [
        node_id
        for node_id in sorted(common_ids, key=lambda item: _node_sort_key(left_nodes, item))
        if _node_semantic_changed(left_nodes[node_id], right_nodes[node_id])
    ]
    left_only = sorted(left_ids - right_ids, key=lambda item: _node_sort_key(left_nodes, item))
    right_only = sorted(right_ids - left_ids, key=lambda item: _node_sort_key(right_nodes, item))
    result = {
        "available": True,
        "identical": not changed and not left_only and not right_only,
        "nodeCounts": {
            "left": len(left_nodes),
            "right": len(right_nodes),
            "common": len(common_ids),
            "changed": len(changed),
            "leftOnly": len(left_only),
            "rightOnly": len(right_only),
        },
        "changed": changed,
        "leftOnly": left_only,
        "rightOnly": right_only,
        "changedNodes": [_node_diff(node_id, left_nodes[node_id], right_nodes[node_id]) for node_id in changed],
        "leftOnlyNodes": [_node_summary(node_id, left_nodes[node_id]) for node_id in left_only],
        "rightOnlyNodes": [_node_summary(node_id, right_nodes[node_id]) for node_id in right_only],
    }
    if alignment.get("applied"):
        result.update(alignment.get("reports", {}))
    return result


def _manual_page_break_semantic_alignment(
    left_docx: Path,
    right_docx: Path,
    left_nodes: dict[str, Any],
    right_nodes: dict[str, Any],
) -> dict[str, Any]:
    left_breaks = _manual_page_break_paragraphs(left_docx)
    right_breaks = _manual_page_break_paragraphs(right_docx)
    left_empty_paragraphs = _empty_paragraphs(left_docx)
    right_empty_paragraphs = _empty_paragraphs(right_docx)
    left_break_indices = {int(item["paragraphIndex"]) for item in left_breaks}
    right_break_indices = {int(item["paragraphIndex"]) for item in right_breaks}
    left_empty_indices = {int(item["paragraphIndex"]) for item in left_empty_paragraphs}
    right_empty_indices = {int(item["paragraphIndex"]) for item in right_empty_paragraphs}
    if left_break_indices == right_break_indices and left_empty_indices == right_empty_indices:
        return {"applied": False}

    left_manual_skip = _structural_paragraph_node_ids(left_nodes, left_break_indices)
    right_manual_skip = _structural_paragraph_node_ids(right_nodes, right_break_indices)
    left_empty_skip = _structural_paragraph_node_ids(left_nodes, left_empty_indices)
    right_empty_skip = _structural_paragraph_node_ids(right_nodes, right_empty_indices)
    left_skip = left_manual_skip | left_empty_skip
    right_skip = right_manual_skip | right_empty_skip
    if not left_skip and not right_skip:
        return {"applied": False}

    left_order = [node_id for node_id in sorted(left_nodes, key=lambda item: _node_sort_key(left_nodes, item)) if node_id not in left_skip]
    right_order = [node_id for node_id in sorted(right_nodes, key=lambda item: _node_sort_key(right_nodes, item)) if node_id not in right_skip]
    pair_count = min(len(left_order), len(right_order))
    right_to_left: dict[str, str] = {}
    for index in range(pair_count):
        left_id = left_order[index]
        right_id = right_order[index]
        left_node = left_nodes[left_id]
        right_node = right_nodes[right_id]
        if left_node.get("kind") != right_node.get("kind") or left_node.get("partPath") != right_node.get("partPath"):
            return {"applied": False}
        right_to_left[right_id] = left_id

    aligned_left = {
        node_id: _drop_structural_children(left_nodes[node_id], left_skip)
        for node_id in left_order
    }
    aligned_right: dict[str, Any] = {}
    for right_id, left_id in right_to_left.items():
        aligned_right[left_id] = _remap_semantic_node_for_structural_paragraph(
            right_nodes[right_id],
            left_nodes[left_id],
            right_to_left,
            right_skip,
        )
    for right_id in right_order[pair_count:]:
        aligned_right[right_id] = right_nodes[right_id]

    return {
        "applied": True,
        "leftNodes": aligned_left,
        "rightNodes": aligned_right,
        "reports": {
            "manualPageBreakAlignment": {
                "strategy": "ignore-manual-page-break-paragraphs",
                "ignoredLeftBreakParagraphs": len(left_break_indices),
                "ignoredRightBreakParagraphs": len(right_break_indices),
                "ignoredLeftNodeIds": sorted(left_manual_skip, key=lambda item: _node_sort_key(left_nodes, item)),
                "ignoredRightNodeIds": sorted(right_manual_skip, key=lambda item: _node_sort_key(right_nodes, item)),
            },
            "emptyParagraphAlignment": {
                "strategy": "ignore-empty-paragraphs",
                "ignoredLeftEmptyParagraphs": len(left_empty_indices),
                "ignoredRightEmptyParagraphs": len(right_empty_indices),
                "ignoredLeftNodeIds": sorted(left_empty_skip, key=lambda item: _node_sort_key(left_nodes, item)),
                "ignoredRightNodeIds": sorted(right_empty_skip, key=lambda item: _node_sort_key(right_nodes, item)),
            },
        },
    }


def _structural_paragraph_node_ids(nodes: dict[str, Any], paragraph_indices: set[int]) -> set[str]:
    if not paragraph_indices:
        return set()
    paragraph_ids = {
        node_id
        for node_id, node in nodes.items()
        if node.get("kind") == "paragraph" and int((node.get("locator") or {}).get("paragraphIndex") or 0) in paragraph_indices
    }
    skipped = set(paragraph_ids)
    for node_id, node in nodes.items():
        if node.get("parent") in paragraph_ids:
            skipped.add(node_id)
    return skipped


def _remap_semantic_node_for_structural_paragraph(
    right_node: dict[str, Any],
    left_node: dict[str, Any],
    right_to_left: dict[str, str],
    right_skip: set[str],
) -> dict[str, Any]:
    remapped = copy.deepcopy(right_node)
    remapped["locator"] = copy.deepcopy(left_node.get("locator"))
    parent = remapped.get("parent")
    if parent in right_to_left:
        remapped["parent"] = right_to_left[parent]
    if "children" in remapped:
        remapped["children"] = [
            right_to_left.get(child_id, child_id)
            for child_id in remapped.get("children", [])
            if child_id not in right_skip
        ]
    if all(remapped.get(field) == left_node.get(field) for field in NODE_DIFF_FIELDS):
        remapped["hash"] = left_node.get("hash")
    return remapped


def _drop_structural_children(node: dict[str, Any], skip: set[str]) -> dict[str, Any]:
    if "children" not in node or not skip:
        return node
    remapped = copy.deepcopy(node)
    remapped["children"] = [child_id for child_id in remapped.get("children", []) if child_id not in skip]
    return remapped


def _fragment_diff_docx(
    left_docx: Path,
    right_docx: Path,
    package_diff: dict[str, Any],
    semantic_diff: dict[str, Any],
) -> dict[str, Any]:
    try:
        left_entries = _read_docx_entry_bytes(left_docx)
        right_entries = _read_docx_entry_bytes(right_docx)
    except Exception as exc:
        return {
            "available": False,
            "error": {
                "code": "FRAGMENT_DIFF_FAILED",
                "message": str(exc),
            },
        }

    semantic_links = _semantic_links_by_entry(semantic_diff)
    changed_entries = []
    for entry_name in package_diff.get("changed", []):
        changed_entries.append(
            {
                "path": entry_name,
                "status": "changed",
                "kind": _package_diff_kind(package_diff, entry_name, "changedEntries"),
                "byteDiff": _byte_change_region(left_entries[entry_name], right_entries[entry_name]),
                "linkedNodes": semantic_links.get(entry_name, []),
            }
        )
    for entry_name in package_diff.get("leftOnly", []):
        data = left_entries[entry_name]
        changed_entries.append(
            {
                "path": entry_name,
                "status": "left-only",
                "kind": _package_diff_kind(package_diff, entry_name, "leftOnlyEntries"),
                "byteDiff": _one_sided_byte_region(data, "left"),
                "linkedNodes": semantic_links.get(entry_name, []),
            }
        )
    for entry_name in package_diff.get("rightOnly", []):
        data = right_entries[entry_name]
        changed_entries.append(
            {
                "path": entry_name,
                "status": "right-only",
                "kind": _package_diff_kind(package_diff, entry_name, "rightOnlyEntries"),
                "byteDiff": _one_sided_byte_region(data, "right"),
                "linkedNodes": semantic_links.get(entry_name, []),
            }
        )
    linked_node_count = sum(len(entry.get("linkedNodes", [])) for entry in changed_entries)
    return {
        "available": True,
        "entryCount": len(changed_entries),
        "linkedNodeCount": linked_node_count,
        "summary": {
            "changedEntries": len(package_diff.get("changed", [])),
            "leftOnlyEntries": len(package_diff.get("leftOnly", [])),
            "rightOnlyEntries": len(package_diff.get("rightOnly", [])),
            "linkedNodes": linked_node_count,
        },
        "entries": changed_entries,
    }


def _manual_page_break_diff_docx(left_docx: Path, right_docx: Path) -> dict[str, Any]:
    try:
        left_breaks = _manual_page_break_paragraphs(left_docx)
        right_breaks = _manual_page_break_paragraphs(right_docx)
    except Exception as exc:
        return {
            "available": False,
            "error": {
                "code": "MANUAL_PAGE_BREAK_DIFF_FAILED",
                "message": str(exc),
            },
        }
    left_keys = {_manual_break_key(item) for item in left_breaks}
    right_keys = {_manual_break_key(item) for item in right_breaks}
    added = [item for item in right_breaks if _manual_break_key(item) not in left_keys]
    removed = [item for item in left_breaks if _manual_break_key(item) not in right_keys]
    return {
        "available": True,
        "identical": not added and not removed,
        "summary": {
            "leftCount": len(left_breaks),
            "rightCount": len(right_breaks),
            "added": len(added),
            "removed": len(removed),
        },
        "added": added,
        "removed": removed,
    }


def _empty_paragraph_diff_docx(left_docx: Path, right_docx: Path) -> dict[str, Any]:
    try:
        left_empty = _empty_paragraphs(left_docx)
        right_empty = _empty_paragraphs(right_docx)
    except Exception as exc:
        return {
            "available": False,
            "error": {
                "code": "EMPTY_PARAGRAPH_DIFF_FAILED",
                "message": str(exc),
            },
        }
    left_keys = {_empty_paragraph_key(item) for item in left_empty}
    right_keys = {_empty_paragraph_key(item) for item in right_empty}
    left_by_key = {_empty_paragraph_key(item): item for item in left_empty}
    right_by_key = {_empty_paragraph_key(item): item for item in right_empty}
    added = [item for item in right_empty if _empty_paragraph_key(item) not in left_keys]
    removed = [item for item in left_empty if _empty_paragraph_key(item) not in right_keys]
    changed = []
    for key in sorted(left_keys & right_keys):
        left_item = left_by_key[key]
        right_item = right_by_key[key]
        property_changes = _dict_value_changes(left_item.get("properties", {}), right_item.get("properties", {}))
        if property_changes:
            changed.append(
                {
                    "kind": "empty-paragraph",
                    "paragraphIndex": right_item.get("paragraphIndex"),
                    "afterParagraphTextPreview": right_item.get("afterParagraphTextPreview", ""),
                    "beforeParagraphTextPreview": right_item.get("beforeParagraphTextPreview", ""),
                    "propertyChanges": property_changes,
                }
            )
    return {
        "available": True,
        "identical": not added and not removed and not changed,
        "summary": {
            "leftCount": len(left_empty),
            "rightCount": len(right_empty),
            "added": len(added),
            "removed": len(removed),
            "changed": len(changed),
        },
        "added": added,
        "removed": removed,
        "changed": changed,
    }


def _empty_paragraphs(docx_path: Path) -> list[dict[str, Any]]:
    entries = _read_docx_entry_bytes(docx_path)
    document_xml = entries.get("word/document.xml")
    if document_xml is None:
        return []
    root = ET.fromstring(document_xml)
    paragraphs = root.findall(".//w:p", NS)
    result: list[dict[str, Any]] = []
    for index, paragraph in enumerate(paragraphs, start=1):
        if not _is_empty_paragraph_xml(paragraph):
            continue
        previous_text = _paragraph_text(paragraphs[index - 2]) if index > 1 else ""
        next_text = _paragraph_text(paragraphs[index]) if index < len(paragraphs) else ""
        result.append(
            {
                "paragraphIndex": index,
                "kind": "empty-paragraph",
                "afterParagraphTextPreview": previous_text[:80],
                "beforeParagraphTextPreview": next_text[:80],
                "properties": _paragraph_spacing_properties_xml(paragraph),
            }
        )
    return result


def _manual_page_break_paragraphs(docx_path: Path) -> list[dict[str, Any]]:
    entries = _read_docx_entry_bytes(docx_path)
    document_xml = entries.get("word/document.xml")
    if document_xml is None:
        return []
    root = ET.fromstring(document_xml)
    paragraphs = root.findall(".//w:p", NS)
    result: list[dict[str, Any]] = []
    for index, paragraph in enumerate(paragraphs, start=1):
        if not _is_manual_page_break_paragraph_xml(paragraph):
            continue
        next_text = _paragraph_text(paragraphs[index]) if index < len(paragraphs) else ""
        result.append(
            {
                "paragraphIndex": index,
                "kind": "manual-page-break",
                "beforeParagraphTextPreview": next_text[:80],
            }
        )
    return result


def _manual_break_key(item: dict[str, Any]) -> tuple[str, str]:
    return (str(item.get("kind", "")), str(item.get("beforeParagraphTextPreview", "")))


def _empty_paragraph_key(item: dict[str, Any]) -> tuple[str, str, str]:
    return (
        str(item.get("kind", "")),
        str(item.get("afterParagraphTextPreview", "")),
        str(item.get("beforeParagraphTextPreview", "")),
    )


def _dict_value_changes(left: dict[str, Any], right: dict[str, Any]) -> dict[str, dict[str, Any]]:
    changes: dict[str, dict[str, Any]] = {}
    for key in sorted(set(left) | set(right)):
        if left.get(key) != right.get(key):
            changes[key] = {"left": left.get(key), "right": right.get(key)}
    return changes


def _is_empty_paragraph_xml(paragraph: ET.Element) -> bool:
    if paragraph.tag != _tag("p"):
        return False
    if _is_manual_page_break_paragraph_xml(paragraph):
        return False
    if any((text_node.text or "") for text_node in paragraph.findall(".//w:t", NS)):
        return False
    if paragraph.findall(".//w:drawing", NS):
        return False
    if paragraph.findall(".//w:br", NS):
        return False
    return True


def _paragraph_spacing_properties_xml(paragraph: ET.Element) -> dict[str, str]:
    props: dict[str, str] = {}
    p_pr = paragraph.find("w:pPr", NS)
    if p_pr is None:
        return props
    p_style = p_pr.find("w:pStyle", NS)
    if p_style is not None and p_style.attrib.get(_tag("val")):
        props["style-id"] = p_style.attrib[_tag("val")]
    spacing = p_pr.find("w:spacing", NS)
    if spacing is not None:
        before = spacing.attrib.get(_tag("before"))
        before_lines = spacing.attrib.get(_tag("beforeLines"))
        if before and before.isdigit():
            props["space-before"] = f"{int(before) / 20:g}pt"
        elif before_lines and before_lines.isdigit():
            props["space-before"] = f"{int(before_lines) / 100:g}line"
        after = spacing.attrib.get(_tag("after"))
        after_lines = spacing.attrib.get(_tag("afterLines"))
        if after and after.isdigit():
            props["space-after"] = f"{int(after) / 20:g}pt"
        elif after_lines and after_lines.isdigit():
            props["space-after"] = f"{int(after_lines) / 100:g}line"
        line = spacing.attrib.get(_tag("line"))
        line_rule = spacing.attrib.get(_tag("lineRule"))
        if line and line.isdigit():
            if line_rule in {None, "auto"}:
                props["line-spacing"] = f"{int(line) / 240:g}"
            elif line_rule in {"exact", "atLeast"}:
                props["line-spacing"] = f"{int(line) / 20:g}pt"
    return props


def _is_manual_page_break_paragraph_xml(paragraph: ET.Element) -> bool:
    meaningful_children = [child for child in list(paragraph) if child.tag != _tag("pPr")]
    if len(meaningful_children) != 1 or meaningful_children[0].tag != _tag("r"):
        return False
    run_children = list(meaningful_children[0])
    if len(run_children) != 1 or run_children[0].tag != _tag("br"):
        return False
    return run_children[0].attrib.get(_tag("type")) == "page"


def _paragraph_text(paragraph: ET.Element) -> str:
    chunks: list[str] = []
    for text_node in paragraph.findall(".//w:t", NS):
        chunks.append(text_node.text or "")
    for br in paragraph.findall(".//w:br", NS):
        if br.attrib.get(_tag("type")) == "page":
            chunks.append("[manual-page-break]")
        else:
            chunks.append("\n")
    return "".join(chunks)


def _package_diff_kind(package_diff: dict[str, Any], entry_name: str, collection: str) -> str | None:
    for entry in package_diff.get(collection, []):
        if entry.get("path") == entry_name:
            return entry.get("kind")
    return None


def _byte_change_region(left: bytes, right: bytes) -> dict[str, Any]:
    if left == right:
        return {
            "identical": True,
            "commonPrefixBytes": len(left),
            "commonSuffixBytes": 0,
            "left": _byte_fragment_snapshot(left, 0, 0),
            "right": _byte_fragment_snapshot(right, 0, 0),
        }
    prefix = 0
    max_prefix = min(len(left), len(right))
    while prefix < max_prefix and left[prefix] == right[prefix]:
        prefix += 1

    suffix = 0
    max_suffix = min(len(left) - prefix, len(right) - prefix)
    while suffix < max_suffix and left[len(left) - 1 - suffix] == right[len(right) - 1 - suffix]:
        suffix += 1

    left_end = len(left) - suffix
    right_end = len(right) - suffix
    return {
        "identical": False,
        "commonPrefixBytes": prefix,
        "commonSuffixBytes": suffix,
        "left": _byte_fragment_snapshot(left, prefix, left_end),
        "right": _byte_fragment_snapshot(right, prefix, right_end),
    }


def _one_sided_byte_region(data: bytes, side: str) -> dict[str, Any]:
    empty = b""
    return {
        "identical": False,
        "commonPrefixBytes": 0,
        "commonSuffixBytes": 0,
        "left": _byte_fragment_snapshot(data if side == "left" else empty, 0, len(data) if side == "left" else 0),
        "right": _byte_fragment_snapshot(data if side == "right" else empty, 0, len(data) if side == "right" else 0),
    }


def _byte_fragment_snapshot(data: bytes, start: int, end: int) -> dict[str, Any]:
    fragment = data[start:end]
    return {
        "size": len(data),
        "changedRange": {
            "start": start,
            "endExclusive": end,
            "length": len(fragment),
        },
        "changedSha256": sha256_bytes(fragment),
    }


def _semantic_links_by_entry(semantic_diff: dict[str, Any]) -> dict[str, list[dict[str, Any]]]:
    if not semantic_diff.get("available"):
        return {}
    links: dict[str, list[dict[str, Any]]] = {}
    for item in semantic_diff.get("changedNodes", []):
        _append_semantic_link(links, _fragment_node_link(item, "changed"))
    for item in semantic_diff.get("leftOnlyNodes", []):
        _append_semantic_link(links, _fragment_node_link(item, "left-only"))
    for item in semantic_diff.get("rightOnlyNodes", []):
        _append_semantic_link(links, _fragment_node_link(item, "right-only"))
    return links


def _append_semantic_link(links: dict[str, list[dict[str, Any]]], link: dict[str, Any]) -> None:
    entry_name = _part_path_to_entry_name(link.get("partPath"))
    if entry_name is None:
        return
    links.setdefault(entry_name, []).append(link)


def _fragment_node_link(item: dict[str, Any], status: str) -> dict[str, Any]:
    left = item.get("left") if isinstance(item.get("left"), dict) else {}
    right = item.get("right") if isinstance(item.get("right"), dict) else {}
    node = right or left or item
    link: dict[str, Any] = {
        "nodeId": item.get("nodeId") or node.get("nodeId"),
        "status": status,
        "kind": item.get("kind") or node.get("kind"),
        "partPath": item.get("partPath") or node.get("partPath"),
        "locator": node.get("locator", {}),
    }
    if item.get("changedFields"):
        link["changedFields"] = item["changedFields"]
    left_hash = left.get("hash")
    right_hash = right.get("hash")
    if left_hash is not None or right_hash is not None:
        link["hash"] = {"left": left_hash, "right": right_hash}
    left_text_hash = left.get("textHash")
    right_text_hash = right.get("textHash")
    if left_text_hash is not None or right_text_hash is not None:
        link["textHash"] = {"left": left_text_hash, "right": right_text_hash}
    return link


def _part_path_to_entry_name(part_path: Any) -> str | None:
    if not isinstance(part_path, str) or not part_path:
        return None
    return part_path[1:] if part_path.startswith("/") else part_path


def _read_docx_entry_bytes(docx_path: Path) -> dict[str, bytes]:
    entries: dict[str, bytes] = {}
    with zipfile.ZipFile(docx_path, "r") as zf:
        for info in zf.infolist():
            if info.is_dir():
                continue
            entries[info.filename] = zf.read(info.filename)
    return entries


def _entry_feature(entries: list[str], entry_name_or_prefix: str) -> dict[str, Any]:
    if entry_name_or_prefix.endswith("/"):
        matched = [entry for entry in entries if entry.startswith(entry_name_or_prefix)]
    else:
        matched = [entry for entry in entries if entry == entry_name_or_prefix]
    return _feature_payload(bool(matched), entries=matched)


def _entry_any_prefix_feature(entries: list[str], prefixes: list[str]) -> dict[str, Any]:
    matched = [entry for entry in entries if any(entry.startswith(prefix) for prefix in prefixes)]
    return _feature_payload(bool(matched), entries=matched)


def _entry_or_relationship_feature(
    entries: list[str],
    rel_entries: dict[str, bytes],
    prefixes: list[str],
    relationship_needles: list[str],
) -> dict[str, Any]:
    matched_entries = [entry for entry in entries if any(entry.startswith(prefix) for prefix in prefixes)]
    matched_relationships = _relationship_matches(rel_entries, relationship_needles)
    return _feature_payload(bool(matched_entries or matched_relationships), entries=matched_entries, relationships=matched_relationships)


def _entry_or_xml_feature(
    entries: list[str],
    xml_entries: dict[str, bytes],
    entry_names: list[str],
    local_names: list[str],
) -> dict[str, Any]:
    matched_entries = [entry for entry in entries if entry in entry_names]
    xml_feature = _xml_any_local_feature(xml_entries, local_names)
    payload = _feature_payload(bool(matched_entries or xml_feature["present"]), entries=matched_entries)
    payload["xmlCounts"] = xml_feature.get("counts", {})
    payload["xmlEntries"] = xml_feature.get("xmlEntries", [])
    return payload


def _xml_local_feature(xml_entries: dict[str, bytes], local_name: str) -> dict[str, Any]:
    return _xml_any_local_feature(xml_entries, [local_name])


def _xml_any_local_feature(xml_entries: dict[str, bytes], local_names: list[str]) -> dict[str, Any]:
    counts: dict[str, int] = {}
    entry_hits: dict[str, dict[str, int]] = {}
    for entry_name, data in xml_entries.items():
        text = data.decode("utf-8", errors="ignore")
        for local_name in local_names:
            count = _count_xml_local_name(text, local_name)
            if count:
                counts[local_name] = counts.get(local_name, 0) + count
                entry_hits.setdefault(entry_name, {})[local_name] = count
    return _feature_payload(bool(counts), counts=counts, xmlEntries=entry_hits)


def _count_xml_local_name(text: str, local_name: str) -> int:
    return len(re.findall(rf"<(?:[A-Za-z0-9_.-]+:)?{re.escape(local_name)}(?:\s|/|>)", text))


def _relationship_matches(rel_entries: dict[str, bytes], needles: list[str]) -> dict[str, list[str]]:
    matches: dict[str, list[str]] = {}
    for entry_name, data in rel_entries.items():
        text = data.decode("utf-8", errors="ignore")
        hits = []
        for rel_type in re.findall(r'\bType="([^"]+)"', text):
            if any(needle in rel_type for needle in needles):
                hits.append(rel_type)
        if hits:
            matches[entry_name] = sorted(set(hits))
    return matches


def _feature_payload(present: bool, **fields: Any) -> dict[str, Any]:
    payload = {"present": present}
    for key, value in fields.items():
        if value:
            payload[key] = value
    return payload


def _audit_policy(feature_name: str) -> str:
    policies = {
        "customXml": "preserve-package-entry",
        "chart": "preserve-protected-object",
        "smartArt": "preserve-protected-object",
        "ole": "preserve-protected-object",
        "alternateContent": "preserve-protected-markup",
        "vml": "preserve-protected-markup",
        "textBox": "preserve-protected-markup",
        "field": "project-as-protected",
        "equation": "project-as-protected-controlled-omml-replacement",
        "revision": "project-as-protected-controlled-action",
        "comment": "project-as-protected-controlled-comment-text",
        "footnote": "project-secondary-part",
        "endnote": "project-secondary-part",
        "headerFooter": "project-secondary-part",
        "image": "project-metadata-controlled-media",
    }
    return policies.get(feature_name, "audit-only")


NODE_DIFF_FIELDS = (
    "kind",
    "partPath",
    "parent",
    "locator",
    "lock",
    "styleId",
    "numbering",
    "text",
    "properties",
    "objectKind",
    "objectProperties",
    "protectedKind",
    "children",
)


def _node_semantic_changed(left: dict[str, Any], right: dict[str, Any]) -> bool:
    if left.get("hash") != right.get("hash"):
        return True
    return any(left.get(field) != right.get(field) for field in NODE_DIFF_FIELDS)


def _node_diff(node_id: str, left: dict[str, Any], right: dict[str, Any]) -> dict[str, Any]:
    changed_fields = [
        field
        for field in ("hash", *NODE_DIFF_FIELDS)
        if left.get(field) != right.get(field)
    ]
    result: dict[str, Any] = {
        "nodeId": node_id,
        "kind": left.get("kind", right.get("kind")),
        "partPath": left.get("partPath", right.get("partPath")),
        "changedFields": changed_fields,
        "left": _node_summary(node_id, left),
        "right": _node_summary(node_id, right),
    }
    changes: dict[str, dict[str, Any]] = {}
    for field in changed_fields:
        changes[field] = {"left": left.get(field), "right": right.get(field)}
    result["changes"] = changes
    return result


def _node_summary(node_id: str, node: dict[str, Any]) -> dict[str, Any]:
    summary: dict[str, Any] = {
        "nodeId": node_id,
        "kind": node.get("kind"),
        "partPath": node.get("partPath"),
        "parent": node.get("parent"),
        "locator": node.get("locator"),
        "lock": node.get("lock"),
        "hash": node.get("hash"),
    }
    for key in (
        "styleId",
        "numbering",
        "text",
        "textHash",
        "properties",
        "objectKind",
        "objectProperties",
        "protectedKind",
        "simpleEditable",
        "children",
    ):
        if key in node:
            summary[key] = node[key]
    return summary


def _node_sort_key(nodes: dict[str, Any], node_id: str) -> tuple[Any, ...]:
    node = nodes.get(node_id, {})
    locator = node.get("locator") or {}
    return (
        node.get("partPath", ""),
        locator.get("paragraphIndex", 0),
        locator.get("runInParagraphIndex", 0),
        node.get("kind", ""),
        node_id,
    )


def _load_manifest(bundle_dir: Path) -> dict[str, Any]:
    manifest_path = bundle_dir / MANIFEST_NAME
    if not manifest_path.exists():
        raise HDocxError("MANIFEST_MISSING", "H-DOCX manifest is missing.", {"path": str(manifest_path)})
    return read_json(manifest_path)


def _collect_docx_inputs(input_path: Path) -> list[Path]:
    if not input_path.exists():
        raise HDocxError("CHECK_INPUT_NOT_FOUND", "DOCX check input does not exist.", {"path": str(input_path)})
    if input_path.is_file():
        if input_path.suffix.lower() != ".docx":
            raise HDocxError("CHECK_INPUT_NOT_DOCX", "DOCX check input file must have a .docx suffix.", {"path": str(input_path)})
        if input_path.name.startswith("~$"):
            raise HDocxError("CHECK_INPUT_TEMP_DOCX", "DOCX check input appears to be a temporary Word lock file.", {"path": str(input_path)})
        return [input_path]
    if not input_path.is_dir():
        raise HDocxError("CHECK_INPUT_UNSUPPORTED", "DOCX check input must be a .docx file or directory.", {"path": str(input_path)})
    inputs = sorted(
        path
        for path in input_path.rglob("*.docx")
        if path.is_file() and not path.name.startswith("~$") and not _inside_hdocx_bundle(path)
    )
    if not inputs:
        raise HDocxError("CHECK_INPUT_EMPTY", "DOCX check directory does not contain any .docx files.", {"path": str(input_path)})
    return inputs


def _validate_batch_output_paths(input_path: Path, work_dir: Path, output_dir: Path) -> None:
    for label, generated_dir in (("work", work_dir), ("out", output_dir)):
        if generated_dir == input_path:
            raise HDocxError(
                "BATCH_CHECK_OUTPUT_OVERLAPS_INPUT",
                "Batch generated directory must not be the input path.",
                {"input": str(input_path), label: str(generated_dir)},
            )
        if _path_contains(generated_dir, input_path):
            raise HDocxError(
                "BATCH_CHECK_OUTPUT_CONTAINS_INPUT",
                "Batch generated directory must not contain the input path.",
                {"input": str(input_path), label: str(generated_dir)},
            )


def _validate_path_encodable(path: Path, argument: str) -> None:
    text = str(path)
    try:
        text.encode("utf-8")
    except UnicodeEncodeError as exc:
        raise HDocxError(
            "PATH_ENCODING_ERROR",
            "Path contains characters that cannot be encoded safely.",
            {"argument": argument, "pathRepr": repr(text)},
        ) from exc


def _path_contains(container: Path, target: Path) -> bool:
    try:
        target.relative_to(container)
        return True
    except ValueError:
        return False


def _inside_hdocx_bundle(path: Path) -> bool:
    return any(part.lower().endswith(".hdocx") for part in path.parts)


def _safe_output_stem(relative_path: str, index: int) -> str:
    normalized = relative_path.replace("\\", "/").removesuffix(".docx")
    safe = re.sub(r"[^A-Za-z0-9._-]+", "-", normalized.replace("/", "__")).strip("-._")
    if not safe:
        safe = "document"
    return f"{index:04d}-{safe}"


def _part_manifest(entries: list[dict[str, Any]]) -> dict[str, Any]:
    return {
        "/" + entry["path"]: {
            "sha256": entry["sha256"],
            "uncompressedSize": entry["uncompressedSize"],
            "zipOrder": entry["zipOrder"],
        }
        for entry in entries
        if not entry["path"].endswith("/")
    }


def _default_preview_css() -> str:
    return """/* Generated preview CSS. Do not use this file as the DOCX truth source. */
body {
  font-family: "Times New Roman", "SimSun", serif;
  line-height: 1.5;
}
.hdocx-p {
  margin: 0 0 0.75rem 0;
}
.hlock-protected {
  background: #fff4cc;
}
.hdocx-image-preview {
  display: inline-block;
  max-width: 100%;
  vertical-align: top;
}
"""


def _default_agent_hcss() -> str:
    return """/* H-CSS edit script.
   Supported declarations are listed in hdocx_guidance(topic="hcss").
   Run hdocx_plan before hdocx_apply; it reports selector matches,
   declaration support, OOXML mappings, and patch ids.
*/
"""


def _detect_modifications(bundle_dir: Path) -> list[str]:
    manifest = _load_manifest(bundle_dir)
    modified: list[str] = []
    for rel_path in [HTML_NAME, AGENT_HCSS_NAME]:
        path = bundle_dir / rel_path
        recorded = manifest.get("files", {}).get(rel_path, {}).get("sha256")
        if recorded and path.exists() and sha256_file(path) != recorded:
            modified.append(rel_path)
    return modified


def _compile_part_replacements(
    bundle_dir: Path,
    manifest: dict[str, Any],
    edit_index: int,
) -> tuple[list[dict[str, Any]], list[dict[str, Any]], int]:
    patches: list[dict[str, Any]] = []
    errors: list[dict[str, Any]] = []
    parts_root = bundle_dir / "parts"
    for part_path, recorded in sorted(manifest.get("parts", {}).items()):
        entry_name = part_path.lstrip("/")
        local_path = parts_root.joinpath(*entry_name.split("/"))
        if not local_path.exists():
            errors.append(
                {
                    "code": "BUNDLE_PART_FILE_MISSING",
                    "message": "An extracted package part is missing from the H-DOCX bundle.",
                    "partPath": part_path,
                }
            )
            continue
        if not local_path.is_file():
            errors.append(
                {
                    "code": "BUNDLE_PART_FILE_NOT_REGULAR",
                    "message": "An extracted package part path is not a regular file.",
                    "partPath": part_path,
                }
            )
            continue
        actual_hash = sha256_file(local_path)
        old_hash = recorded.get("sha256")
        if actual_hash == old_hash:
            continue
        if not entry_name.startswith("word/media/"):
            errors.append(
                {
                    "code": "UNSUPPORTED_PART_REPLACEMENT",
                    "message": "Only existing word/media/* package parts can be replaced in this phase.",
                    "partPath": part_path,
                    "expectedSha256": old_hash,
                    "actualSha256": actual_hash,
                }
            )
            continue
        edit_index += 1
        patches.append(
            {
                "id": f"patch-{edit_index:06d}",
                "operation": "replace-media",
                "partPath": part_path,
                "entryName": entry_name,
                "sourcePath": f"parts/{entry_name}",
                "expectedOldHash": old_hash,
                "newSha256": actual_hash,
                "newSize": local_path.stat().st_size,
            }
        )
    return patches, errors, edit_index


def _analyze_edits(bundle_dir: Path) -> dict[str, Any]:
    manifest = _load_manifest(bundle_dir)
    modified_files = _detect_modifications(bundle_dir)
    errors: list[dict[str, Any]] = []
    patches: list[dict[str, Any]] = []

    html_modified = HTML_NAME in modified_files
    html_text = (bundle_dir / HTML_NAME).read_text(encoding="utf-8")
    html_nodes = collect_hdocx_nodes(html_text)
    edit_index = 0

    part_patches, part_errors, edit_index = _compile_part_replacements(bundle_dir, manifest, edit_index)
    patches.extend(part_patches)
    errors.extend(part_errors)

    hcss_diagnostics = None
    if AGENT_HCSS_NAME in modified_files:
        hcss_patches, hcss_errors, edit_index, hcss_diagnostics = _compile_hcss_edits(
            bundle_dir,
            manifest,
            html_nodes,
            edit_index,
        )
        patches.extend(hcss_patches)
        errors.extend(hcss_errors)

    if not html_modified:
        return {
            "modified": bool(modified_files or part_patches or part_errors),
            "modifiedFiles": modified_files,
            "modifiedParts": [patch["partPath"] for patch in part_patches],
            "patches": patches,
            "errors": errors,
            "hcss": hcss_diagnostics,
        }

    manifest_nodes = manifest.get("nodes", {})
    html_patch_base = len(patches)
    html_error_base = len(errors)

    for node_id, node in manifest_nodes.items():
        if node.get("kind") == "paragraph":
            current = html_nodes.get(node_id)
            if current is None:
                errors.append(
                    {
                        "code": "HTML_PARAGRAPH_NODE_MISSING",
                        "message": "A manifest paragraph node is missing from document.html.",
                        "nodeId": node_id,
                    }
                )
                continue
            if current.kind != "paragraph":
                errors.append(
                    {
                        "code": "HTML_NODE_KIND_CHANGED",
                        "message": "A paragraph node changed its data-hdocx-type.",
                        "nodeId": node_id,
                        "expected": "paragraph",
                        "actual": current.kind,
                    }
                )
                continue
            if current.lock != node.get("lock"):
                errors.append(
                    {
                        "code": "HTML_NODE_LOCK_CHANGED",
                        "message": "A node changed its data-hdocx-lock value.",
                        "nodeId": node_id,
                        "expected": node.get("lock"),
                        "actual": current.lock,
                    }
                )
                continue
            metadata_errors = _diff_readonly_paragraph_metadata(node, current.attrs)
            errors.extend({"nodeId": node_id, **error} for error in metadata_errors)
            property_errors, changed_properties = _diff_paragraph_properties(node, current.attrs)
            errors.extend({"nodeId": node_id, **error} for error in property_errors)
            if changed_properties:
                edit_index += 1
                patches.append(
                    {
                        "id": f"patch-{edit_index:06d}",
                        "operation": "patch-paragraph",
                        "partPath": node["partPath"],
                        "entryName": node["partPath"].lstrip("/"),
                        "nodeId": node_id,
                        "locator": node["locator"],
                        "expectedOldHash": node["hash"],
                        "oldProperties": node.get("properties", {}),
                        "newProperties": changed_properties,
                    }
                )
            continue
        if node.get("kind") == "protected":
            current = html_nodes.get(node_id)
            if current is None:
                errors.append(
                    {
                        "code": "HTML_PROTECTED_NODE_MISSING",
                        "message": "A protected node is missing from document.html.",
                        "nodeId": node_id,
                    }
                )
                continue
            if current.kind != "protected":
                errors.append(
                    {
                        "code": "HTML_NODE_KIND_CHANGED",
                        "message": "A protected node changed its data-hdocx-type.",
                        "nodeId": node_id,
                        "expected": "protected",
                        "actual": current.kind,
                    }
                )
                continue
            if current.lock != "protected":
                errors.append(
                    {
                        "code": "PROTECTED_NODE_LOCK_CHANGED",
                        "message": "A protected node changed its data-hdocx-lock value.",
                        "nodeId": node_id,
                        "actual": current.lock,
                    }
                )
                continue
            if current.text != node.get("text", ""):
                errors.append(
                    {
                        "code": "PROTECTED_NODE_TEXT_MODIFIED",
                        "message": "Protected node text was modified.",
                        "nodeId": node_id,
                        "protectedKind": node.get("protectedKind"),
                    }
                )
            continue
        if node.get("kind") != "run":
            continue
        current = html_nodes.get(node_id)
        if current is None:
            errors.append(
                {
                    "code": "HTML_RUN_NODE_MISSING",
                    "message": "A manifest run node is missing from document.html.",
                    "nodeId": node_id,
                }
            )
            continue
        if current.kind != "run":
            errors.append(
                {
                    "code": "HTML_NODE_KIND_CHANGED",
                    "message": "A run node changed its data-hdocx-type.",
                    "nodeId": node_id,
                    "expected": "run",
                    "actual": current.kind,
                }
            )
            continue
        if current.lock != node.get("lock"):
            errors.append(
                {
                    "code": "HTML_NODE_LOCK_CHANGED",
                    "message": "A node changed its data-hdocx-lock value.",
                    "nodeId": node_id,
                    "expected": node.get("lock"),
                    "actual": current.lock,
                }
            )
            continue

        old_text = node.get("text", "")
        new_text = current.text
        if current.has_segments:
            split_errors, split_segments = _build_split_segments(node, current)
            errors.extend({"nodeId": node_id, **error} for error in split_errors)
            if node.get("lock") != "editable":
                errors.append(
                    {
                        "code": "PROTECTED_RUN_SPLIT_MODIFIED",
                        "message": "Protected run was split or segment-formatted.",
                        "nodeId": node_id,
                    }
                )
                continue
            if not node.get("simpleEditable"):
                errors.append(
                    {
                        "code": "RUN_SPLIT_UNSUPPORTED",
                        "message": "Only simple-editable text runs can be split in this phase.",
                        "nodeId": node_id,
                    }
                )
                continue
            if split_errors:
                continue
            edit_index += 1
            patches.append(
                {
                    "id": f"patch-{edit_index:06d}",
                    "operation": "split-run",
                    "partPath": node["partPath"],
                    "entryName": node["partPath"].lstrip("/"),
                    "nodeId": node_id,
                    "locator": node["locator"],
                    "expectedOldHash": node["hash"],
                    "splitSegments": split_segments,
                }
            )
            continue

        property_errors, changed_properties = _diff_run_properties(node, current.attrs)
        errors.extend({"nodeId": node_id, **error} for error in property_errors)
        drawing_errors, changed_drawing_properties = _diff_drawing_properties(node, current.attrs)
        errors.extend({"nodeId": node_id, **error} for error in drawing_errors)
        if old_text == new_text:
            text_changed = False
        else:
            text_changed = True
        if node.get("lock") == "editable-metadata":
            metadata_errors: list[dict[str, Any]] = []
            if text_changed:
                metadata_errors.append(
                    {
                        "code": "PROTECTED_RUN_TEXT_MODIFIED",
                        "message": "Metadata-only run text was modified.",
                        "nodeId": node_id,
                    }
                )
            if changed_properties:
                metadata_errors.append(
                    {
                        "code": "PROTECTED_RUN_FORMAT_MODIFIED",
                        "message": "Metadata-only run formatting was modified.",
                        "nodeId": node_id,
                    }
                )
            errors.extend(metadata_errors)
            if metadata_errors or drawing_errors:
                continue
            if changed_drawing_properties:
                edit_index += 1
                operation = (
                    "patch-drawing-alt"
                    if set(changed_drawing_properties) == {"alt"}
                    else "patch-drawing-properties"
                )
                patches.append(
                    {
                        "id": f"patch-{edit_index:06d}",
                        "operation": operation,
                        "partPath": node["partPath"],
                        "entryName": node["partPath"].lstrip("/"),
                        "nodeId": node_id,
                        "locator": node["locator"],
                        "expectedOldHash": node["hash"],
                        "oldProperties": node.get("objectProperties", {}),
                        "newProperties": changed_drawing_properties,
                    }
                )
            continue
        if node.get("lock") != "editable":
            if text_changed:
                errors.append(
                    {
                        "code": "PROTECTED_RUN_TEXT_MODIFIED",
                        "message": "Protected run text was modified.",
                        "nodeId": node_id,
                    }
                )
            if changed_properties:
                errors.append(
                    {
                        "code": "PROTECTED_RUN_FORMAT_MODIFIED",
                        "message": "Protected run formatting was modified.",
                        "nodeId": node_id,
                    }
                )
            continue
        if text_changed and not node.get("simpleEditable"):
            errors.append(
                {
                    "code": "RUN_TEXT_PATCH_UNSUPPORTED",
                    "message": "This run is not simple-editable in the current patch compiler.",
                    "nodeId": node_id,
                }
            )
            continue
        if not text_changed and not changed_properties:
            continue
        edit_index += 1
        patch = {
            "id": f"patch-{edit_index:06d}",
            "operation": "patch-run",
            "partPath": node["partPath"],
            "entryName": node["partPath"].lstrip("/"),
            "nodeId": node_id,
            "locator": node["locator"],
            "expectedOldHash": node["hash"],
        }
        if text_changed:
            patch["oldText"] = old_text
            patch["newText"] = new_text
        if changed_properties:
            patch["oldProperties"] = node.get("properties", {})
            patch["newProperties"] = changed_properties
        patches.append(patch)

    if html_modified and len(patches) == html_patch_base and len(errors) == html_error_base:
        errors.append(
            {
                "code": "HTML_UNSUPPORTED_MODIFICATION",
                "message": "document.html changed, but no supported run text edits were detected.",
            }
        )

    return {
        "modified": bool(modified_files or part_patches or part_errors),
        "modifiedFiles": modified_files,
        "modifiedParts": [patch["partPath"] for patch in part_patches],
        "patches": patches,
        "errors": errors,
        "hcss": hcss_diagnostics,
    }


def _diff_run_properties(
    manifest_node: dict[str, Any],
    html_attrs: dict[str, str],
) -> tuple[list[dict[str, Any]], dict[str, str | None]]:
    errors: list[dict[str, Any]] = []
    changed: dict[str, str | None] = {}
    old_properties = manifest_node.get("properties", {})
    for prop_name, attr_name in RUN_PROPERTY_ATTRS.items():
        old_value = old_properties.get(prop_name)
        raw_new = html_attrs.get(attr_name)
        try:
            new_value = _normalize_run_property(prop_name, raw_new)
        except ValueError as exc:
            errors.append(
                {
                    "code": "HTML_INVALID_RUN_PROPERTY",
                    "message": str(exc),
                    "property": prop_name,
                    "value": raw_new,
                }
            )
            continue
        if old_value != new_value:
            changed[prop_name] = new_value
    return errors, changed


def _diff_readonly_paragraph_metadata(
    manifest_node: dict[str, Any],
    html_attrs: dict[str, str],
) -> list[dict[str, Any]]:
    errors: list[dict[str, Any]] = []
    checks: list[tuple[str, str, str | None]] = [
        ("styleId", "data-hdocx-style-id", manifest_node.get("styleId")),
    ]
    numbering = manifest_node.get("numbering", {})
    for prop_name, attr_name in PARAGRAPH_NUMBERING_ATTRS.items():
        checks.append((f"numbering.{prop_name}", attr_name, numbering.get(prop_name)))
    for prop_name, attr_name, old_value in checks:
        raw_new = html_attrs.get(attr_name)
        if old_value is None:
            if raw_new not in {None, ""}:
                errors.append(
                    {
                        "code": "HTML_READONLY_METADATA_MODIFIED",
                        "message": "Read-only paragraph metadata was added or modified.",
                        "property": prop_name,
                        "expected": None,
                        "actual": raw_new,
                    }
                )
            continue
        expected = str(old_value)
        if raw_new != expected:
            errors.append(
                {
                    "code": "HTML_READONLY_METADATA_MODIFIED",
                    "message": "Read-only paragraph metadata was modified.",
                    "property": prop_name,
                    "expected": expected,
                    "actual": raw_new,
                }
            )
    return errors


def _diff_drawing_properties(
    manifest_node: dict[str, Any],
    html_attrs: dict[str, str],
) -> tuple[list[dict[str, Any]], dict[str, str | None]]:
    errors: list[dict[str, Any]] = []
    changed: dict[str, str | None] = {}
    if manifest_node.get("objectKind") != "drawing":
        return errors, changed
    old_properties = manifest_node.get("objectProperties", {})
    for prop_name, attr_name in DRAWING_PROPERTY_ATTRS.items():
        old_value = old_properties.get(prop_name)
        raw_new = html_attrs.get(attr_name)
        try:
            new_value = _normalize_drawing_property(prop_name, raw_new, old_value)
        except ValueError as exc:
            errors.append(
                {
                    "code": "HTML_INVALID_DRAWING_PROPERTY",
                    "message": str(exc),
                    "property": prop_name,
                    "value": raw_new,
                }
            )
            continue
        if old_value != new_value:
            changed[prop_name] = new_value
    return errors, changed


def _normalize_drawing_property(prop_name: str, raw_value: str | None, old_value: str | None) -> str | None:
    if prop_name == "alt":
        return raw_value if raw_value not in {None, ""} else None
    if prop_name in {"width-emu", "height-emu"}:
        if raw_value is None or raw_value == "":
            if old_value is None:
                return None
            raise ValueError(f"{prop_name} cannot be removed; set a positive EMU integer.")
        value = raw_value.strip()
        if not value.isdigit():
            raise ValueError(f"{prop_name} must be a positive EMU integer.")
        normalized = str(int(value))
        if int(normalized) <= 0:
            raise ValueError(f"{prop_name} must be greater than zero.")
        if old_value is None:
            raise ValueError(f"{prop_name} cannot be added when the drawing has no wp:extent.")
        return normalized
    raise ValueError(f"Unsupported drawing property: {prop_name}")


def _canonical_image_property_name(prop_name: str) -> str:
    return prop_name


def _canonical_image_paragraph_property_name(prop_name: str) -> str:
    if not prop_name.startswith(IMAGE_PARAGRAPH_PROPERTY_PREFIX):
        return prop_name
    raw = prop_name.removeprefix(IMAGE_PARAGRAPH_PROPERTY_PREFIX)
    return _canonical_paragraph_property_name(raw)


def _normalize_image_hcss_property(prop_name: str, raw_value: str | None, old_value: str | None) -> str | None:
    value = None if raw_value is None else _strip_hcss_value(raw_value.strip())
    return _normalize_drawing_property(prop_name, value, old_value)


def _image_ooxml_mapping(prop_name: str) -> str:
    if prop_name == "alt":
        return "wp:docPr @descr"
    if prop_name in {"width-emu", "height-emu"}:
        return "wp:extent @cx/@cy"
    return OOXML_PROPERTY_MAP.get(prop_name, "drawing metadata")


def _diff_paragraph_properties(
    manifest_node: dict[str, Any],
    html_attrs: dict[str, str],
) -> tuple[list[dict[str, Any]], dict[str, str | None]]:
    errors: list[dict[str, Any]] = []
    changed: dict[str, str | None] = {}
    old_properties = manifest_node.get("properties", {})
    for prop_name, attr_name in PARAGRAPH_PROPERTY_ATTRS.items():
        old_value = old_properties.get(prop_name)
        raw_new = html_attrs.get(attr_name)
        try:
            new_value = _normalize_paragraph_property(prop_name, raw_new)
        except ValueError as exc:
            errors.append(
                {
                    "code": "HTML_INVALID_PARAGRAPH_PROPERTY",
                    "message": str(exc),
                    "property": prop_name,
                    "value": raw_new,
                }
            )
            continue
        if old_value != new_value:
            changed[prop_name] = new_value
    return errors, changed


def _build_split_segments(
    manifest_node: dict[str, Any],
    current: Any,
) -> tuple[list[dict[str, Any]], list[dict[str, Any]]]:
    errors: list[dict[str, Any]] = []
    old_text = manifest_node.get("text", "")
    if current.text != old_text:
        errors.append(
            {
                "code": "RUN_SPLIT_TEXT_CHANGED_UNSUPPORTED",
                "message": "Run split currently requires concatenated segment text to equal the original run text.",
                "oldText": old_text,
                "newText": current.text,
            }
        )
    segments: list[dict[str, Any]] = []
    for index, chunk in enumerate(current.chunks, start=1):
        if chunk.text == "":
            continue
        property_errors, properties = _extract_run_property_overrides(chunk.attrs)
        errors.extend(property_errors)
        segment: dict[str, Any] = {"index": index, "text": chunk.text}
        if properties:
            segment["properties"] = properties
        segments.append(segment)
    if len(segments) < 2:
        errors.append(
            {
                "code": "RUN_SPLIT_REQUIRES_MULTIPLE_SEGMENTS",
                "message": "Run split requires at least two non-empty text segments.",
            }
        )
    return errors, segments


def _extract_run_property_overrides(html_attrs: dict[str, str]) -> tuple[list[dict[str, Any]], dict[str, str | None]]:
    errors: list[dict[str, Any]] = []
    properties: dict[str, str | None] = {}
    for prop_name, attr_name in RUN_PROPERTY_ATTRS.items():
        if attr_name not in html_attrs:
            continue
        raw_value = html_attrs.get(attr_name)
        try:
            properties[prop_name] = _normalize_run_property(prop_name, raw_value)
        except ValueError as exc:
            errors.append(
                {
                    "code": "HTML_INVALID_RUN_PROPERTY",
                    "message": str(exc),
                    "property": prop_name,
                    "value": raw_value,
                }
            )
    return errors, properties


def _canonical_run_property_name(prop_name: str) -> str:
    return RUN_PROPERTY_ALIASES.get(prop_name, RUN_PROPERTY_ALIASES.get(prop_name.lower(), prop_name))


def _canonical_paragraph_property_name(prop_name: str) -> str:
    return PARAGRAPH_PROPERTY_ALIASES.get(
        prop_name,
        PARAGRAPH_PROPERTY_ALIASES.get(prop_name.lower(), prop_name),
    )


def _canonical_empty_paragraph_property_name(prop_name: str) -> str:
    return EMPTY_PARAGRAPH_PROPERTY_ALIASES.get(
        prop_name,
        EMPTY_PARAGRAPH_PROPERTY_ALIASES.get(prop_name.lower(), prop_name),
    )


def _canonical_hcss_property_name(prop_name: str) -> str:
    paragraph_name = _canonical_paragraph_property_name(prop_name)
    if paragraph_name in PARAGRAPH_PROPERTY_ATTRS:
        return paragraph_name
    run_name = _canonical_run_property_name(prop_name)
    if run_name in RUN_PROPERTY_ATTRS:
        return run_name
    return prop_name


def _normalize_run_property(prop_name: str, raw_value: str | None) -> str | None:
    prop_name = _canonical_run_property_name(prop_name)
    if raw_value is None or raw_value == "":
        return None
    value = _strip_hcss_value(raw_value.strip())
    if prop_name in {"bold", "italic"}:
        lowered = value.lower()
        if lowered in {"true", "1", "on", "yes"}:
            return "true"
        if lowered in {"false", "0", "off", "no"}:
            return "false"
        raise ValueError(f"{prop_name} must be true or false.")
    if prop_name == "font-size":
        if not value.endswith("pt"):
            raise ValueError("font-size must use pt units.")
        try:
            points = float(value[:-2])
        except ValueError as exc:
            raise ValueError("font-size must be numeric pt.") from exc
        if points <= 0:
            raise ValueError("font-size must be positive.")
        return f"{points:g}pt"
    if prop_name == "color":
        normalized = value[1:] if value.startswith("#") else value
        if len(normalized) != 6 or any(ch not in "0123456789abcdefABCDEF" for ch in normalized):
            raise ValueError("color must be #RRGGBB.")
        return f"#{normalized.lower()}"
    if prop_name in {"font-family", "ascii-font", "hansi-font", "east-asia-font", "cs-font"}:
        if not value:
            raise ValueError(f"{prop_name} must not be empty.")
        if any(ord(ch) < 32 for ch in value):
            raise ValueError(f"{prop_name} must not contain control characters.")
        return value
    raise ValueError(f"Unsupported run property: {prop_name}")


def _normalize_paragraph_property(prop_name: str, raw_value: str | None) -> str | None:
    prop_name = _canonical_paragraph_property_name(prop_name)
    if raw_value is None or raw_value == "":
        return None
    value = _strip_hcss_value(raw_value.strip())
    if prop_name == "align":
        aliases = {"justify": "justify", "both": "justify", "left": "left", "right": "right", "center": "center"}
        normalized = aliases.get(value.lower())
        if not normalized:
            raise ValueError("align must be left, center, right, justify, or both.")
        return normalized
    if prop_name == "first-line-indent":
        if value.endswith("char"):
            number = value[:-4]
            try:
                chars = float(number)
            except ValueError as exc:
                raise ValueError("first-line-indent char value must be numeric.") from exc
            if chars < 0:
                raise ValueError("first-line-indent must not be negative.")
            return f"{chars:g}char"
        if value.endswith("pt"):
            number = value[:-2]
            try:
                points = float(number)
            except ValueError as exc:
                raise ValueError("first-line-indent pt value must be numeric.") from exc
            if points < 0:
                raise ValueError("first-line-indent must not be negative.")
            return f"{points:g}pt"
        raise ValueError("first-line-indent must use char or pt units.")
    if prop_name == "line-spacing":
        if value.endswith("pt"):
            try:
                points = float(value[:-2])
            except ValueError as exc:
                raise ValueError("line-spacing pt value must be numeric.") from exc
            if points <= 0:
                raise ValueError("line-spacing must be positive.")
            return f"{points:g}pt"
        try:
            multiple = float(value)
        except ValueError as exc:
            raise ValueError("line-spacing must be a positive multiple or pt value.") from exc
        if multiple <= 0:
            raise ValueError("line-spacing must be positive.")
        return f"{multiple:g}"
    if prop_name in {"space-before", "space-after"}:
        if value == "0":
            return "0pt"
        if value.endswith("pt"):
            try:
                points = float(value[:-2])
            except ValueError as exc:
                raise ValueError(f"{prop_name} pt value must be numeric.") from exc
            if points < 0:
                raise ValueError(f"{prop_name} must not be negative.")
            return f"{points:g}pt"
        if value.endswith("line"):
            number = value[:-4]
            try:
                lines = float(number)
            except ValueError as exc:
                raise ValueError(f"{prop_name} line value must be numeric.") from exc
            if lines < 0:
                raise ValueError(f"{prop_name} must not be negative.")
            return f"{lines:g}line"
        raise ValueError(f"{prop_name} must use 0, pt, or line units.")
    raise ValueError(f"Unsupported paragraph property: {prop_name}")


def _normalize_empty_paragraph_property(prop_name: str, raw_value: str | None) -> str:
    raw_prop_name = prop_name
    prop_name = _canonical_empty_paragraph_property_name(prop_name)
    if prop_name not in EMPTY_PARAGRAPH_SUPPORTED_PROPERTIES:
        raise ValueError(f"Unsupported empty paragraph property: {prop_name}")
    if prop_name == "style-id":
        value = _strip_hcss_value((raw_value or "").strip())
        if not value:
            raise ValueError("empty-paragraph-style-id must not be empty.")
        if not re.fullmatch(r"[A-Za-z0-9_.-]+", value):
            raise ValueError("empty-paragraph-style-id must be a simple style id.")
        return value
    paragraph_prop = _canonical_paragraph_property_name(prop_name)
    if paragraph_prop not in {"line-spacing", "space-before", "space-after"}:
        raise ValueError(f"Unsupported empty paragraph property: {prop_name}")
    if raw_prop_name in {"line-spacing-exact", "line-spacingExact"} and not _strip_hcss_value(raw_value or "").endswith("pt"):
        raise ValueError("empty-paragraph-line-spacing-exact must use pt units.")
    normalized = _normalize_paragraph_property(paragraph_prop, raw_value)
    if normalized is None:
        raise ValueError(f"empty-paragraph-{prop_name} must not be empty.")
    return normalized


def _normalize_hcss_bool(prop_name: str, raw_value: str | None) -> str:
    if raw_value is None or raw_value == "":
        raise ValueError(f"{prop_name} must be true or false.")
    value = _strip_hcss_value(raw_value.strip()).lower()
    if value in {"true", "1", "on", "yes"}:
        return "true"
    if value in {"false", "0", "off", "no"}:
        return "false"
    raise ValueError(f"{prop_name} must be true or false.")


def _compile_hcss_edits(
    bundle_dir: Path,
    manifest: dict[str, Any],
    html_nodes: dict[str, Any],
    edit_index: int,
) -> tuple[list[dict[str, Any]], list[dict[str, Any]], int, dict[str, Any]]:
    text = (bundle_dir / AGENT_HCSS_NAME).read_text(encoding="utf-8-sig")
    program, parse_errors = _parse_hcss(text)
    diagnostics = _hcss_diagnostics_base(modified=True)
    if parse_errors:
        diagnostics["parseErrors"] = parse_errors
        diagnostics["summary"]["errorCount"] = len(parse_errors)
        return [], parse_errors, edit_index, diagnostics

    patches: list[dict[str, Any]] = []
    errors: list[dict[str, Any]] = []
    manifest_nodes = manifest.get("nodes", {})

    style_errors, style_patches, created_styles, edit_index = _compile_hcss_style_creations(
        bundle_dir,
        program["styleCreations"],
        manifest,
        patches,
        edit_index,
    )
    errors.extend(style_errors)
    patches.extend(style_patches)
    if created_styles:
        manifest = {**manifest, "styles": {**manifest.get("styles", {}), **created_styles}}
    style_delete_errors, style_delete_patches, edit_index = _compile_hcss_style_deletions(
        program["styleDeletions"],
        manifest,
        edit_index,
    )
    errors.extend(style_delete_errors)
    patches.extend(style_delete_patches)
    list_errors, list_patches, created_lists, edit_index = _compile_hcss_list_creations(
        bundle_dir,
        program["listCreations"],
        manifest,
        patches,
        edit_index,
    )
    errors.extend(list_errors)
    patches.extend(list_patches)
    if created_lists:
        manifest = {**manifest, "createdLists": created_lists}

    image_errors, image_patches, edit_index = _compile_hcss_image_insertions(
        bundle_dir,
        program["imageInsertions"],
        program["sets"],
        html_nodes,
        manifest,
        patches,
        edit_index,
    )
    errors.extend(image_errors)
    patches.extend(image_patches)
    table_errors, table_patches, edit_index = _compile_hcss_table_row_insertions(
        program["tableRowInsertions"],
        program["sets"],
        html_nodes,
        manifest,
        edit_index,
    )
    errors.extend(table_errors)
    patches.extend(table_patches)
    table_delete_errors, table_delete_patches, edit_index = _compile_hcss_table_row_deletions(
        program["tableRowDeletions"],
        program["sets"],
        html_nodes,
        manifest,
        edit_index,
    )
    errors.extend(table_delete_errors)
    patches.extend(table_delete_patches)
    col_insert_errors, col_insert_patches, edit_index = _compile_hcss_table_column_insertions(
        program["tableColumnInsertions"],
        program["sets"],
        html_nodes,
        manifest,
        edit_index,
    )
    errors.extend(col_insert_errors)
    patches.extend(col_insert_patches)
    col_delete_errors, col_delete_patches, edit_index = _compile_hcss_table_column_deletions(
        program["tableColumnDeletions"],
        program["sets"],
        html_nodes,
        manifest,
        edit_index,
    )
    errors.extend(col_delete_errors)
    patches.extend(col_delete_patches)

    for rule in program["rules"]:
        target_ids = _resolve_hcss_target(rule["target"], program["sets"], html_nodes)
        allow_empty = _hcss_target_allows_empty(rule["target"], program["sets"])
        rule_diag = {
            "line": rule.get("line"),
            "target": rule["target"],
            "mode": rule["mode"],
            "allowEmpty": allow_empty,
            "matchedNodeIds": target_ids,
            "matchCount": len(target_ids),
            "declarations": [],
            "patchIds": [],
            "errors": [],
        }
        diagnostics["rules"].append(rule_diag)
        if not target_ids:
            if allow_empty:
                _update_hcss_diagnostics_summary(diagnostics)
                continue
            error = (
                {
                    "code": "HCSS_SELECTOR_ZERO_MATCH",
                    "message": "H-CSS rule matched no H-DOCX nodes.",
                    "target": rule["target"],
                    "line": rule.get("line"),
                }
            )
            errors.append(error)
            rule_diag["errors"].append(error)
            _update_hcss_diagnostics_summary(diagnostics)
            continue
        expanded = _expand_hcss_declarations(rule["declarations"], program["formats"], program["tokens"])
        if expanded["errors"]:
            _annotate_hcss_errors(expanded["errors"], rule.get("declarationMeta", []), rule.get("line"))
            errors.extend(expanded["errors"])
            rule_diag["errors"].extend(expanded["errors"])
            rule_diag["declarations"] = _hcss_declaration_diagnostics(
                rule["mode"],
                rule["declarations"],
                rule.get("declarationMeta", []),
            )
            _update_hcss_diagnostics_summary(diagnostics)
            continue
        declarations = expanded["declarations"]
        rule_diag["declarations"] = _hcss_declaration_diagnostics(
            rule["mode"],
            declarations,
            rule.get("declarationMeta", []),
        )
        mode = rule["mode"]
        patch_base = len(patches)
        if mode == "paragraph-formatting":
            new_errors, new_patches, edit_index = _compile_hcss_paragraph_rule(
                target_ids, declarations, manifest_nodes, edit_index
            )
        elif mode == "image-formatting":
            new_errors, new_patches, edit_index = _compile_hcss_image_formatting_rule(
                target_ids, declarations, manifest_nodes, edit_index
            )
        elif mode == "paragraph-structure":
            new_errors, new_patches, edit_index = _compile_hcss_paragraph_structure_rule(
                bundle_dir, target_ids, declarations, manifest_nodes, edit_index
            )
        elif mode == "all-runs":
            new_errors, new_patches, edit_index = _compile_hcss_all_runs_rule(
                target_ids, declarations, manifest_nodes, edit_index
            )
        elif mode == "style-definition":
            new_errors, new_patches, edit_index = _compile_hcss_style_definition_rule(
                target_ids, declarations, manifest_nodes, manifest, edit_index
            )
        elif mode == "numbering-definition":
            new_errors, new_patches, edit_index = _compile_hcss_numbering_definition_rule(
                target_ids, declarations, manifest_nodes, manifest, edit_index
            )
        elif mode == "paragraph-style":
            new_errors, new_patches, edit_index = _compile_hcss_paragraph_style_rule(
                target_ids, declarations, manifest_nodes, manifest, edit_index
            )
        elif mode == "paragraph-numbering":
            new_errors, new_patches, edit_index = _compile_hcss_paragraph_numbering_rule(
                target_ids, declarations, manifest_nodes, manifest, edit_index
            )
        elif mode == "comment-text":
            new_errors, new_patches, edit_index = _compile_hcss_comment_text_rule(
                target_ids, declarations, manifest_nodes, edit_index
            )
        elif mode == "revision-action":
            new_errors, new_patches, edit_index = _compile_hcss_revision_action_rule(
                target_ids, declarations, manifest_nodes, edit_index
            )
        elif mode == "equation-omml":
            new_errors, new_patches, edit_index = _compile_hcss_equation_omml_rule(
                bundle_dir, target_ids, declarations, manifest_nodes, edit_index
            )
        elif mode == "direct-formatting":
            new_errors, new_patches, edit_index = _compile_hcss_direct_rule(
                target_ids, declarations, manifest_nodes, edit_index
            )
        else:
            new_errors = [
                {
                    "code": "HCSS_UNSUPPORTED_EDIT_MODE",
                    "message": "Unsupported @hdocx-edit mode.",
                    "mode": mode,
                }
            ]
            new_patches = []
        _annotate_hcss_errors(new_errors, rule_diag["declarations"], rule.get("line"))
        noops = [patch for patch in new_patches if patch.get("noop")]
        real_patches = [patch for patch in new_patches if not patch.get("noop")]
        if noops:
            rule_diag["noops"] = noops
        errors.extend(new_errors)
        patches.extend(real_patches)
        rule_diag["errors"].extend(new_errors)
        rule_diag["patchIds"] = [patch["id"] for patch in patches[patch_base:]]
        _update_hcss_diagnostics_summary(diagnostics)

    _update_hcss_diagnostics_summary(diagnostics)
    return patches, errors, edit_index, diagnostics


def _parse_hcss(text: str) -> tuple[dict[str, Any], list[dict[str, Any]]]:
    text = text.lstrip("\ufeff")
    stripped = _strip_hcss_comments_preserve_lines(text)
    tokens: dict[str, str] = {}
    sets: dict[str, dict[str, Any]] = {}
    formats: dict[str, list[tuple[str, str]]] = {}
    rules: list[dict[str, Any]] = []
    image_insertions: list[dict[str, Any]] = []
    table_row_insertions: list[dict[str, Any]] = []
    table_row_deletions: list[dict[str, Any]] = []
    table_column_insertions: list[dict[str, Any]] = []
    table_column_deletions: list[dict[str, Any]] = []
    style_creations: list[dict[str, Any]] = []
    style_deletions: list[dict[str, Any]] = []
    list_creations: list[dict[str, Any]] = []
    errors: list[dict[str, Any]] = []
    mode = "direct-formatting"

    stripped, image_insertions = _parse_hcss_image_insertion_blocks(stripped)
    stripped, table_row_insertions = _parse_hcss_table_row_insertion_blocks(stripped)
    stripped, table_row_deletions = _parse_hcss_table_row_deletion_rules(stripped)
    stripped, table_column_insertions = _parse_hcss_table_column_insertion_blocks(stripped)
    stripped, table_column_deletions = _parse_hcss_table_column_deletion_rules(stripped)
    stripped, style_creations = _parse_hcss_style_blocks(stripped)
    stripped, style_deletions = _parse_hcss_style_deletion_rules(stripped)
    stripped, list_creations = _parse_hcss_list_blocks(stripped)

    pattern = re.compile(
        r"@hdocx-token\s+([A-Za-z0-9_.-]+)\s+([^;]+);"
        r"|@hdocx-edit\s+mode\(([^)]+)\)\s*;"
        r"|\s*(?:(@hdocx-set|@hdocx-format)\s+)?([^{};]+?)\s*\{([^{}]*)\}",
        flags=re.DOTALL,
    )

    for match in pattern.finditer(stripped):
        token_name, token_value, edit_mode, block_kind, block_name, block_body = match.groups()
        if token_name:
            tokens[token_name] = _strip_hcss_value(token_value)
            continue
        if edit_mode:
            mode = edit_mode.strip()
            continue
        if not block_name:
            continue
        name = block_name.strip()
        rule_line = stripped[: match.start()].count("\n") + 1
        body_line = stripped[: match.start(6)].count("\n") + 1
        declarations, declaration_meta = _parse_hcss_declarations_with_meta(block_body, body_line)
        if block_kind == "@hdocx-set":
            sets[name] = _parse_hcss_set(name, declarations)
        elif block_kind == "@hdocx-format":
            formats[name] = [(key, value) for key, value in declarations if key != "@hdocx-include"]
        else:
            rules.append(
                {
                    "target": name,
                    "mode": mode,
                    "declarations": declarations,
                    "line": rule_line,
                    "declarationMeta": declaration_meta,
                }
            )

    remainder = pattern.sub("", stripped).strip()
    if remainder:
        errors.append(
            {
                "code": "HCSS_PARSE_UNSUPPORTED_SYNTAX",
                "message": "H-CSS contains unsupported syntax.",
                "text": remainder[:120],
            }
        )
    return {
        "tokens": tokens,
        "sets": sets,
        "formats": formats,
        "rules": rules,
        "imageInsertions": image_insertions,
        "tableRowInsertions": table_row_insertions,
        "tableRowDeletions": table_row_deletions,
        "tableColumnInsertions": table_column_insertions,
        "tableColumnDeletions": table_column_deletions,
        "styleCreations": style_creations,
        "styleDeletions": style_deletions,
        "listCreations": list_creations,
    }, errors


def _strip_hcss_comments_preserve_lines(text: str) -> str:
    def replace(match: re.Match[str]) -> str:
        comment = match.group(0)
        return "\n" * comment.count("\n")

    return re.sub(r"/\*.*?\*/", replace, text, flags=re.DOTALL)


def _parse_hcss_image_insertion_blocks(text: str) -> tuple[str, list[dict[str, Any]]]:
    insertions: list[dict[str, Any]] = []
    pattern = re.compile(
        r"@hdocx-insert-image\s+(after|before)\(([^)]+)\)\s*\{([^{}]*)\}",
        flags=re.DOTALL,
    )
    for match in pattern.finditer(text):
        position, target, body = match.groups()
        insertions.append(
            {
                "position": position.strip(),
                "target": target.strip(),
                "declarations": _parse_hcss_declarations(body),
            }
        )
    return pattern.sub("", text), insertions


def _parse_hcss_table_row_insertion_blocks(text: str) -> tuple[str, list[dict[str, Any]]]:
    insertions: list[dict[str, Any]] = []
    pattern = re.compile(
        r"@hdocx-insert-table-row\s+(after|before)\(([^)]+)\)\s*\{([^{}]*)\}",
        flags=re.DOTALL,
    )
    for match in pattern.finditer(text):
        position, target, body = match.groups()
        insertions.append(
            {
                "position": position.strip(),
                "target": target.strip(),
                "declarations": _parse_hcss_declarations(body),
            }
        )
    return pattern.sub("", text), insertions


def _parse_hcss_table_row_deletion_rules(text: str) -> tuple[str, list[dict[str, Any]]]:
    deletions: list[dict[str, Any]] = []
    pattern = re.compile(r"@hdocx-delete-table-row\s*\(([^)]+)\)\s*;", flags=re.DOTALL)
    for match in pattern.finditer(text):
        deletions.append({"target": match.group(1).strip()})
    return pattern.sub("", text), deletions


def _parse_hcss_table_column_insertion_blocks(text: str) -> tuple[str, list[dict[str, Any]]]:
    insertions: list[dict[str, Any]] = []
    pattern = re.compile(
        r"@hdocx-insert-table-column\s+(after|before)\(([^)]+)\)\s*\{([^{}]*)\}",
        flags=re.DOTALL,
    )
    for match in pattern.finditer(text):
        position, target, body = match.groups()
        insertions.append(
            {
                "position": position.strip(),
                "target": target.strip(),
                "declarations": _parse_hcss_declarations(body),
            }
        )
    return pattern.sub("", text), insertions


def _parse_hcss_table_column_deletion_rules(text: str) -> tuple[str, list[dict[str, Any]]]:
    deletions: list[dict[str, Any]] = []
    pattern = re.compile(r"@hdocx-delete-table-column\s*\(([^)]+)\)\s*;", flags=re.DOTALL)
    for match in pattern.finditer(text):
        deletions.append({"target": match.group(1).strip()})
    return pattern.sub("", text), deletions


def _parse_hcss_style_blocks(text: str) -> tuple[str, list[dict[str, Any]]]:
    styles: list[dict[str, Any]] = []
    pattern = re.compile(r"@hdocx-style\s+([A-Za-z0-9_.-]+)\s*\{([^{}]*)\}", flags=re.DOTALL)
    for match in pattern.finditer(text):
        style_id, body = match.groups()
        styles.append({"styleId": style_id.strip(), "declarations": _parse_hcss_declarations(body)})
    return pattern.sub("", text), styles


def _parse_hcss_style_deletion_rules(text: str) -> tuple[str, list[dict[str, Any]]]:
    deletions: list[dict[str, Any]] = []
    pattern = re.compile(r"@hdocx-delete-style\s*\(([^)]+)\)\s*;", flags=re.DOTALL)
    for match in pattern.finditer(text):
        deletions.append({"styleId": match.group(1).strip()})
    return pattern.sub("", text), deletions


def _parse_hcss_list_blocks(text: str) -> tuple[str, list[dict[str, Any]]]:
    lists: list[dict[str, Any]] = []
    pattern = re.compile(r"@hdocx-list\s+([A-Za-z0-9_.-]+)\s*\{([^{}]*)\}", flags=re.DOTALL)
    for match in pattern.finditer(text):
        list_id, body = match.groups()
        lists.append({"listId": list_id.strip(), "declarations": _parse_hcss_declarations(body)})
    return pattern.sub("", text), lists


def _parse_hcss_declarations(body: str) -> list[tuple[str, str]]:
    declarations, _ = _parse_hcss_declarations_with_meta(body)
    return declarations


def _parse_hcss_declarations_with_meta(body: str, base_line: int = 1) -> tuple[list[tuple[str, str]], list[dict[str, Any]]]:
    declarations: list[tuple[str, str]] = []
    metadata: list[dict[str, Any]] = []
    offset = 0
    for raw in body.split(";"):
        raw_start = offset
        offset += len(raw) + 1
        item = raw.strip()
        if not item:
            continue
        leading = len(raw) - len(raw.lstrip())
        line = base_line + body[: raw_start + leading].count("\n")
        if item.startswith("@hdocx-include "):
            value = item[len("@hdocx-include ") :].strip()
            declarations.append(("@hdocx-include", value))
            metadata.append({"line": line, "property": "@hdocx-include", "value": value})
            continue
        if ":" not in item:
            declarations.append(("__parse_error__", item))
            metadata.append({"line": line, "property": "__parse_error__", "value": item})
            continue
        key, value = item.split(":", 1)
        prop = key.strip()
        raw_value = value.strip()
        declarations.append((prop, raw_value))
        metadata.append({"line": line, "property": prop, "value": raw_value})
    return declarations, metadata


def _hcss_diagnostics_base(*, modified: bool) -> dict[str, Any]:
    return {
        "modified": modified,
        "rules": [],
        "parseErrors": [],
        "summary": {
            "ruleCount": 0,
            "matchedNodeCount": 0,
            "supportedDeclarationCount": 0,
            "unsupportedDeclarationCount": 0,
            "patchCount": 0,
            "noopCount": 0,
            "errorCount": 0,
        },
    }


def _update_hcss_diagnostics_summary(diagnostics: dict[str, Any]) -> None:
    rules = diagnostics.get("rules", [])
    declarations = [declaration for rule in rules for declaration in rule.get("declarations", [])]
    diagnostics["summary"] = {
        "ruleCount": len(rules),
        "matchedNodeCount": sum(rule.get("matchCount", 0) for rule in rules),
        "supportedDeclarationCount": sum(1 for declaration in declarations if declaration.get("supported")),
        "unsupportedDeclarationCount": sum(1 for declaration in declarations if not declaration.get("supported")),
        "patchCount": sum(len(rule.get("patchIds", [])) for rule in rules),
        "noopCount": sum(len(rule.get("noops", [])) for rule in rules),
        "errorCount": len(diagnostics.get("parseErrors", []))
        + sum(len(rule.get("errors", [])) for rule in rules),
    }


def _hcss_declaration_diagnostics(
    mode: str,
    declarations: list[tuple[str, str]],
    declaration_meta: list[dict[str, Any]] | None = None,
) -> list[dict[str, Any]]:
    metadata = declaration_meta or []
    diagnostics: list[dict[str, Any]] = []
    for index, (key, value) in enumerate(declarations):
        meta = metadata[index] if index < len(metadata) else {}
        diagnostic: dict[str, Any] = {
            "line": meta.get("line"),
            "property": key,
            "value": value,
            "supported": False,
        }
        if key == "__parse_error__":
            diagnostic["reason"] = "Invalid declaration syntax; expected property: value;"
            diagnostics.append(diagnostic)
            continue
        if key == "@hdocx-include":
            diagnostic["supported"] = True
            diagnostic["kind"] = "format-include"
            diagnostics.append(diagnostic)
            continue
        if not key.startswith("hdocx-"):
            diagnostic["reason"] = "Only hdocx-* declarations are accepted."
            diagnostics.append(diagnostic)
            continue
        raw_prop = key.removeprefix("hdocx-")
        paragraph_prop = _canonical_paragraph_property_name(raw_prop)
        run_prop = _canonical_run_property_name(raw_prop)
        allowed_paragraph = mode in {"paragraph-formatting", "style-definition", "direct-formatting"}
        allowed_run = mode in {"all-runs", "style-definition", "direct-formatting"}
        if mode == "image-formatting":
            image_prop = _canonical_image_property_name(raw_prop)
            paragraph_image_prop = _canonical_image_paragraph_property_name(raw_prop)
            if image_prop in IMAGE_FORMATTING_PROPERTIES:
                try:
                    normalized = _normalize_image_hcss_property(image_prop, value, old_value="1" if image_prop in {"width-emu", "height-emu"} else "")
                except ValueError as exc:
                    diagnostic.update(
                        {
                            "normalizedProperty": image_prop,
                            "reason": str(exc),
                            "ooxml": _image_ooxml_mapping(image_prop),
                        }
                    )
                else:
                    diagnostic.update(
                        {
                            "supported": True,
                            "kind": "image-format",
                            "normalizedProperty": image_prop,
                            "normalizedValue": normalized,
                            "ooxml": _image_ooxml_mapping(image_prop),
                        }
                    )
                diagnostics.append(diagnostic)
                continue
            if paragraph_image_prop in PARAGRAPH_PROPERTY_ATTRS:
                try:
                    if raw_prop in {"paragraph-line-spacing-exact", "paragraph-line-spacingExact"} and not _strip_hcss_value(value).endswith("pt"):
                        raise ValueError("paragraph-line-spacing-exact must use pt units.")
                    normalized = _normalize_paragraph_property(paragraph_image_prop, value)
                except ValueError as exc:
                    diagnostic.update(
                        {
                            "normalizedProperty": f"{IMAGE_PARAGRAPH_PROPERTY_PREFIX}{paragraph_image_prop}",
                            "reason": str(exc),
                            "ooxml": OOXML_PROPERTY_MAP.get(paragraph_image_prop),
                        }
                    )
                else:
                    diagnostic.update(
                        {
                            "supported": True,
                            "kind": "image-paragraph-format",
                            "normalizedProperty": f"{IMAGE_PARAGRAPH_PROPERTY_PREFIX}{paragraph_image_prop}",
                            "normalizedValue": normalized,
                            "ooxml": OOXML_PROPERTY_MAP.get(paragraph_image_prop),
                        }
                    )
                diagnostics.append(diagnostic)
                continue
        if (mode == "paragraph-formatting" and raw_prop == "manual-page-break-before") or (
            mode == "paragraph-structure" and raw_prop in PARAGRAPH_STRUCTURAL_PROPERTIES
        ):
            try:
                normalized = _normalize_hcss_bool(raw_prop, value)
            except ValueError as exc:
                diagnostic.update(
                    {
                        "normalizedProperty": raw_prop,
                        "reason": str(exc),
                        "ooxml": OOXML_PROPERTY_MAP.get(raw_prop),
                    }
                )
            else:
                diagnostic.update(
                    {
                        "supported": True,
                        "kind": "structural-paragraph",
                        "normalizedProperty": raw_prop,
                        "normalizedValue": normalized,
                        "ooxml": OOXML_PROPERTY_MAP.get(raw_prop),
                    }
                )
            diagnostics.append(diagnostic)
            continue
        if mode == "paragraph-structure" and raw_prop.startswith(EMPTY_PARAGRAPH_PROPERTY_PREFIX):
            raw_empty_prop = raw_prop.removeprefix(EMPTY_PARAGRAPH_PROPERTY_PREFIX)
            empty_prop = _canonical_empty_paragraph_property_name(raw_empty_prop)
            try:
                normalized = _normalize_empty_paragraph_property(raw_empty_prop, value)
            except ValueError as exc:
                diagnostic.update(
                    {
                        "normalizedProperty": f"{EMPTY_PARAGRAPH_PROPERTY_PREFIX}{empty_prop}",
                        "reason": str(exc),
                        "ooxml": OOXML_PROPERTY_MAP.get(f"{EMPTY_PARAGRAPH_PROPERTY_PREFIX}{empty_prop}"),
                    }
                )
            else:
                diagnostic.update(
                    {
                        "supported": True,
                        "kind": "empty-paragraph-format",
                        "normalizedProperty": f"{EMPTY_PARAGRAPH_PROPERTY_PREFIX}{empty_prop}",
                        "normalizedValue": normalized,
                        "ooxml": OOXML_PROPERTY_MAP.get(f"{EMPTY_PARAGRAPH_PROPERTY_PREFIX}{empty_prop}"),
                    }
                )
            diagnostics.append(diagnostic)
            continue
        if allowed_paragraph and paragraph_prop in PARAGRAPH_PROPERTY_ATTRS:
            try:
                if raw_prop in {"line-spacing-exact", "line-spacingExact"} and not _strip_hcss_value(value).endswith("pt"):
                    raise ValueError("line-spacing-exact must use pt units.")
                normalized = _normalize_paragraph_property(paragraph_prop, value)
            except ValueError as exc:
                diagnostic.update(
                    {
                        "normalizedProperty": paragraph_prop,
                        "reason": str(exc),
                        "ooxml": OOXML_PROPERTY_MAP.get(paragraph_prop),
                    }
                )
            else:
                diagnostic.update(
                    {
                        "supported": True,
                        "normalizedProperty": paragraph_prop,
                        "normalizedValue": normalized,
                        "ooxml": OOXML_PROPERTY_MAP.get(paragraph_prop),
                    }
                )
            diagnostics.append(diagnostic)
            continue
        if allowed_run and run_prop in RUN_PROPERTY_ATTRS:
            try:
                normalized = _normalize_run_property(run_prop, value)
            except ValueError as exc:
                diagnostic.update(
                    {
                        "normalizedProperty": run_prop,
                        "reason": str(exc),
                        "ooxml": OOXML_PROPERTY_MAP.get(run_prop),
                    }
                )
            else:
                diagnostic.update(
                    {
                        "supported": True,
                        "normalizedProperty": run_prop,
                        "normalizedValue": normalized,
                        "ooxml": OOXML_PROPERTY_MAP.get(run_prop),
                    }
                )
            diagnostics.append(diagnostic)
            continue
        diagnostic["normalizedProperty"] = _canonical_hcss_property_name(raw_prop)
        if paragraph_prop in PARAGRAPH_PROPERTY_ATTRS or run_prop in RUN_PROPERTY_ATTRS:
            diagnostic["reason"] = f"Property is supported, but not in {mode} mode."
        else:
            diagnostic["reason"] = "Unsupported H-CSS property."
        diagnostic["ooxml"] = OOXML_PROPERTY_MAP.get(diagnostic["normalizedProperty"])
        diagnostics.append(diagnostic)
    return diagnostics


def _annotate_hcss_errors(
    errors: list[dict[str, Any]],
    declaration_diagnostics: list[dict[str, Any]],
    rule_line: int | None,
) -> None:
    for error in errors:
        if "line" in error:
            continue
        prop = error.get("property")
        matched_line = None
        if prop is not None:
            prop_text = str(prop)
            for declaration in declaration_diagnostics:
                names = {
                    str(declaration.get("property")),
                    str(declaration.get("normalizedProperty")),
                }
                if prop_text in names or f"hdocx-{prop_text}" in names:
                    matched_line = declaration.get("line")
                    break
        error["line"] = matched_line or rule_line


def _parse_hcss_set(name: str, declarations: list[tuple[str, str]]) -> dict[str, Any]:
    result: dict[str, Any] = {"name": name, "select": None, "exclude": None, "allowEmpty": False}
    for key, value in declarations:
        if key == "select":
            result["select"] = value
        elif key == "exclude":
            result["exclude"] = value
        elif key == "allow-empty":
            result["allowEmpty"] = value.lower() == "true"
    return result


def _expand_hcss_declarations(
    declarations: list[tuple[str, str]],
    formats: dict[str, list[tuple[str, str]]],
    tokens: dict[str, str],
) -> dict[str, Any]:
    expanded: list[tuple[str, str]] = []
    errors: list[dict[str, Any]] = []
    for key, value in declarations:
        if key == "__parse_error__":
            errors.append({"code": "HCSS_DECLARATION_PARSE_ERROR", "message": "Invalid declaration.", "value": value})
            continue
        if key == "@hdocx-include":
            included = formats.get(value)
            if included is None:
                errors.append({"code": "HCSS_UNKNOWN_FORMAT", "message": "Unknown @hdocx-format.", "format": value})
                continue
            for inc_key, inc_value in included:
                expanded.append((inc_key, _resolve_hcss_tokens(inc_value, tokens)))
            continue
        expanded.append((key, _resolve_hcss_tokens(value, tokens)))
    return {"declarations": expanded, "errors": errors}


def _resolve_hcss_tokens(value: str, tokens: dict[str, str]) -> str:
    def replace(match: re.Match[str]) -> str:
        name = match.group(1)
        return tokens.get(name, match.group(0))

    return re.sub(r"token\(([A-Za-z0-9_.-]+)\)", replace, value)


def _strip_hcss_value(value: str) -> str:
    value = value.strip()
    if (value.startswith('"') and value.endswith('"')) or (value.startswith("'") and value.endswith("'")):
        return value[1:-1]
    return value


def _compile_hcss_image_insertions(
    bundle_dir: Path,
    insertions: list[dict[str, Any]],
    sets: dict[str, dict[str, Any]],
    html_nodes: dict[str, Any],
    manifest: dict[str, Any],
    planned_patches: list[dict[str, Any]],
    edit_index: int,
) -> tuple[list[dict[str, Any]], list[dict[str, Any]], int]:
    errors: list[dict[str, Any]] = []
    patches: list[dict[str, Any]] = []
    if not insertions:
        return errors, patches, edit_index

    existing_entries = {entry["path"] for entry in manifest.get("package", {}).get("entries", [])}
    planned_entries: set[str] = set()
    next_rel_nums: dict[str, int] = {}
    next_doc_pr_id = _next_doc_pr_id(bundle_dir)

    for insertion in insertions:
        target_ids = _resolve_hcss_target(insertion["target"], sets, html_nodes)
        if len(target_ids) != 1:
            errors.append(
                {
                    "code": "HCSS_INSERT_IMAGE_TARGET_COUNT",
                    "message": "Image insertion requires exactly one paragraph target.",
                    "target": insertion["target"],
                    "matched": target_ids,
                }
            )
            continue
        target_id = target_ids[0]
        target_node = manifest.get("nodes", {}).get(target_id)
        if not target_node or target_node.get("kind") != "paragraph":
            errors.append(
                {
                    "code": "HCSS_INSERT_IMAGE_TARGET_KIND",
                    "message": "Image insertion target must be a paragraph.",
                    "target": insertion["target"],
                    "nodeId": target_id,
                }
            )
            continue
        part_path = target_node.get("partPath")
        if not _image_insertion_part_supported(part_path):
            errors.append(
                {
                    "code": "HCSS_INSERT_IMAGE_PART_UNSUPPORTED",
                    "message": "Image insertion supports projected Word XML parts only.",
                    "nodeId": target_id,
                    "partPath": part_path,
                }
            )
            continue
        if insertion.get("position") not in {"after", "before"}:
            errors.append(
                {
                    "code": "HCSS_INSERT_IMAGE_POSITION_UNSUPPORTED",
                    "message": "Image insertion currently supports after(...) and before(...).",
                    "position": insertion.get("position"),
                }
            )
            continue
        props, prop_errors = _image_insertion_properties(bundle_dir, insertion["declarations"])
        if prop_errors:
            errors.extend(prop_errors)
            continue
        source_path = props["sourcePath"]
        extension = source_path.suffix.lower()
        content_type = IMAGE_CONTENT_TYPES[extension]
        media_entry = _next_media_entry(extension, existing_entries | planned_entries)
        planned_entries.add(media_entry)
        rels_entry = _rels_entry_for_part(part_path)
        if rels_entry not in next_rel_nums:
            next_rel_nums[rels_entry] = _next_planned_relationship_number(bundle_dir, rels_entry, planned_patches + patches)
        relationship_id = f"rId{next_rel_nums[rels_entry]}"
        next_rel_nums[rels_entry] += 1
        doc_pr_id = next_doc_pr_id
        next_doc_pr_id += 1

        edit_index += 1
        operation = f"insert-image-{insertion['position']}-paragraph"
        patches.append(
            {
                "id": f"patch-{edit_index:06d}",
                "operation": operation,
                "partPath": part_path,
                "entryName": part_path.lstrip("/"),
                "nodeId": target_id,
                "locator": target_node["locator"],
                "relationshipId": relationship_id,
                "docPrId": doc_pr_id,
                "mediaEntryName": media_entry,
                "mediaFileName": Path(media_entry).name,
                "alt": props.get("alt"),
                "widthEmu": props["widthEmu"],
                "heightEmu": props["heightEmu"],
            }
        )
        edit_index += 1
        patches.append(
            {
                "id": f"patch-{edit_index:06d}",
                "operation": "patch-relationships",
                "partPath": "/" + rels_entry,
                "entryName": rels_entry,
                "relationshipId": relationship_id,
                "target": "media/" + Path(media_entry).name,
            }
        )
        edit_index += 1
        patches.append(
            {
                "id": f"patch-{edit_index:06d}",
                "operation": "patch-content-types",
                "partPath": "/[Content_Types].xml",
                "entryName": "[Content_Types].xml",
                "extension": extension.lstrip("."),
                "contentType": content_type,
            }
        )
        edit_index += 1
        patches.append(
            {
                "id": f"patch-{edit_index:06d}",
                "operation": "add-media",
                "partPath": "/" + media_entry,
                "entryName": media_entry,
                "sourcePath": str(source_path.relative_to(bundle_dir).as_posix()),
                "newSha256": sha256_file(source_path),
                "newSize": source_path.stat().st_size,
            }
        )
    return errors, patches, edit_index


def _compile_hcss_table_row_insertions(
    insertions: list[dict[str, Any]],
    sets: dict[str, dict[str, Any]],
    html_nodes: dict[str, Any],
    manifest: dict[str, Any],
    edit_index: int,
) -> tuple[list[dict[str, Any]], list[dict[str, Any]], int]:
    errors: list[dict[str, Any]] = []
    patches: list[dict[str, Any]] = []
    if not insertions:
        return errors, patches, edit_index
    manifest_nodes = manifest.get("nodes", {})
    for insertion in insertions:
        target_ids = _resolve_hcss_target(insertion["target"], sets, html_nodes)
        if len(target_ids) != 1:
            errors.append(
                {
                    "code": "HCSS_INSERT_TABLE_ROW_TARGET_COUNT",
                    "message": "Table row insertion requires exactly one row target.",
                    "target": insertion["target"],
                    "matched": target_ids,
                }
            )
            continue
        target_id = target_ids[0]
        target_node = manifest_nodes.get(target_id)
        if not target_node or target_node.get("kind") != "table-row":
            errors.append(
                {
                    "code": "HCSS_INSERT_TABLE_ROW_TARGET_KIND",
                    "message": "Table row insertion target must be a table row.",
                    "target": insertion["target"],
                    "nodeId": target_id,
                }
            )
            continue
        if insertion.get("position") != "after":
            errors.append(
                {
                    "code": "HCSS_INSERT_TABLE_ROW_POSITION_UNSUPPORTED",
                    "message": "Table row insertion currently supports after(...) only.",
                    "position": insertion.get("position"),
                }
            )
            continue
        table_node = manifest_nodes.get(target_node.get("parent", ""))
        if not target_node.get("simpleEditable") or not table_node or not table_node.get("simpleEditable"):
            errors.append(
                {
                    "code": "HCSS_INSERT_TABLE_ROW_COMPLEX_TABLE",
                    "message": "Table row insertion is only supported for simple tables without merged cells.",
                    "nodeId": target_id,
                }
            )
            continue
        cell_ids = target_node.get("children", [])
        cell_texts, prop_errors = _table_row_insertion_properties(insertion["declarations"], len(cell_ids))
        if prop_errors:
            errors.extend(prop_errors)
            continue
        edit_index += 1
        patches.append(
            {
                "id": f"patch-{edit_index:06d}",
                "operation": "insert-table-row-after",
                "partPath": target_node["partPath"],
                "entryName": target_node["partPath"].lstrip("/"),
                "nodeId": target_id,
                "locator": target_node["locator"],
                "cellTexts": cell_texts,
            }
        )
    return errors, patches, edit_index


def _compile_hcss_table_row_deletions(
    deletions: list[dict[str, Any]],
    sets: dict[str, dict[str, Any]],
    html_nodes: dict[str, Any],
    manifest: dict[str, Any],
    edit_index: int,
) -> tuple[list[dict[str, Any]], list[dict[str, Any]], int]:
    errors: list[dict[str, Any]] = []
    patches: list[dict[str, Any]] = []
    if not deletions:
        return errors, patches, edit_index
    manifest_nodes = manifest.get("nodes", {})
    for deletion in deletions:
        target_ids = _resolve_hcss_target(deletion["target"], sets, html_nodes)
        if len(target_ids) != 1:
            errors.append(
                {
                    "code": "HCSS_DELETE_TABLE_ROW_TARGET_COUNT",
                    "message": "Table row deletion requires exactly one row target.",
                    "target": deletion["target"],
                    "matched": target_ids,
                }
            )
            continue
        target_id = target_ids[0]
        target_node = manifest_nodes.get(target_id)
        if not target_node or target_node.get("kind") != "table-row":
            errors.append(
                {
                    "code": "HCSS_DELETE_TABLE_ROW_TARGET_KIND",
                    "message": "Table row deletion target must be a table row.",
                    "target": deletion["target"],
                    "nodeId": target_id,
                }
            )
            continue
        table_node = manifest_nodes.get(target_node.get("parent", ""))
        if not target_node.get("simpleEditable") or not table_node or not table_node.get("simpleEditable"):
            errors.append(
                {
                    "code": "HCSS_DELETE_TABLE_ROW_COMPLEX_TABLE",
                    "message": "Table row deletion is only supported for simple tables without merged cells.",
                    "nodeId": target_id,
                }
            )
            continue
        if len(table_node.get("children", [])) <= 1:
            errors.append(
                {
                    "code": "HCSS_DELETE_TABLE_ROW_LAST_ROW",
                    "message": "Deleting the only row in a table is not supported.",
                    "nodeId": target_id,
                }
            )
            continue
        edit_index += 1
        patches.append(
            {
                "id": f"patch-{edit_index:06d}",
                "operation": "delete-table-row",
                "partPath": target_node["partPath"],
                "entryName": target_node["partPath"].lstrip("/"),
                "nodeId": target_id,
                "locator": target_node["locator"],
            }
        )
    return errors, patches, edit_index


def _compile_hcss_style_creations(
    bundle_dir: Path,
    style_creations: list[dict[str, Any]],
    manifest: dict[str, Any],
    planned_patches: list[dict[str, Any]],
    edit_index: int,
) -> tuple[list[dict[str, Any]], list[dict[str, Any]], dict[str, Any], int]:
    errors: list[dict[str, Any]] = []
    patches: list[dict[str, Any]] = []
    created: dict[str, Any] = {}
    if not style_creations:
        return errors, patches, created, edit_index
    if not _manifest_has_entry(manifest, "word/styles.xml"):
        bootstrap_patches, edit_index = _bootstrap_word_part_patches(
            bundle_dir,
            edit_index,
            WORD_DOCUMENT_RELS_ENTRY,
            "word/styles.xml",
            "/word/styles.xml",
            STYLES_CONTENT_TYPE,
            STYLES_REL_TYPE,
            "styles.xml",
            planned_patches + patches,
        )
        patches.extend(bootstrap_patches)
    existing = set(manifest.get("styles", {}))
    for style_def in style_creations:
        style_id = style_def["styleId"]
        if style_id in existing or style_id in created:
            errors.append(
                {
                    "code": "HCSS_STYLE_ALREADY_EXISTS",
                    "message": "Style id already exists.",
                    "styleId": style_id,
                }
            )
            continue
        metadata, paragraph_decls, run_decls, style_errors = _style_creation_declarations(style_def["declarations"])
        errors.extend(style_errors)
        if style_errors:
            continue
        if metadata["type"] != "paragraph":
            errors.append(
                {
                    "code": "HCSS_STYLE_TYPE_UNSUPPORTED",
                    "message": "Only paragraph style creation is supported in this phase.",
                    "styleId": style_id,
                    "type": metadata["type"],
                }
            )
            continue
        paragraph_props, paragraph_errors = _hcss_paragraph_properties(paragraph_decls)
        run_props, run_errors = _hcss_run_properties(run_decls)
        errors.extend(paragraph_errors)
        errors.extend(run_errors)
        if paragraph_errors or run_errors:
            continue
        snapshot = {
            "styleId": style_id,
            "type": metadata["type"],
            "name": metadata.get("name") or style_id,
            "basedOn": metadata.get("basedOn"),
            "paragraphProperties": paragraph_props,
            "runProperties": run_props,
        }
        created[style_id] = snapshot
        edit_index += 1
        patches.append(
            {
                "id": f"patch-{edit_index:06d}",
                "operation": "create-style",
                "partPath": "/word/styles.xml",
                "entryName": "word/styles.xml",
                "styleId": style_id,
                "styleType": metadata["type"],
                "name": snapshot["name"],
                "basedOn": metadata.get("basedOn"),
                "next": metadata.get("next"),
                "qFormat": metadata.get("qFormat", True),
                "newProperties": {
                    "paragraph": paragraph_props,
                    "run": run_props,
                },
            }
        )
    return errors, patches, created, edit_index


def _compile_hcss_style_deletions(
    style_deletions: list[dict[str, Any]],
    manifest: dict[str, Any],
    edit_index: int,
) -> tuple[list[dict[str, Any]], list[dict[str, Any]], int]:
    errors: list[dict[str, Any]] = []
    patches: list[dict[str, Any]] = []
    if not style_deletions:
        return errors, patches, edit_index
    if not _manifest_has_entry(manifest, "word/styles.xml"):
        return (
            [{"code": "HCSS_STYLES_PART_MISSING", "message": "@hdocx-delete-style requires word/styles.xml."}],
            patches,
            edit_index,
        )
    styles = manifest.get("styles", {})
    used_style_ids = {
        node.get("styleId")
        for node in manifest.get("nodes", {}).values()
        if node.get("kind") == "paragraph" and node.get("styleId")
    }
    deleted: set[str] = set()
    for deletion in style_deletions:
        style_id = deletion["styleId"]
        style = styles.get(style_id)
        if style is None:
            errors.append({"code": "HCSS_DELETE_STYLE_NOT_FOUND", "message": "Style id does not exist.", "styleId": style_id})
            continue
        if style_id in used_style_ids:
            errors.append({"code": "HCSS_DELETE_STYLE_IN_USE", "message": "Style is still used by projected paragraphs.", "styleId": style_id})
            continue
        if style.get("default"):
            errors.append({"code": "HCSS_DELETE_STYLE_DEFAULT", "message": "Default styles cannot be deleted safely.", "styleId": style_id})
            continue
        if style_id in deleted:
            errors.append({"code": "HCSS_DELETE_STYLE_DUPLICATE", "message": "Style id is requested for deletion more than once.", "styleId": style_id})
            continue
        deleted.add(style_id)
        edit_index += 1
        patches.append(
            {
                "id": f"patch-{edit_index:06d}",
                "operation": "delete-style",
                "partPath": "/word/styles.xml",
                "entryName": "word/styles.xml",
                "styleId": style_id,
            }
        )
    return errors, patches, edit_index


def _compile_hcss_list_creations(
    bundle_dir: Path,
    list_creations: list[dict[str, Any]],
    manifest: dict[str, Any],
    planned_patches: list[dict[str, Any]],
    edit_index: int,
) -> tuple[list[dict[str, Any]], list[dict[str, Any]], dict[str, Any], int]:
    errors: list[dict[str, Any]] = []
    patches: list[dict[str, Any]] = []
    created: dict[str, Any] = {}
    if not list_creations:
        return errors, patches, created, edit_index
    if not _manifest_has_entry(manifest, "word/numbering.xml"):
        bootstrap_patches, edit_index = _bootstrap_word_part_patches(
            bundle_dir,
            edit_index,
            WORD_DOCUMENT_RELS_ENTRY,
            "word/numbering.xml",
            "/word/numbering.xml",
            NUMBERING_CONTENT_TYPE,
            NUMBERING_REL_TYPE,
            "numbering.xml",
            planned_patches + patches,
        )
        patches.extend(bootstrap_patches)
    numbering = manifest.get("numbering", {})
    next_abstract = _next_numeric_id(numbering.get("abstractNums", {}))
    next_num = _next_numeric_id(numbering.get("nums", {}))
    for list_def in list_creations:
        alias = list_def["listId"]
        if alias in created:
            errors.append({"code": "HCSS_LIST_ALREADY_EXISTS", "message": "List alias already exists.", "listId": alias})
            continue
        levels, property_errors = _hcss_list_level_properties(list_def["declarations"])
        errors.extend(property_errors)
        if property_errors:
            continue
        properties = levels["0"]
        abstract_num_id = str(next_abstract)
        num_id = str(next_num)
        next_abstract += 1
        next_num += 1
        created[alias] = {"abstractNumId": abstract_num_id, "numId": num_id, "ilvl": "0", "levels": levels}
        edit_index += 1
        patches.append(
            {
                "id": f"patch-{edit_index:06d}",
                "operation": "create-numbering-list",
                "partPath": "/word/numbering.xml",
                "entryName": "word/numbering.xml",
                "listId": alias,
                "abstractNumId": abstract_num_id,
                "numId": num_id,
                "ilvl": "0",
                "newProperties": properties,
                "levels": levels,
            }
        )
    return errors, patches, created, edit_index


def _next_numeric_id(items: dict[str, Any]) -> int:
    used = [int(key) for key in items if str(key).isdigit()]
    return max(used, default=-1) + 1


def _style_creation_declarations(
    declarations: list[tuple[str, str]],
) -> tuple[dict[str, Any], list[tuple[str, str]], list[tuple[str, str]], list[dict[str, Any]]]:
    metadata: dict[str, Any] = {"type": "paragraph", "qFormat": True}
    paragraph_decls: list[tuple[str, str]] = []
    run_decls: list[tuple[str, str]] = []
    errors: list[dict[str, Any]] = []
    for key, value in declarations:
        if key == "__parse_error__":
            errors.append({"code": "HCSS_DECLARATION_PARSE_ERROR", "message": "Invalid declaration.", "value": value})
            continue
        stripped = _strip_hcss_value(value)
        if key == "type":
            metadata["type"] = stripped
        elif key == "name":
            metadata["name"] = stripped
        elif key == "based-on":
            metadata["basedOn"] = stripped
        elif key == "next":
            metadata["next"] = stripped
        elif key == "q-format":
            metadata["qFormat"] = stripped.lower() in {"true", "1", "yes", "on"}
        elif key.startswith("hdocx-"):
            prop = key.removeprefix("hdocx-")
            if prop in PARAGRAPH_PROPERTY_ATTRS:
                paragraph_decls.append((key, value))
            elif prop in RUN_PROPERTY_ATTRS:
                run_decls.append((key, value))
            else:
                errors.append({"code": "HCSS_STYLE_PROPERTY_UNSUPPORTED", "message": "Unsupported style property.", "property": key})
        else:
            errors.append({"code": "HCSS_STYLE_DECLARATION_UNSUPPORTED", "message": "Unsupported @hdocx-style declaration.", "property": key})
    return metadata, paragraph_decls, run_decls, errors


def _compile_hcss_table_column_insertions(
    insertions: list[dict[str, Any]],
    sets: dict[str, dict[str, Any]],
    html_nodes: dict[str, Any],
    manifest: dict[str, Any],
    edit_index: int,
) -> tuple[list[dict[str, Any]], list[dict[str, Any]], int]:
    errors: list[dict[str, Any]] = []
    patches: list[dict[str, Any]] = []
    if not insertions:
        return errors, patches, edit_index
    manifest_nodes = manifest.get("nodes", {})
    for insertion in insertions:
        target_ids = _resolve_hcss_target(insertion["target"], sets, html_nodes)
        if len(target_ids) != 1:
            errors.append(
                {
                    "code": "HCSS_INSERT_TABLE_COLUMN_TARGET_COUNT",
                    "message": "Table column insertion requires exactly one cell target.",
                    "target": insertion["target"],
                    "matched": target_ids,
                }
            )
            continue
        target_id = target_ids[0]
        target_node = manifest_nodes.get(target_id)
        table_node, table_errors = _validate_table_column_target(target_id, target_node, manifest_nodes, "insert")
        if table_errors:
            errors.extend(table_errors)
            continue
        if insertion.get("position") != "after":
            errors.append(
                {
                    "code": "HCSS_INSERT_TABLE_COLUMN_POSITION_UNSUPPORTED",
                    "message": "Table column insertion currently supports after(...) only.",
                    "position": insertion.get("position"),
                }
            )
            continue
        row_ids = table_node.get("children", [])
        cell_texts, prop_errors = _table_column_text_properties(insertion["declarations"], len(row_ids), "insert")
        if prop_errors:
            errors.extend(prop_errors)
            continue
        edit_index += 1
        patches.append(
            {
                "id": f"patch-{edit_index:06d}",
                "operation": "insert-table-column-after",
                "partPath": target_node["partPath"],
                "entryName": target_node["partPath"].lstrip("/"),
                "nodeId": target_id,
                "locator": target_node["locator"],
                "cellTexts": cell_texts,
            }
        )
    return errors, patches, edit_index


def _compile_hcss_table_column_deletions(
    deletions: list[dict[str, Any]],
    sets: dict[str, dict[str, Any]],
    html_nodes: dict[str, Any],
    manifest: dict[str, Any],
    edit_index: int,
) -> tuple[list[dict[str, Any]], list[dict[str, Any]], int]:
    errors: list[dict[str, Any]] = []
    patches: list[dict[str, Any]] = []
    if not deletions:
        return errors, patches, edit_index
    manifest_nodes = manifest.get("nodes", {})
    for deletion in deletions:
        target_ids = _resolve_hcss_target(deletion["target"], sets, html_nodes)
        if len(target_ids) != 1:
            errors.append(
                {
                    "code": "HCSS_DELETE_TABLE_COLUMN_TARGET_COUNT",
                    "message": "Table column deletion requires exactly one cell target.",
                    "target": deletion["target"],
                    "matched": target_ids,
                }
            )
            continue
        target_id = target_ids[0]
        target_node = manifest_nodes.get(target_id)
        table_node, table_errors = _validate_table_column_target(target_id, target_node, manifest_nodes, "delete")
        if table_errors:
            errors.extend(table_errors)
            continue
        first_row = manifest_nodes.get(table_node.get("children", [""])[0], {})
        if len(first_row.get("children", [])) <= 1:
            errors.append(
                {
                    "code": "HCSS_DELETE_TABLE_COLUMN_LAST_COLUMN",
                    "message": "Deleting the only column in a table is not supported.",
                    "nodeId": target_id,
                }
            )
            continue
        edit_index += 1
        patches.append(
            {
                "id": f"patch-{edit_index:06d}",
                "operation": "delete-table-column",
                "partPath": target_node["partPath"],
                "entryName": target_node["partPath"].lstrip("/"),
                "nodeId": target_id,
                "locator": target_node["locator"],
            }
        )
    return errors, patches, edit_index


def _validate_table_column_target(
    target_id: str,
    target_node: dict[str, Any] | None,
    manifest_nodes: dict[str, Any],
    action: str,
) -> tuple[dict[str, Any], list[dict[str, Any]]]:
    errors: list[dict[str, Any]] = []
    if not target_node or target_node.get("kind") != "table-cell":
        errors.append(
            {
                "code": f"HCSS_{action.upper()}_TABLE_COLUMN_TARGET_KIND",
                "message": "Table column operation target must be a table cell.",
                "nodeId": target_id,
            }
        )
        return {}, errors
    row_node = manifest_nodes.get(target_node.get("parent", ""))
    table_node = manifest_nodes.get(row_node.get("parent", "") if row_node else "")
    if not target_node.get("simpleEditable") or not row_node or not table_node or not table_node.get("simpleEditable"):
        errors.append(
            {
                "code": f"HCSS_{action.upper()}_TABLE_COLUMN_COMPLEX_TABLE",
                "message": "Table column operation is only supported for simple tables without merged cells.",
                "nodeId": target_id,
            }
        )
        return {}, errors
    expected_cols = len(row_node.get("children", []))
    for row_id in table_node.get("children", []):
        row = manifest_nodes.get(row_id, {})
        if len(row.get("children", [])) != expected_cols:
            errors.append(
                {
                    "code": f"HCSS_{action.upper()}_TABLE_COLUMN_IRREGULAR_TABLE",
                    "message": "Table column operation requires every row to have the same cell count.",
                    "nodeId": target_id,
                }
            )
            break
    return table_node, errors


def _table_column_text_properties(
    declarations: list[tuple[str, str]],
    expected_count: int,
    action: str,
) -> tuple[list[str], list[dict[str, Any]]]:
    errors: list[dict[str, Any]] = []
    raw: dict[str, str] = {}
    for key, value in declarations:
        if key == "__parse_error__":
            errors.append({"code": "HCSS_DECLARATION_PARSE_ERROR", "message": "Invalid declaration.", "value": value})
            continue
        raw[key] = _strip_hcss_value(value)
    if "cells" not in raw:
        return ["" for _ in range(expected_count)], errors
    values = [item.strip() for item in raw["cells"].split("|")]
    if len(values) != expected_count:
        errors.append(
            {
                "code": f"HCSS_{action.upper()}_TABLE_COLUMN_CELL_COUNT",
                "message": "Inserted column cell count must match the table row count.",
                "expected": expected_count,
                "actual": len(values),
            }
        )
        return [], errors
    return values, errors


def _table_row_insertion_properties(
    declarations: list[tuple[str, str]],
    expected_count: int,
) -> tuple[list[str], list[dict[str, Any]]]:
    errors: list[dict[str, Any]] = []
    raw: dict[str, str] = {}
    for key, value in declarations:
        if key == "__parse_error__":
            errors.append({"code": "HCSS_DECLARATION_PARSE_ERROR", "message": "Invalid declaration.", "value": value})
            continue
        raw[key] = _strip_hcss_value(value)
    if "cells" not in raw:
        return ["" for _ in range(expected_count)], errors
    values = [item.strip() for item in raw["cells"].split("|")]
    if len(values) != expected_count:
        errors.append(
            {
                "code": "HCSS_INSERT_TABLE_ROW_CELL_COUNT",
                "message": "Inserted row cell count must match the target row.",
                "expected": expected_count,
                "actual": len(values),
            }
        )
        return [], errors
    return values, errors


def _image_insertion_properties(
    bundle_dir: Path,
    declarations: list[tuple[str, str]],
) -> tuple[dict[str, Any], list[dict[str, Any]]]:
    raw: dict[str, str] = {}
    errors: list[dict[str, Any]] = []
    for key, value in declarations:
        if key == "__parse_error__":
            errors.append({"code": "HCSS_DECLARATION_PARSE_ERROR", "message": "Invalid declaration.", "value": value})
            continue
        raw[key] = _strip_hcss_value(value)
    for required in ("source", "width-emu", "height-emu"):
        if required not in raw:
            errors.append(
                {
                    "code": "HCSS_INSERT_IMAGE_MISSING_PROPERTY",
                    "message": "Image insertion is missing a required property.",
                    "property": required,
                }
            )
    if errors:
        return {}, errors

    source = Path(raw["source"])
    if source.is_absolute() or any(part in {"", ".", ".."} for part in source.parts):
        errors.append(
            {
                "code": "HCSS_INSERT_IMAGE_UNSAFE_SOURCE",
                "message": "Image source must be a safe relative path inside the H-DOCX bundle.",
                "source": raw["source"],
            }
        )
        return {}, errors
    source_path = (bundle_dir / source).resolve()
    bundle_root = bundle_dir.resolve()
    try:
        source_path.relative_to(bundle_root)
    except ValueError:
        errors.append(
            {
                "code": "HCSS_INSERT_IMAGE_UNSAFE_SOURCE",
                "message": "Image source resolved outside the H-DOCX bundle.",
                "source": raw["source"],
            }
        )
        return {}, errors
    if not source_path.is_file():
        errors.append(
            {
                "code": "HCSS_INSERT_IMAGE_SOURCE_MISSING",
                "message": "Image source file was not found.",
                "source": raw["source"],
            }
        )
        return {}, errors
    extension = source_path.suffix.lower()
    if extension not in IMAGE_CONTENT_TYPES:
        errors.append(
            {
                "code": "HCSS_INSERT_IMAGE_UNSUPPORTED_TYPE",
                "message": "Unsupported image source extension.",
                "extension": extension,
            }
        )
        return {}, errors

    width = _normalize_positive_int(raw["width-emu"], "width-emu", errors)
    height = _normalize_positive_int(raw["height-emu"], "height-emu", errors)
    if errors:
        return {}, errors
    return {
        "sourcePath": source_path,
        "widthEmu": width,
        "heightEmu": height,
        "alt": raw.get("alt"),
    }, errors


def _normalize_positive_int(raw_value: str, property_name: str, errors: list[dict[str, Any]]) -> str:
    value = raw_value.strip()
    if not value.isdigit() or int(value) <= 0:
        errors.append(
            {
                "code": "HCSS_INSERT_IMAGE_INVALID_SIZE",
                "message": "Image size must be a positive EMU integer.",
                "property": property_name,
                "value": raw_value,
            }
        )
        return "0"
    return str(int(value))


def _image_insertion_part_supported(part_path: str | None) -> bool:
    if not part_path:
        return False
    return part_path.startswith("/word/") and part_path.endswith(".xml") and "/_rels/" not in part_path


def _rels_entry_for_part(part_path: str) -> str:
    entry = part_path.lstrip("/")
    entry_path = Path(entry)
    return str(entry_path.parent / "_rels" / f"{entry_path.name}.rels").replace("\\", "/")


def _next_media_entry(extension: str, used_entries: set[str]) -> str:
    index = 1
    while True:
        candidate = f"word/media/hdocx-image-{index:06d}{extension}"
        if candidate not in used_entries:
            return candidate
        index += 1


def _next_relationship_number(bundle_dir: Path, rels_entry: str = "word/_rels/document.xml.rels") -> int:
    rels_path = bundle_dir / "parts" / rels_entry
    max_id = 0
    if rels_path.exists():
        try:
            root = ET.parse(rels_path).getroot()
            for rel in root.findall(f"{{{PKG_REL_NS}}}Relationship"):
                rel_id = rel.attrib.get("Id", "")
                match = re.fullmatch(r"rId(\d+)", rel_id)
                if match:
                    max_id = max(max_id, int(match.group(1)))
        except ET.ParseError:
            pass
    return max_id + 1


def _next_planned_relationship_number(
    bundle_dir: Path,
    rels_entry: str,
    planned_patches: list[dict[str, Any]],
) -> int:
    next_id = _next_relationship_number(bundle_dir, rels_entry)
    for patch in planned_patches:
        if patch.get("operation") != "patch-relationships" or patch.get("entryName") != rels_entry:
            continue
        rel_id = str(patch.get("relationshipId", ""))
        match = re.fullmatch(r"rId(\d+)", rel_id)
        if match:
            next_id = max(next_id, int(match.group(1)) + 1)
    return next_id


def _bootstrap_word_part_patches(
    bundle_dir: Path,
    edit_index: int,
    rels_entry: str,
    entry_name: str,
    part_name: str,
    content_type: str,
    relationship_type: str,
    relationship_target: str,
    planned_patches: list[dict[str, Any]],
) -> tuple[list[dict[str, Any]], int]:
    relationship_id = f"rId{_next_planned_relationship_number(bundle_dir, rels_entry, planned_patches)}"
    edit_index += 1
    content_type_patch = {
        "id": f"patch-{edit_index:06d}",
        "operation": "patch-content-types",
        "partPath": "/[Content_Types].xml",
        "entryName": "[Content_Types].xml",
        "partName": part_name,
        "contentType": content_type,
    }
    edit_index += 1
    relationship_patch = {
        "id": f"patch-{edit_index:06d}",
        "operation": "patch-relationships",
        "partPath": "/" + rels_entry,
        "entryName": rels_entry,
        "relationshipId": relationship_id,
        "relationshipType": relationship_type,
        "target": relationship_target,
        "createdPartEntryName": entry_name,
    }
    return [content_type_patch, relationship_patch], edit_index


def _manifest_has_entry(manifest: dict[str, Any], entry_name: str) -> bool:
    if "/" + entry_name in manifest.get("parts", {}):
        return True
    return any(entry.get("path") == entry_name for entry in manifest.get("package", {}).get("entries", []))


def _next_doc_pr_id(bundle_dir: Path) -> int:
    word_dir = bundle_dir / "parts" / "word"
    if not word_dir.exists():
        return 1
    ids: list[int] = []
    for xml_path in word_dir.rglob("*.xml"):
        if "_rels" in xml_path.parts:
            continue
        text = xml_path.read_text(encoding="utf-8", errors="ignore")
        ids.extend(int(match) for match in re.findall(r"\bdocPr\b[^>]*\bid=\"(\d+)\"", text))
    return (max(ids) + 1) if ids else 1


def _resolve_hcss_target(
    target: str,
    sets: dict[str, dict[str, Any]],
    html_nodes: dict[str, Any],
) -> list[str]:
    if target in sets:
        definition = sets[target]
        selected = set(_match_hcss_selector(definition.get("select"), html_nodes))
        excluded = set(_match_hcss_selector(definition.get("exclude"), html_nodes))
        return sorted(selected - excluded)
    return _match_hcss_selector(target, html_nodes)


def _hcss_target_allows_empty(target: str, sets: dict[str, dict[str, Any]]) -> bool:
    definition = sets.get(target)
    return bool(definition and definition.get("allowEmpty"))


def _match_hcss_selector(selector: str | None, html_nodes: dict[str, Any]) -> list[str]:
    if not selector:
        return []
    selector = selector.strip()
    selector_parts = _split_hcss_selector_list(selector)
    if len(selector_parts) > 1:
        matched: set[str] = set()
        for selector_part in selector_parts:
            matched.update(_match_hcss_selector(selector_part, html_nodes))
        return sorted(matched)
    if selector in html_nodes:
        return [selector]
    function_match = re.fullmatch(r"([A-Za-z0-9_-]+)\(([^)]*)\)", selector)
    if function_match:
        return _match_hcss_selector_function(function_match.group(1), function_match.group(2), html_nodes)
    attr_matches = re.findall(r"\[([A-Za-z0-9_.:-]+)=\"([^\"]*)\"\]", selector)
    id_matches = re.findall(r"#([A-Za-z0-9_.:-]+)", selector)
    class_matches = re.findall(r"\.([A-Za-z0-9_-]+)", selector)
    stripped = re.sub(r"\[[A-Za-z0-9_.:-]+=\"[^\"]*\"\]", "", selector)
    stripped = re.sub(r"#[A-Za-z0-9_.:-]+", "", stripped)
    stripped = re.sub(r"\.[A-Za-z0-9_-]+", "", stripped).strip()
    if not attr_matches and not id_matches and not class_matches:
        return []
    if stripped and stripped != "*":
        return []
    matched: list[str] = []
    for node_id, node in html_nodes.items():
        classes = set(node.attrs.get("class", "").split())
        if id_matches and node_id not in id_matches and node.attrs.get("id") not in id_matches:
            continue
        if not all(class_name in classes for class_name in class_matches):
            continue
        if all(node.attrs.get(attr) == value for attr, value in attr_matches):
            matched.append(node_id)
    return sorted(matched)


def _match_hcss_selector_function(function_name: str, raw_args: str, html_nodes: dict[str, Any]) -> list[str]:
    args = [_strip_hcss_value(arg.strip()) for arg in raw_args.split(",") if arg.strip()]
    name = function_name.lower()
    matched: list[str] = []
    if name == "id" and len(args) == 1:
        return [args[0]] if args[0] in html_nodes else []
    for node_id, node in html_nodes.items():
        attrs = node.attrs
        if name == "type" and len(args) == 1:
            if attrs.get("data-hdocx-type") == args[0]:
                matched.append(node_id)
        elif name == "style" and len(args) == 1:
            if attrs.get("data-hdocx-style-id") == args[0]:
                matched.append(node_id)
        elif name == "part" and len(args) in {1, 2}:
            if attrs.get("data-hdocx-part") != args[0]:
                continue
            node_type = attrs.get("data-hdocx-type")
            if node_type in {"part", "section"}:
                continue
            if len(args) == 2 and node_type != args[1]:
                continue
            if len(args) == 1 or node_type == args[1]:
                matched.append(node_id)
        elif name == "list" and len(args) in {1, 2}:
            if attrs.get("data-hdocx-num-id") != args[0]:
                continue
            if len(args) == 2 and attrs.get("data-hdocx-ilvl") != args[1]:
                continue
            matched.append(node_id)
    return sorted(matched)


def _split_hcss_selector_list(selector: str) -> list[str]:
    parts: list[str] = []
    start = 0
    paren_depth = 0
    bracket_depth = 0
    quote: str | None = None
    for index, ch in enumerate(selector):
        if quote:
            if ch == quote:
                quote = None
            continue
        if ch in {"'", '"'}:
            quote = ch
            continue
        if ch == "(":
            paren_depth += 1
            continue
        if ch == ")" and paren_depth:
            paren_depth -= 1
            continue
        if ch == "[":
            bracket_depth += 1
            continue
        if ch == "]" and bracket_depth:
            bracket_depth -= 1
            continue
        if ch == "," and paren_depth == 0 and bracket_depth == 0:
            part = selector[start:index].strip()
            if part:
                parts.append(part)
            start = index + 1
    tail = selector[start:].strip()
    if tail:
        parts.append(tail)
    return parts


def _compile_hcss_paragraph_rule(
    target_ids: list[str],
    declarations: list[tuple[str, str]],
    manifest_nodes: dict[str, Any],
    edit_index: int,
) -> tuple[list[dict[str, Any]], list[dict[str, Any]], int]:
    errors: list[dict[str, Any]] = []
    patches: list[dict[str, Any]] = []
    for key, value in declarations:
        if not key.startswith("hdocx-"):
            continue
        prop_name = key.removeprefix("hdocx-")
        if prop_name in {"insert-empty-paragraph-before", "insert-empty-paragraph-after"} or prop_name.startswith(EMPTY_PARAGRAPH_PROPERTY_PREFIX):
            errors.append(
                {
                    "code": "HCSS_PARAGRAPH_STRUCTURE_MODE_REQUIRED",
                    "message": "Empty paragraph insertion must use @hdocx-edit mode(paragraph-structure).",
                    "property": prop_name,
                    "value": value,
                }
            )
    formatting_declarations = [
        (key, value)
        for key, value in declarations
        if not (key.startswith("hdocx-") and key.removeprefix("hdocx-") in PARAGRAPH_STRUCTURAL_PROPERTIES)
    ]
    new_properties, property_errors = _hcss_paragraph_properties(formatting_declarations)
    errors.extend(property_errors)
    manual_break, manual_break_errors = _hcss_manual_page_break_before(declarations)
    errors.extend(manual_break_errors)
    if errors:
        return errors, patches, edit_index
    for node_id in sorted(set(target_ids)):
        node = manifest_nodes.get(node_id)
        if not node or node.get("kind") != "paragraph":
            errors.append({"code": "HCSS_TARGET_KIND_MISMATCH", "message": "paragraph-formatting requires paragraph targets.", "nodeId": node_id})
            continue
        if node.get("lock") != "editable":
            errors.append({"code": "HCSS_TARGET_PROTECTED", "message": "H-CSS matched a protected paragraph.", "nodeId": node_id})
            continue
        if manual_break == "true":
            edit_index += 1
            patches.append(_manual_page_break_patch(edit_index, node_id, node))
        changed = {key: value for key, value in new_properties.items() if node.get("properties", {}).get(key) != value}
        if not changed:
            continue
        edit_index += 1
        patches.append(_paragraph_patch(edit_index, node_id, node, changed))
    return errors, patches, edit_index


def _compile_hcss_paragraph_structure_rule(
    bundle_dir: Path,
    target_ids: list[str],
    declarations: list[tuple[str, str]],
    manifest_nodes: dict[str, Any],
    edit_index: int,
) -> tuple[list[dict[str, Any]], list[dict[str, Any]], int]:
    errors: list[dict[str, Any]] = []
    patches: list[dict[str, Any]] = []
    structure, structure_errors = _hcss_paragraph_structure_properties(declarations)
    errors.extend(structure_errors)
    if errors:
        return errors, patches, edit_index
    for node_id in sorted(set(target_ids)):
        node = manifest_nodes.get(node_id)
        if not node or node.get("kind") != "paragraph":
            errors.append({"code": "HCSS_TARGET_KIND_MISMATCH", "message": "paragraph-structure requires paragraph targets.", "nodeId": node_id})
            continue
        if node.get("lock") != "editable":
            errors.append({"code": "HCSS_TARGET_PROTECTED", "message": "H-CSS matched a protected paragraph.", "nodeId": node_id})
            continue
        if structure.get("manual-page-break-before") == "true":
            edit_index += 1
            patches.append(_manual_page_break_patch(edit_index, node_id, node))
        if structure.get("insert-empty-paragraph-before") == "true":
            noop = _empty_paragraph_insert_noop(bundle_dir, node, "before", structure.get("emptyProperties", {}))
            if noop:
                patches.append(_empty_paragraph_noop(node_id, node, "before", noop))
            else:
                edit_index += 1
                patches.append(_empty_paragraph_patch(edit_index, node_id, node, "before", structure.get("emptyProperties", {})))
        if structure.get("insert-empty-paragraph-after") == "true":
            noop = _empty_paragraph_insert_noop(bundle_dir, node, "after", structure.get("emptyProperties", {}))
            if noop:
                patches.append(_empty_paragraph_noop(node_id, node, "after", noop))
            else:
                edit_index += 1
                patches.append(_empty_paragraph_patch(edit_index, node_id, node, "after", structure.get("emptyProperties", {})))
    return errors, patches, edit_index


def _compile_hcss_image_formatting_rule(
    target_ids: list[str],
    declarations: list[tuple[str, str]],
    manifest_nodes: dict[str, Any],
    edit_index: int,
) -> tuple[list[dict[str, Any]], list[dict[str, Any]], int]:
    errors: list[dict[str, Any]] = []
    patches: list[dict[str, Any]] = []
    requested_image_props, requested_paragraph_props, property_errors = _hcss_image_formatting_properties(declarations)
    errors.extend(property_errors)
    if errors:
        return errors, patches, edit_index

    run_ids: list[str] = []
    for node_id in target_ids:
        node = manifest_nodes.get(node_id)
        if not node:
            errors.append({"code": "HCSS_TARGET_NOT_FOUND", "message": "H-CSS target is not in manifest.", "nodeId": node_id})
            continue
        if node.get("kind") == "run":
            run_ids.append(node_id)
        elif node.get("kind") == "paragraph":
            for child_id in node.get("children", []):
                child = manifest_nodes.get(child_id, {})
                if child.get("kind") == "run" and child.get("objectKind") == "drawing":
                    run_ids.append(child_id)
        else:
            errors.append({"code": "HCSS_TARGET_KIND_MISMATCH", "message": "image-formatting requires drawing run or paragraph targets.", "nodeId": node_id})
    for run_id in sorted(set(run_ids)):
        node = manifest_nodes.get(run_id, {})
        if node.get("kind") != "run" or node.get("objectKind") != "drawing":
            errors.append({"code": "HCSS_TARGET_KIND_MISMATCH", "message": "image-formatting matched a non-image run.", "nodeId": run_id})
            continue
        old_image_props = node.get("objectProperties", {})
        changed_image_props = {
            key: value
            for key, value in requested_image_props.items()
            if old_image_props.get(key) != value
        }
        if changed_image_props:
            edit_index += 1
            patches.append(
                {
                    "id": f"patch-{edit_index:06d}",
                    "operation": "patch-drawing-properties",
                    "partPath": node["partPath"],
                    "entryName": node["partPath"].lstrip("/"),
                    "nodeId": run_id,
                    "locator": node["locator"],
                    "expectedOldHash": node["hash"],
                    "oldProperties": old_image_props,
                    "newProperties": changed_image_props,
                }
            )
        if requested_paragraph_props:
            paragraph_id = node.get("parent")
            paragraph = manifest_nodes.get(paragraph_id or "", {})
            if paragraph.get("kind") != "paragraph":
                errors.append({"code": "HCSS_IMAGE_PARENT_NOT_FOUND", "message": "Image run parent paragraph was not found.", "nodeId": run_id})
                continue
            changed_paragraph_props = {
                key: value
                for key, value in requested_paragraph_props.items()
                if paragraph.get("properties", {}).get(key) != value
            }
            if changed_paragraph_props:
                edit_index += 1
                patches.append(_paragraph_patch(edit_index, paragraph_id, paragraph, changed_paragraph_props))
    return errors, patches, edit_index


def _compile_hcss_all_runs_rule(
    target_ids: list[str],
    declarations: list[tuple[str, str]],
    manifest_nodes: dict[str, Any],
    edit_index: int,
) -> tuple[list[dict[str, Any]], list[dict[str, Any]], int]:
    errors: list[dict[str, Any]] = []
    patches: list[dict[str, Any]] = []
    new_properties, property_errors = _hcss_run_properties(declarations)
    errors.extend(property_errors)
    if errors:
        return errors, patches, edit_index
    run_ids: list[str] = []
    for node_id in target_ids:
        node = manifest_nodes.get(node_id)
        if not node:
            errors.append({"code": "HCSS_TARGET_NOT_FOUND", "message": "H-CSS target is not in manifest.", "nodeId": node_id})
            continue
        if node.get("kind") == "run":
            run_ids.append(node_id)
        elif node.get("kind") == "paragraph":
            run_ids.extend(node.get("children", []))
        else:
            errors.append({"code": "HCSS_TARGET_KIND_MISMATCH", "message": "all-runs requires run or paragraph targets.", "nodeId": node_id})
    for run_id in sorted(set(run_ids)):
        node = manifest_nodes.get(run_id)
        if not node:
            continue
        if node.get("lock") != "editable":
            errors.append({"code": "HCSS_TARGET_PROTECTED", "message": "H-CSS matched a protected run.", "nodeId": run_id})
            continue
        changed = {key: value for key, value in new_properties.items() if node.get("properties", {}).get(key) != value}
        if not changed:
            continue
        edit_index += 1
        patches.append(_run_property_patch(edit_index, run_id, node, changed))
    return errors, patches, edit_index


def _compile_hcss_direct_rule(
    target_ids: list[str],
    declarations: list[tuple[str, str]],
    manifest_nodes: dict[str, Any],
    edit_index: int,
) -> tuple[list[dict[str, Any]], list[dict[str, Any]], int]:
    prop_names = {_canonical_hcss_property_name(key.removeprefix("hdocx-")) for key, _ in declarations if key.startswith("hdocx-")}
    if prop_names and prop_names.issubset(PARAGRAPH_PROPERTY_ATTRS):
        return _compile_hcss_paragraph_rule(target_ids, declarations, manifest_nodes, edit_index)
    if prop_names and prop_names.issubset(RUN_PROPERTY_ATTRS):
        return _compile_hcss_all_runs_rule(target_ids, declarations, manifest_nodes, edit_index)
    return (
        [{"code": "HCSS_DIRECT_FORMATTING_AMBIGUOUS", "message": "direct-formatting cannot mix paragraph and run properties."}],
        [],
        edit_index,
    )


def _compile_hcss_style_definition_rule(
    target_ids: list[str],
    declarations: list[tuple[str, str]],
    manifest_nodes: dict[str, Any],
    manifest: dict[str, Any],
    edit_index: int,
) -> tuple[list[dict[str, Any]], list[dict[str, Any]], int]:
    errors: list[dict[str, Any]] = []
    patches: list[dict[str, Any]] = []
    if "/word/styles.xml" not in manifest.get("parts", {}):
        return (
            [{"code": "HCSS_STYLES_PART_MISSING", "message": "style-definition mode requires word/styles.xml."}],
            [],
            edit_index,
        )
    paragraph_props, paragraph_errors = _hcss_paragraph_properties(
        [
            (key, value)
            for key, value in declarations
            if key.startswith("hdocx-")
            and _canonical_paragraph_property_name(key.removeprefix("hdocx-")) in PARAGRAPH_PROPERTY_ATTRS
        ]
    )
    run_props, run_errors = _hcss_run_properties(
        [
            (key, value)
            for key, value in declarations
            if key.startswith("hdocx-")
            and _canonical_run_property_name(key.removeprefix("hdocx-")) in RUN_PROPERTY_ATTRS
        ]
    )
    known_props = set(PARAGRAPH_PROPERTY_ATTRS) | set(RUN_PROPERTY_ATTRS)
    for key, _ in declarations:
        if not key.startswith("hdocx-") or _canonical_hcss_property_name(key.removeprefix("hdocx-")) not in known_props:
            errors.append({"code": "HCSS_STYLE_PROPERTY_UNSUPPORTED", "message": "Unsupported style-definition property.", "property": key})
    errors.extend(paragraph_errors)
    errors.extend(run_errors)
    if errors:
        return errors, patches, edit_index

    style_ids: set[str] = set()
    for node_id in target_ids:
        node = manifest_nodes.get(node_id)
        if not node:
            errors.append({"code": "HCSS_TARGET_NOT_FOUND", "message": "H-CSS target is not in manifest.", "nodeId": node_id})
            continue
        if node.get("kind") != "paragraph":
            errors.append({"code": "HCSS_TARGET_KIND_MISMATCH", "message": "style-definition mode requires paragraph targets.", "nodeId": node_id})
            continue
        style_id = node.get("styleId")
        if not style_id:
            errors.append({"code": "HCSS_TARGET_HAS_NO_STYLE", "message": "Matched paragraph has no styleId.", "nodeId": node_id})
            continue
        style_ids.add(style_id)
    if errors:
        return errors, patches, edit_index
    for style_id in sorted(style_ids):
        edit_index += 1
        patches.append(
            {
                "id": f"patch-{edit_index:06d}",
                "operation": "patch-style",
                "partPath": "/word/styles.xml",
                "entryName": "word/styles.xml",
                "styleId": style_id,
                "newProperties": {
                    "paragraph": paragraph_props,
                    "run": run_props,
                },
            }
        )
    return errors, patches, edit_index


def _compile_hcss_numbering_definition_rule(
    target_ids: list[str],
    declarations: list[tuple[str, str]],
    manifest_nodes: dict[str, Any],
    manifest: dict[str, Any],
    edit_index: int,
) -> tuple[list[dict[str, Any]], list[dict[str, Any]], int]:
    errors: list[dict[str, Any]] = []
    patches: list[dict[str, Any]] = []
    if "/word/numbering.xml" not in manifest.get("parts", {}):
        return (
            [{"code": "HCSS_NUMBERING_PART_MISSING", "message": "numbering-definition mode requires word/numbering.xml."}],
            [],
            edit_index,
        )
    new_properties, property_errors = _hcss_numbering_properties(declarations)
    errors.extend(property_errors)
    if errors:
        return errors, patches, edit_index

    levels: dict[tuple[str, str], dict[str, str]] = {}
    for node_id in target_ids:
        node = manifest_nodes.get(node_id)
        if not node:
            errors.append({"code": "HCSS_TARGET_NOT_FOUND", "message": "H-CSS target is not in manifest.", "nodeId": node_id})
            continue
        if node.get("kind") != "paragraph":
            errors.append({"code": "HCSS_TARGET_KIND_MISMATCH", "message": "numbering-definition mode requires paragraph targets.", "nodeId": node_id})
            continue
        numbering = node.get("numbering", {})
        abstract_num_id = numbering.get("abstractNumId")
        ilvl = numbering.get("ilvl")
        if abstract_num_id is None or ilvl is None:
            errors.append({"code": "HCSS_TARGET_HAS_NO_NUMBERING", "message": "Matched paragraph has no resolved numbering level.", "nodeId": node_id})
            continue
        levels[(str(abstract_num_id), str(ilvl))] = numbering
    if errors:
        return errors, patches, edit_index
    for abstract_num_id, ilvl in sorted(levels):
        old_numbering = levels[(abstract_num_id, ilvl)]
        changed = {
            key: value
            for key, value in new_properties.items()
            if _numbering_old_value(old_numbering, key) != value
        }
        if not changed:
            continue
        edit_index += 1
        patches.append(
            {
                "id": f"patch-{edit_index:06d}",
                "operation": "patch-numbering-level",
                "partPath": "/word/numbering.xml",
                "entryName": "word/numbering.xml",
                "abstractNumId": abstract_num_id,
                "ilvl": ilvl,
                "oldProperties": old_numbering,
                "newProperties": changed,
            }
        )
    return errors, patches, edit_index


def _compile_hcss_paragraph_style_rule(
    target_ids: list[str],
    declarations: list[tuple[str, str]],
    manifest_nodes: dict[str, Any],
    manifest: dict[str, Any],
    edit_index: int,
) -> tuple[list[dict[str, Any]], list[dict[str, Any]], int]:
    errors: list[dict[str, Any]] = []
    patches: list[dict[str, Any]] = []
    style_id = None
    for key, value in declarations:
        if key != "hdocx-style-id":
            errors.append(
                {
                    "code": "HCSS_PARAGRAPH_STYLE_PROPERTY_UNSUPPORTED",
                    "message": "paragraph-style mode only supports hdocx-style-id.",
                    "property": key,
                }
            )
            continue
        style_id = _strip_hcss_value(value)
    if not style_id:
        errors.append(
            {
                "code": "HCSS_PARAGRAPH_STYLE_ID_MISSING",
                "message": "paragraph-style mode requires hdocx-style-id.",
            }
        )
    styles = manifest.get("styles", {})
    if style_id and style_id not in styles:
        errors.append(
            {
                "code": "HCSS_PARAGRAPH_STYLE_UNKNOWN",
                "message": "Requested style id does not exist in the manifest.",
                "styleId": style_id,
            }
        )
    if errors:
        return errors, patches, edit_index
    for node_id in target_ids:
        node = manifest_nodes.get(node_id)
        if not node or node.get("kind") != "paragraph":
            errors.append({"code": "HCSS_TARGET_KIND_MISMATCH", "message": "paragraph-style mode requires paragraph targets.", "nodeId": node_id})
            continue
        if node.get("lock") != "editable":
            errors.append({"code": "HCSS_TARGET_PROTECTED", "message": "H-CSS matched a protected paragraph.", "nodeId": node_id})
            continue
        if node.get("styleId") == style_id:
            continue
        edit_index += 1
        patches.append(
            {
                "id": f"patch-{edit_index:06d}",
                "operation": "patch-paragraph-style",
                "partPath": node["partPath"],
                "entryName": node["partPath"].lstrip("/"),
                "nodeId": node_id,
                "locator": node["locator"],
                "oldStyleId": node.get("styleId"),
                "newStyleId": style_id,
            }
        )
    return errors, patches, edit_index


def _compile_hcss_paragraph_numbering_rule(
    target_ids: list[str],
    declarations: list[tuple[str, str]],
    manifest_nodes: dict[str, Any],
    manifest: dict[str, Any],
    edit_index: int,
) -> tuple[list[dict[str, Any]], list[dict[str, Any]], int]:
    errors: list[dict[str, Any]] = []
    patches: list[dict[str, Any]] = []
    list_id = None
    ilvl = "0"
    for key, value in declarations:
        stripped = _strip_hcss_value(value)
        if key == "hdocx-list-id":
            list_id = stripped
        elif key == "hdocx-ilvl":
            if not stripped.isdigit():
                errors.append({"code": "HCSS_PARAGRAPH_NUMBERING_INVALID_ILVL", "message": "hdocx-ilvl must be numeric.", "value": stripped})
            else:
                ilvl = str(int(stripped))
        else:
            errors.append({"code": "HCSS_PARAGRAPH_NUMBERING_PROPERTY_UNSUPPORTED", "message": "Unsupported paragraph-numbering property.", "property": key})
    if not list_id:
        errors.append({"code": "HCSS_PARAGRAPH_NUMBERING_LIST_MISSING", "message": "paragraph-numbering mode requires hdocx-list-id."})
    target_list = _resolve_hcss_list_id(list_id, manifest) if list_id else None
    if list_id and target_list is None:
        errors.append({"code": "HCSS_PARAGRAPH_NUMBERING_UNKNOWN_LIST", "message": "Unknown list id or numId.", "listId": list_id})
    if errors:
        return errors, patches, edit_index
    for node_id in target_ids:
        node = manifest_nodes.get(node_id)
        if not node or node.get("kind") != "paragraph":
            errors.append({"code": "HCSS_TARGET_KIND_MISMATCH", "message": "paragraph-numbering mode requires paragraph targets.", "nodeId": node_id})
            continue
        if node.get("lock") != "editable":
            errors.append({"code": "HCSS_TARGET_PROTECTED", "message": "H-CSS matched a protected paragraph.", "nodeId": node_id})
            continue
        old_numbering = node.get("numbering", {})
        if old_numbering.get("numId") == target_list["numId"] and old_numbering.get("ilvl") == ilvl:
            continue
        edit_index += 1
        patches.append(
            {
                "id": f"patch-{edit_index:06d}",
                "operation": "patch-paragraph-numbering",
                "partPath": node["partPath"],
                "entryName": node["partPath"].lstrip("/"),
                "nodeId": node_id,
                "locator": node["locator"],
                "oldNumbering": old_numbering,
                "numId": target_list["numId"],
                "ilvl": ilvl,
            }
        )
    return errors, patches, edit_index


def _compile_hcss_comment_text_rule(
    target_ids: list[str],
    declarations: list[tuple[str, str]],
    manifest_nodes: dict[str, Any],
    edit_index: int,
) -> tuple[list[dict[str, Any]], list[dict[str, Any]], int]:
    errors: list[dict[str, Any]] = []
    patches: list[dict[str, Any]] = []
    new_text = None
    for key, value in declarations:
        if key != "hdocx-text":
            errors.append({"code": "HCSS_COMMENT_TEXT_PROPERTY_UNSUPPORTED", "message": "comment-text mode only supports hdocx-text.", "property": key})
            continue
        new_text = _strip_hcss_value(value)
    if new_text is None:
        errors.append({"code": "HCSS_COMMENT_TEXT_MISSING", "message": "comment-text mode requires hdocx-text."})
    if errors:
        return errors, patches, edit_index
    for node_id in target_ids:
        node = manifest_nodes.get(node_id)
        if not node or node.get("kind") != "protected" or node.get("protectedKind") != "comment":
            errors.append({"code": "HCSS_TARGET_KIND_MISMATCH", "message": "comment-text mode requires protected comment targets.", "nodeId": node_id})
            continue
        if node.get("partPath") != "/word/comments.xml":
            errors.append({"code": "HCSS_COMMENT_TEXT_PART_UNSUPPORTED", "message": "comment-text mode requires /word/comments.xml targets.", "nodeId": node_id})
            continue
        edit_index += 1
        patches.append(
            {
                "id": f"patch-{edit_index:06d}",
                "operation": "patch-comment-text",
                "partPath": "/word/comments.xml",
                "entryName": "word/comments.xml",
                "nodeId": node_id,
                "locator": node.get("locator", {}),
                "oldText": node.get("text"),
                "newText": new_text,
            }
        )
    return errors, patches, edit_index


def _compile_hcss_revision_action_rule(
    target_ids: list[str],
    declarations: list[tuple[str, str]],
    manifest_nodes: dict[str, Any],
    edit_index: int,
) -> tuple[list[dict[str, Any]], list[dict[str, Any]], int]:
    errors: list[dict[str, Any]] = []
    patches: list[dict[str, Any]] = []
    action = None
    for key, value in declarations:
        if key != "hdocx-action":
            errors.append({"code": "HCSS_REVISION_ACTION_PROPERTY_UNSUPPORTED", "message": "revision-action mode only supports hdocx-action.", "property": key})
            continue
        action = _strip_hcss_value(value)
    if action not in {"accept", "reject"}:
        errors.append({"code": "HCSS_REVISION_ACTION_INVALID", "message": "hdocx-action must be accept or reject.", "action": action})
    if errors:
        return errors, patches, edit_index
    for node_id in target_ids:
        node = manifest_nodes.get(node_id)
        protected_kind = node.get("protectedKind") if node else None
        if not node or node.get("kind") != "protected" or protected_kind not in {"revision-insert", "revision-delete"}:
            errors.append({"code": "HCSS_TARGET_KIND_MISMATCH", "message": "revision-action mode requires revision protected targets.", "nodeId": node_id})
            continue
        edit_index += 1
        patches.append(
            {
                "id": f"patch-{edit_index:06d}",
                "operation": "patch-revision-action",
                "partPath": node["partPath"],
                "entryName": node["partPath"].lstrip("/"),
                "nodeId": node_id,
                "locator": node.get("locator", {}),
                "revisionKind": protected_kind,
                "action": action,
            }
        )
    return errors, patches, edit_index


def _compile_hcss_equation_omml_rule(
    bundle_dir: Path,
    target_ids: list[str],
    declarations: list[tuple[str, str]],
    manifest_nodes: dict[str, Any],
    edit_index: int,
) -> tuple[list[dict[str, Any]], list[dict[str, Any]], int]:
    errors: list[dict[str, Any]] = []
    patches: list[dict[str, Any]] = []
    source = None
    for key, value in declarations:
        if key != "hdocx-omml-source":
            errors.append({"code": "HCSS_EQUATION_OMML_PROPERTY_UNSUPPORTED", "message": "equation-omml mode only supports hdocx-omml-source.", "property": key})
            continue
        source = _strip_hcss_value(value)
    if not source:
        errors.append({"code": "HCSS_EQUATION_OMML_SOURCE_MISSING", "message": "equation-omml mode requires hdocx-omml-source."})
    source_path = _resolve_equation_source(bundle_dir, source, errors) if source else None
    new_omml = None
    if source_path is not None:
        try:
            new_omml = source_path.read_text(encoding="utf-8").strip()
            _validate_omml_fragment(new_omml)
        except UnicodeDecodeError:
            errors.append({"code": "HCSS_EQUATION_OMML_SOURCE_ENCODING", "message": "OMML source must be UTF-8 text.", "source": source})
        except ET.ParseError as exc:
            errors.append({"code": "HCSS_EQUATION_OMML_PARSE_ERROR", "message": str(exc), "source": source})
        except ValueError as exc:
            errors.append({"code": "HCSS_EQUATION_OMML_INVALID_ROOT", "message": str(exc), "source": source})
    if errors:
        return errors, patches, edit_index
    for node_id in target_ids:
        node = manifest_nodes.get(node_id)
        if not node or node.get("kind") != "protected" or node.get("protectedKind") != "equation":
            errors.append({"code": "HCSS_TARGET_KIND_MISMATCH", "message": "equation-omml mode requires protected equation targets.", "nodeId": node_id})
            continue
        locator = node.get("locator", {})
        if locator.get("paragraphIndex") is None or locator.get("protectedInParagraphIndex") is None:
            errors.append({"code": "HCSS_EQUATION_LOCATOR_MISSING", "message": "Equation target does not have a resolvable paragraph locator.", "nodeId": node_id})
            continue
        edit_index += 1
        patches.append(
            {
                "id": f"patch-{edit_index:06d}",
                "operation": "patch-equation-omml",
                "partPath": node["partPath"],
                "entryName": node["partPath"].lstrip("/"),
                "nodeId": node_id,
                "locator": locator,
                "expectedOldHash": node.get("hash"),
                "newOmml": new_omml,
                "sourcePath": str(source_path.relative_to(bundle_dir.resolve()).as_posix()) if source_path else None,
            }
        )
    return errors, patches, edit_index


def _resolve_equation_source(bundle_dir: Path, source: str | None, errors: list[dict[str, Any]]) -> Path | None:
    if not source:
        return None
    source_rel = Path(source)
    if source_rel.is_absolute() or any(part in {"", ".", ".."} for part in source_rel.parts):
        errors.append(
            {
                "code": "HCSS_EQUATION_OMML_UNSAFE_SOURCE",
                "message": "OMML source must be a safe relative path inside the H-DOCX bundle.",
                "source": source,
            }
        )
        return None
    source_path = (bundle_dir / source_rel).resolve()
    bundle_root = bundle_dir.resolve()
    try:
        source_path.relative_to(bundle_root)
    except ValueError:
        errors.append(
            {
                "code": "HCSS_EQUATION_OMML_UNSAFE_SOURCE",
                "message": "OMML source resolved outside the H-DOCX bundle.",
                "source": source,
            }
        )
        return None
    if not source_path.is_file():
        errors.append(
            {
                "code": "HCSS_EQUATION_OMML_SOURCE_MISSING",
                "message": "OMML source file was not found.",
                "source": source,
            }
        )
        return None
    return source_path


def _validate_omml_fragment(text: str) -> None:
    root = ET.fromstring(text)
    if root.tag not in {f"{{{M_NS}}}oMath", f"{{{M_NS}}}oMathPara"}:
        raise ValueError("OMML root must be m:oMath or m:oMathPara.")


def _resolve_hcss_list_id(list_id: str | None, manifest: dict[str, Any]) -> dict[str, str] | None:
    if not list_id:
        return None
    created = manifest.get("createdLists", {}).get(list_id)
    if created:
        return {"numId": created["numId"], "abstractNumId": created["abstractNumId"]}
    nums = manifest.get("numbering", {}).get("nums", {})
    if list_id in nums:
        return {"numId": list_id, "abstractNumId": str(nums[list_id].get("abstractNumId"))}
    return None


def _hcss_numbering_properties(declarations: list[tuple[str, str]]) -> tuple[dict[str, str], list[dict[str, Any]]]:
    properties: dict[str, str] = {}
    errors: list[dict[str, Any]] = []
    for key, value in declarations:
        if not key.startswith("hdocx-"):
            errors.append({"code": "HCSS_UNSUPPORTED_DECLARATION", "message": "Only hdocx-* declarations are allowed.", "property": key})
            continue
        prop_name = key.removeprefix("hdocx-")
        if prop_name not in NUMBERING_PROPERTY_NAMES:
            errors.append({"code": "HCSS_NUMBERING_PROPERTY_UNSUPPORTED", "message": "Unsupported numbering property.", "property": prop_name})
            continue
        try:
            properties[prop_name] = _normalize_numbering_property(prop_name, value)
        except ValueError as exc:
            errors.append({"code": "HCSS_INVALID_NUMBERING_PROPERTY", "message": str(exc), "property": prop_name, "value": value})
    return properties, errors


def _hcss_list_level_properties(declarations: list[tuple[str, str]]) -> tuple[dict[str, dict[str, str]], list[dict[str, Any]]]:
    levels: dict[str, dict[str, str]] = {"0": {}}
    errors: list[dict[str, Any]] = []
    for key, value in declarations:
        if key == "__parse_error__":
            errors.append({"code": "HCSS_DECLARATION_PARSE_ERROR", "message": "Invalid declaration.", "value": value})
            continue
        if not key.startswith("hdocx-"):
            errors.append({"code": "HCSS_UNSUPPORTED_DECLARATION", "message": "Only hdocx-* declarations are allowed.", "property": key})
            continue
        prop_name = key.removeprefix("hdocx-")
        ilvl = "0"
        level_match = re.fullmatch(r"level-([0-8])-(.+)", prop_name)
        if level_match:
            ilvl, prop_name = level_match.groups()
        if prop_name not in NUMBERING_PROPERTY_NAMES:
            errors.append({"code": "HCSS_NUMBERING_PROPERTY_UNSUPPORTED", "message": "Unsupported numbering property.", "property": prop_name})
            continue
        try:
            levels.setdefault(str(int(ilvl)), {})[prop_name] = _normalize_numbering_property(prop_name, value)
        except ValueError as exc:
            errors.append({"code": "HCSS_INVALID_NUMBERING_PROPERTY", "message": str(exc), "property": prop_name, "value": value})
    if errors:
        return levels, errors
    for ilvl, props in levels.items():
        display_level = int(ilvl) + 1
        props.setdefault("num-format", "decimal")
        props.setdefault("level-text", f"%{display_level}.")
        props.setdefault("start", "1")
    return dict(sorted(levels.items(), key=lambda item: int(item[0]))), errors


def _normalize_numbering_property(prop_name: str, raw_value: str) -> str:
    value = _strip_hcss_value(raw_value)
    if prop_name == "num-format":
        if not re.fullmatch(r"[A-Za-z][A-Za-z0-9]*", value):
            raise ValueError("num-format must be an OOXML numbering format token.")
        return value
    if prop_name == "level-text":
        if value == "":
            raise ValueError("level-text must not be empty.")
        return value
    if prop_name == "number-suffix":
        if value not in {"tab", "space", "nothing"}:
            raise ValueError("number-suffix must be tab, space, or nothing.")
        return value
    if prop_name == "start":
        if not value.isdigit() or int(value) <= 0:
            raise ValueError("start must be a positive integer.")
        return str(int(value))
    if prop_name in {"num-indent-left", "num-indent-hanging", "num-indent-first-line"}:
        if not value.isdigit() or int(value) < 0:
            raise ValueError(f"{prop_name} must be a non-negative twip integer.")
        return str(int(value))
    raise ValueError(f"Unsupported numbering property: {prop_name}")


def _numbering_old_value(numbering: dict[str, Any], prop_name: str) -> str | None:
    mapping = {
        "num-format": "numFmt",
        "level-text": "lvlText",
        "number-suffix": "suffix",
        "start": "start",
        "num-indent-left": "indent.left",
        "num-indent-hanging": "indent.hanging",
        "num-indent-first-line": "indent.firstLine",
    }
    return numbering.get(mapping[prop_name])


def _hcss_run_properties(declarations: list[tuple[str, str]]) -> tuple[dict[str, str | None], list[dict[str, Any]]]:
    properties: dict[str, str | None] = {}
    errors: list[dict[str, Any]] = []
    for key, value in declarations:
        if not key.startswith("hdocx-"):
            errors.append({"code": "HCSS_UNSUPPORTED_DECLARATION", "message": "Only hdocx-* declarations are allowed.", "property": key})
            continue
        prop_name = _canonical_run_property_name(key.removeprefix("hdocx-"))
        if prop_name not in RUN_PROPERTY_ATTRS:
            errors.append({"code": "HCSS_RUN_PROPERTY_UNSUPPORTED", "message": "Unsupported run property for all-runs mode.", "property": prop_name})
            continue
        try:
            properties[prop_name] = _normalize_run_property(prop_name, value)
        except ValueError as exc:
            errors.append({"code": "HCSS_INVALID_RUN_PROPERTY", "message": str(exc), "property": prop_name, "value": value})
    return properties, errors


def _hcss_paragraph_properties(declarations: list[tuple[str, str]]) -> tuple[dict[str, str | None], list[dict[str, Any]]]:
    properties: dict[str, str | None] = {}
    errors: list[dict[str, Any]] = []
    for key, value in declarations:
        if not key.startswith("hdocx-"):
            errors.append({"code": "HCSS_UNSUPPORTED_DECLARATION", "message": "Only hdocx-* declarations are allowed.", "property": key})
            continue
        raw_prop = key.removeprefix("hdocx-")
        prop_name = _canonical_paragraph_property_name(raw_prop)
        if prop_name not in PARAGRAPH_PROPERTY_ATTRS:
            errors.append({"code": "HCSS_PARAGRAPH_PROPERTY_UNSUPPORTED", "message": "Unsupported paragraph property for paragraph-formatting mode.", "property": prop_name})
            continue
        try:
            if raw_prop in {"line-spacing-exact", "line-spacingExact"} and not _strip_hcss_value(value).endswith("pt"):
                raise ValueError("line-spacing-exact must use pt units.")
            properties[prop_name] = _normalize_paragraph_property(prop_name, value)
        except ValueError as exc:
            errors.append({"code": "HCSS_INVALID_PARAGRAPH_PROPERTY", "message": str(exc), "property": prop_name, "value": value})
    return properties, errors


def _hcss_manual_page_break_before(declarations: list[tuple[str, str]]) -> tuple[str | None, list[dict[str, Any]]]:
    errors: list[dict[str, Any]] = []
    result: str | None = None
    for key, value in declarations:
        if not key.startswith("hdocx-"):
            continue
        prop_name = key.removeprefix("hdocx-")
        if prop_name != "manual-page-break-before":
            continue
        try:
            result = _normalize_hcss_bool(prop_name, value)
        except ValueError as exc:
            errors.append({"code": "HCSS_INVALID_PARAGRAPH_PROPERTY", "message": str(exc), "property": prop_name, "value": value})
    return result, errors


def _hcss_paragraph_structure_properties(declarations: list[tuple[str, str]]) -> tuple[dict[str, Any], list[dict[str, Any]]]:
    errors: list[dict[str, Any]] = []
    result: dict[str, Any] = {"emptyProperties": {}}
    saw_empty_insert = False
    for key, value in declarations:
        if key == "__parse_error__":
            errors.append({"code": "HCSS_DECLARATION_PARSE_ERROR", "message": "Invalid declaration.", "value": value})
            continue
        if not key.startswith("hdocx-"):
            errors.append({"code": "HCSS_DECLARATION_UNPREFIXED", "message": "Only hdocx-* declarations are accepted.", "property": key})
            continue
        prop_name = key.removeprefix("hdocx-")
        if prop_name in PARAGRAPH_STRUCTURAL_PROPERTIES:
            try:
                normalized = _normalize_hcss_bool(prop_name, value)
            except ValueError as exc:
                errors.append({"code": "HCSS_INVALID_PARAGRAPH_STRUCTURE_PROPERTY", "message": str(exc), "property": prop_name, "value": value})
                continue
            result[prop_name] = normalized
            if prop_name in {"insert-empty-paragraph-before", "insert-empty-paragraph-after"} and normalized == "true":
                saw_empty_insert = True
            continue
        if prop_name.startswith(EMPTY_PARAGRAPH_PROPERTY_PREFIX):
            raw_empty_prop = prop_name.removeprefix(EMPTY_PARAGRAPH_PROPERTY_PREFIX)
            empty_prop = _canonical_empty_paragraph_property_name(raw_empty_prop)
            try:
                normalized = _normalize_empty_paragraph_property(raw_empty_prop, value)
            except ValueError as exc:
                errors.append({"code": "HCSS_INVALID_EMPTY_PARAGRAPH_PROPERTY", "message": str(exc), "property": prop_name, "value": value})
                continue
            result["emptyProperties"][empty_prop] = normalized
            continue
        errors.append({"code": "HCSS_PARAGRAPH_STRUCTURE_PROPERTY_UNSUPPORTED", "message": "Unsupported property for paragraph-structure mode.", "property": prop_name})
    if result["emptyProperties"] and not saw_empty_insert:
        errors.append(
            {
                "code": "HCSS_EMPTY_PARAGRAPH_FORMAT_WITHOUT_INSERT",
                "message": "empty-paragraph-* declarations require insert-empty-paragraph-before or insert-empty-paragraph-after to be true.",
            }
        )
    return result, errors


def _hcss_image_formatting_properties(
    declarations: list[tuple[str, str]]
) -> tuple[dict[str, str | None], dict[str, str | None], list[dict[str, Any]]]:
    errors: list[dict[str, Any]] = []
    image_props: dict[str, str | None] = {}
    paragraph_props: dict[str, str | None] = {}
    for key, value in declarations:
        if key == "__parse_error__":
            errors.append({"code": "HCSS_DECLARATION_PARSE_ERROR", "message": "Invalid declaration.", "value": value})
            continue
        if not key.startswith("hdocx-"):
            errors.append({"code": "HCSS_DECLARATION_UNPREFIXED", "message": "Only hdocx-* declarations are accepted.", "property": key})
            continue
        raw_prop = key.removeprefix("hdocx-")
        image_prop = _canonical_image_property_name(raw_prop)
        if image_prop in IMAGE_FORMATTING_PROPERTIES:
            try:
                normalized = _normalize_image_hcss_property(image_prop, value, old_value="1" if image_prop in {"width-emu", "height-emu"} else "")
            except ValueError as exc:
                errors.append({"code": "HCSS_INVALID_IMAGE_PROPERTY", "message": str(exc), "property": raw_prop, "value": value})
                continue
            image_props[image_prop] = normalized
            continue
        paragraph_prop = _canonical_image_paragraph_property_name(raw_prop)
        if paragraph_prop in PARAGRAPH_PROPERTY_ATTRS and raw_prop.startswith(IMAGE_PARAGRAPH_PROPERTY_PREFIX):
            try:
                if raw_prop in {"paragraph-line-spacing-exact", "paragraph-line-spacingExact"} and not _strip_hcss_value(value).endswith("pt"):
                    raise ValueError("paragraph-line-spacing-exact must use pt units.")
                normalized = _normalize_paragraph_property(paragraph_prop, value)
            except ValueError as exc:
                errors.append({"code": "HCSS_INVALID_IMAGE_PARAGRAPH_PROPERTY", "message": str(exc), "property": raw_prop, "value": value})
                continue
            paragraph_props[paragraph_prop] = normalized
            continue
        errors.append({"code": "HCSS_IMAGE_PROPERTY_UNSUPPORTED", "message": "Unsupported property for image-formatting mode.", "property": raw_prop})
    return image_props, paragraph_props, errors


def _paragraph_patch(edit_index: int, node_id: str, node: dict[str, Any], changed: dict[str, str | None]) -> dict[str, Any]:
    return {
        "id": f"patch-{edit_index:06d}",
        "operation": "patch-paragraph",
        "partPath": node["partPath"],
        "entryName": node["partPath"].lstrip("/"),
        "nodeId": node_id,
        "locator": node["locator"],
        "expectedOldHash": node["hash"],
        "oldProperties": node.get("properties", {}),
        "newProperties": changed,
    }


def _manual_page_break_patch(edit_index: int, node_id: str, node: dict[str, Any]) -> dict[str, Any]:
    return {
        "id": f"patch-{edit_index:06d}",
        "operation": "insert-manual-page-break-before",
        "partPath": node["partPath"],
        "entryName": node["partPath"].lstrip("/"),
        "nodeId": node_id,
        "locator": node["locator"],
        "expectedOldHash": node["hash"],
    }


def _empty_paragraph_patch(
    edit_index: int,
    node_id: str,
    node: dict[str, Any],
    position: str,
    empty_properties: dict[str, str],
) -> dict[str, Any]:
    return {
        "id": f"patch-{edit_index:06d}",
        "operation": f"insert-empty-paragraph-{position}",
        "partPath": node["partPath"],
        "entryName": node["partPath"].lstrip("/"),
        "nodeId": node_id,
        "locator": node["locator"],
        "expectedOldHash": node["hash"],
        "newProperties": dict(empty_properties),
    }


def _empty_paragraph_noop(node_id: str, node: dict[str, Any], position: str, details: dict[str, Any]) -> dict[str, Any]:
    return {
        "noop": True,
        "operation": f"insert-empty-paragraph-{position}",
        "reason": "idempotent-empty-paragraph-already-present",
        "partPath": node["partPath"],
        "entryName": node["partPath"].lstrip("/"),
        "nodeId": node_id,
        "locator": node["locator"],
        **details,
    }


def _empty_paragraph_insert_noop(
    bundle_dir: Path,
    node: dict[str, Any],
    position: str,
    requested_properties: dict[str, str],
) -> dict[str, Any] | None:
    part_path = node.get("partPath")
    locator = node.get("locator") or {}
    paragraph_index = int(locator.get("paragraphIndex") or 0)
    if not part_path or paragraph_index <= 0:
        return None
    local_part = bundle_dir / "parts" / Path(part_path.lstrip("/"))
    if not local_part.exists():
        return None
    try:
        root = ET.fromstring(local_part.read_bytes())
    except ET.ParseError:
        return None
    paragraphs = root.findall(".//w:p", NS)
    target_index = paragraph_index - 1
    neighbor_index = target_index - 1 if position == "before" else target_index + 1
    if neighbor_index < 0 or neighbor_index >= len(paragraphs):
        return None
    neighbor = paragraphs[neighbor_index]
    if not _is_empty_paragraph_xml(neighbor):
        return None
    existing_properties = _paragraph_spacing_properties_xml(neighbor)
    mismatched = {
        key: {"expected": value, "actual": existing_properties.get(key)}
        for key, value in requested_properties.items()
        if existing_properties.get(key) != value
    }
    if mismatched:
        return None
    return {
        "neighborParagraphIndex": neighbor_index + 1,
        "existingProperties": existing_properties,
        "requestedProperties": dict(requested_properties),
    }


def _run_property_patch(edit_index: int, node_id: str, node: dict[str, Any], changed: dict[str, str | None]) -> dict[str, Any]:
    return {
        "id": f"patch-{edit_index:06d}",
        "operation": "patch-run",
        "partPath": node["partPath"],
        "entryName": node["partPath"].lstrip("/"),
        "nodeId": node_id,
        "locator": node["locator"],
        "expectedOldHash": node["hash"],
        "oldProperties": node.get("properties", {}),
        "newProperties": changed,
    }


def _validation_report(
    bundle_dir: Path,
    ok: bool,
    errors: list[dict[str, Any]],
    warnings: list[dict[str, Any]],
) -> dict[str, Any]:
    return {
        "ok": ok,
        "bundle": str(bundle_dir),
        "errors": errors,
        "warnings": warnings,
    }


def _append_audit(bundle_dir: Path, payload: dict[str, Any]) -> None:
    audit_path = bundle_dir / "audit.log.jsonl"
    payload = {"time": datetime.now(timezone.utc).isoformat(), **payload}
    audit_path.parent.mkdir(parents=True, exist_ok=True)
    with audit_path.open("a", encoding="utf-8", newline="\n") as f:
        import json

        f.write(json.dumps(payload, ensure_ascii=False, sort_keys=True) + "\n")
