from __future__ import annotations

import zipfile
from pathlib import Path, PurePosixPath
from typing import Any

from .errors import HDocxError
from .utils import sha256_bytes


REQUIRED_DOCX_ENTRIES = {"[Content_Types].xml", "_rels/.rels"}


def _safe_entry_path(name: str) -> PurePosixPath:
    if not name or name.startswith("/") or "\\" in name:
        raise HDocxError(
            "PACKAGE_UNSAFE_ENTRY_NAME",
            "DOCX contains an unsafe ZIP entry name.",
            {"entry": name},
        )
    path = PurePosixPath(name)
    if any(part in ("", ".", "..") for part in path.parts):
        raise HDocxError(
            "PACKAGE_UNSAFE_ENTRY_NAME",
            "DOCX contains an unsafe ZIP entry name.",
            {"entry": name},
        )
    return path


def read_docx_entries(docx_path: Path) -> list[dict[str, Any]]:
    try:
        with zipfile.ZipFile(docx_path, "r") as zf:
            names = zf.namelist()
            missing = sorted(REQUIRED_DOCX_ENTRIES.difference(names))
            if missing:
                raise HDocxError(
                    "PACKAGE_NOT_DOCX",
                    "Input is a ZIP file but does not look like a DOCX package.",
                    {"missing": missing},
                )

            entries: list[dict[str, Any]] = []
            for order, info in enumerate(zf.infolist()):
                _safe_entry_path(info.filename)
                data = zf.read(info.filename)
                entries.append(
                    {
                        "path": info.filename,
                        "zipOrder": order,
                        "compressType": info.compress_type,
                        "compressedSize": info.compress_size,
                        "uncompressedSize": info.file_size,
                        "crc": f"{info.CRC:08x}",
                        "dateTime": list(info.date_time),
                        "sha256": sha256_bytes(data),
                    }
                )
            return entries
    except zipfile.BadZipFile as exc:
        raise HDocxError(
            "PACKAGE_BAD_ZIP",
            "Input is not a readable unencrypted DOCX ZIP package.",
            {"path": str(docx_path)},
        ) from exc


def extract_entries(docx_path: Path, target_dir: Path) -> None:
    target_dir.mkdir(parents=True, exist_ok=True)
    with zipfile.ZipFile(docx_path, "r") as zf:
        for info in zf.infolist():
            safe_path = _safe_entry_path(info.filename)
            if info.is_dir():
                continue
            out_path = target_dir.joinpath(*safe_path.parts)
            out_path.parent.mkdir(parents=True, exist_ok=True)
            out_path.write_bytes(zf.read(info.filename))


def compare_docx_entries(left: Path, right: Path) -> dict[str, Any]:
    left_list = read_docx_entries(left)
    right_list = read_docx_entries(right)
    left_entries = {entry["path"]: entry for entry in left_list}
    right_entries = {entry["path"]: entry for entry in right_list}
    left_names = set(left_entries)
    right_names = set(right_entries)

    changed: list[str] = []
    unchanged: list[str] = []
    metadata_changed: list[str] = []
    for name in sorted(left_names & right_names):
        if left_entries[name]["sha256"] != right_entries[name]["sha256"]:
            changed.append(name)
        else:
            unchanged.append(name)
            if _metadata_changed(left_entries[name], right_entries[name]):
                metadata_changed.append(name)

    left_only = sorted(left_names - right_names)
    right_only = sorted(right_names - left_names)
    return {
        "entryCounts": {
            "left": len(left_list),
            "right": len(right_list),
            "common": len(left_names & right_names),
            "changed": len(changed),
            "metadataChanged": len(metadata_changed),
            "leftOnly": len(left_only),
            "rightOnly": len(right_only),
            "unchanged": len(unchanged),
        },
        "leftOnly": left_only,
        "rightOnly": right_only,
        "changed": changed,
        "metadataChanged": metadata_changed,
        "changedEntries": [
            _diff_entry(name, "changed", left_entries.get(name), right_entries.get(name)) for name in changed
        ],
        "leftOnlyEntries": [_diff_entry(name, "left-only", left_entries.get(name), None) for name in left_only],
        "rightOnlyEntries": [_diff_entry(name, "right-only", None, right_entries.get(name)) for name in right_only],
        "metadataChangedEntries": [
            _diff_entry(name, "metadata-changed", left_entries.get(name), right_entries.get(name))
            for name in metadata_changed
        ],
        "identical": not changed and not left_only and not right_only,
        "zipMetadataIdentical": not metadata_changed,
    }


def _metadata_changed(left: dict[str, Any], right: dict[str, Any]) -> bool:
    return any(
        left.get(field) != right.get(field)
        for field in ("zipOrder", "compressType", "compressedSize", "crc", "dateTime")
    )


def _diff_entry(
    path: str,
    status: str,
    left: dict[str, Any] | None,
    right: dict[str, Any] | None,
) -> dict[str, Any]:
    result: dict[str, Any] = {
        "path": path,
        "status": status,
        "kind": _entry_kind(path),
    }
    if left is not None:
        result["left"] = _entry_snapshot(left)
    if right is not None:
        result["right"] = _entry_snapshot(right)
    if left is not None and right is not None:
        result["changedFields"] = [
            field
            for field in (
                "sha256",
                "uncompressedSize",
                "compressedSize",
                "crc",
                "compressType",
                "zipOrder",
                "dateTime",
            )
            if left.get(field) != right.get(field)
        ]
    return result


def _entry_snapshot(entry: dict[str, Any]) -> dict[str, Any]:
    return {
        "sha256": entry["sha256"],
        "uncompressedSize": entry["uncompressedSize"],
        "compressedSize": entry["compressedSize"],
        "crc": entry["crc"],
        "compressType": entry["compressType"],
        "zipOrder": entry["zipOrder"],
        "dateTime": entry["dateTime"],
    }


def _entry_kind(path: str) -> str:
    if path == "[Content_Types].xml":
        return "content-types"
    if path.endswith(".rels"):
        return "relationships"
    if path == "word/document.xml":
        return "main-document"
    if path.startswith("word/header") and path.endswith(".xml"):
        return "header"
    if path.startswith("word/footer") and path.endswith(".xml"):
        return "footer"
    if path == "word/footnotes.xml":
        return "footnotes"
    if path == "word/endnotes.xml":
        return "endnotes"
    if path == "word/comments.xml":
        return "comments"
    if path == "word/styles.xml":
        return "styles"
    if path == "word/numbering.xml":
        return "numbering"
    if path.startswith("word/media/"):
        return "media"
    if path.endswith(".xml"):
        return "xml"
    return "binary"


def repack_docx_with_modified_entries(
    original_docx: Path,
    output_docx: Path,
    modified_entries: dict[str, bytes],
) -> None:
    output_docx.parent.mkdir(parents=True, exist_ok=True)
    written: set[str] = set()
    with zipfile.ZipFile(original_docx, "r") as source, zipfile.ZipFile(output_docx, "w") as target:
        for info in source.infolist():
            _safe_entry_path(info.filename)
            data = modified_entries.get(info.filename)
            if data is None:
                data = source.read(info.filename)
            written.add(info.filename)
            out_info = zipfile.ZipInfo(info.filename, info.date_time)
            out_info.compress_type = info.compress_type
            out_info.comment = info.comment
            out_info.extra = info.extra
            out_info.internal_attr = info.internal_attr
            out_info.external_attr = info.external_attr
            out_info.create_system = info.create_system
            target.writestr(out_info, data)
        for entry_name in sorted(set(modified_entries) - written):
            _safe_entry_path(entry_name)
            out_info = zipfile.ZipInfo(entry_name)
            out_info.compress_type = zipfile.ZIP_DEFLATED
            target.writestr(out_info, modified_entries[entry_name])
