from __future__ import annotations

import json
import os
import sys
from dataclasses import dataclass
from pathlib import Path
from typing import Any, Callable

from .errors import HDocxError
from .fixtures import generate_pressure_fixtures
from .hdocx import (
    apply_hdocx,
    audit_docx,
    batch_check_docx,
    check_docx,
    diff_docx,
    doctor_report,
    export_docx,
    inspect_hdocx,
    plan_hdocx,
    roundtrip_docx,
    validate_hdocx,
)
from .rendering import render_check_docx
from .utils import write_json


PROTOCOL_VERSION = "2024-11-05"
SERVER_INFO = {"name": "html-docx-mcp", "version": "0.1.0"}
ROOT_DESCRIPTION = (
    "Workspace root. Defaults to HDOCX_MCP_ROOT, then CLAUDE_PROJECT_DIR, "
    "then the MCP server current directory."
)


@dataclass(frozen=True)
class Tool:
    name: str
    description: str
    input_schema: dict[str, Any]
    handler: Callable[[dict[str, Any]], dict[str, Any]]


def main() -> int:
    server = HDocxMcpServer()
    return server.run()


class HDocxMcpServer:
    def __init__(self) -> None:
        self.tools = _build_tools()

    def run(self) -> int:
        for raw_line in sys.stdin:
            line = raw_line.strip()
            if not line:
                continue
            try:
                request = json.loads(line)
                response = self.handle_message(request)
            except Exception as exc:
                response = _error_response(None, -32700, "Parse error", {"message": str(exc)})
            if response is not None:
                _write_message(response)
        return 0

    def handle_message(self, message: dict[str, Any]) -> dict[str, Any] | None:
        request_id = message.get("id")
        method = message.get("method")
        params = message.get("params") or {}

        if method == "initialize":
            protocol_version = params.get("protocolVersion") or PROTOCOL_VERSION
            return _result_response(
                request_id,
                {
                    "protocolVersion": protocol_version,
                    "capabilities": {
                        "tools": {"listChanged": False},
                        "resources": {"subscribe": False, "listChanged": False},
                        "prompts": {"listChanged": False},
                    },
                    "serverInfo": SERVER_INFO,
                },
            )

        if method == "notifications/initialized":
            return None

        if method == "ping":
            return _result_response(request_id, {})

        if method == "tools/list":
            return _result_response(
                request_id,
                {
                    "tools": [
                        {
                            "name": tool.name,
                            "description": tool.description,
                            "inputSchema": tool.input_schema,
                        }
                        for tool in self.tools.values()
                    ]
                },
            )

        if method == "tools/call":
            return self._call_tool(request_id, params)

        if method == "resources/list":
            return _result_response(request_id, {"resources": []})

        if method == "prompts/list":
            return _result_response(request_id, {"prompts": []})

        if method == "logging/setLevel":
            return _result_response(request_id, {})

        if request_id is None:
            return None
        return _error_response(request_id, -32601, "Method not found", {"method": method})

    def _call_tool(self, request_id: Any, params: dict[str, Any]) -> dict[str, Any]:
        tool_name = params.get("name")
        arguments = params.get("arguments") or {}
        tool = self.tools.get(tool_name)
        if tool is None:
            return _error_response(request_id, -32602, "Unknown tool", {"tool": tool_name})
        if not isinstance(arguments, dict):
            return _error_response(request_id, -32602, "Tool arguments must be an object", {"tool": tool_name})

        try:
            payload = tool.handler(arguments)
            return _tool_response(request_id, payload, is_error=not bool(payload.get("ok", True)))
        except HDocxError as exc:
            return _tool_response(
                request_id,
                {"ok": False, "error": exc.to_dict()},
                is_error=True,
            )
        except Exception as exc:
            return _tool_response(
                request_id,
                {
                    "ok": False,
                    "error": {
                        "code": "UNEXPECTED_ERROR",
                        "severity": "error",
                        "message": str(exc),
                    },
                },
                is_error=True,
            )


def _build_tools() -> dict[str, Tool]:
    tools = [
        Tool(
            "hdocx_doctor",
            "Report html_docx runtime capabilities and optional renderer availability.",
            _schema(properties={}),
            lambda args: doctor_report(),
        ),
        Tool(
            "hdocx_audit",
            "Audit high-risk DOCX structures and preservation policies.",
            _schema(
                required=["input"],
                properties={
                    "root": _string(ROOT_DESCRIPTION),
                    "input": _string("Input DOCX path, relative to root or absolute under root."),
                    "report": _string("Optional JSON report path under root."),
                },
            ),
            lambda args: _with_report(args, audit_docx(_path_arg(args, "input", must_exist=True))),
        ),
        Tool(
            "hdocx_export",
            "Export a DOCX file to an H-DOCX bundle directory.",
            _schema(
                required=["input", "out"],
                properties={
                    "root": _string(ROOT_DESCRIPTION),
                    "input": _string("Input DOCX path under root."),
                    "out": _string("Output .hdocx directory under root."),
                    "force": _boolean("Replace an existing output directory."),
                    "report": _string("Optional JSON report path under root."),
                },
            ),
            lambda args: _with_report(
                args,
                export_docx(_path_arg(args, "input", must_exist=True), _path_arg(args, "out"), force=bool(args.get("force", False))),
            ),
        ),
        Tool(
            "hdocx_validate",
            "Validate an H-DOCX bundle.",
            _schema(
                required=["bundle"],
                properties={
                    "root": _string(ROOT_DESCRIPTION),
                    "bundle": _string("H-DOCX bundle directory under root."),
                    "report": _string("Optional JSON report path under root."),
                },
            ),
            lambda args: _with_report(args, validate_hdocx(_path_arg(args, "bundle", must_exist=True))),
        ),
        Tool(
            "hdocx_inspect",
            "Inspect a projected H-DOCX node, style, list, table, or image by id.",
            _schema(
                required=["bundle", "id"],
                properties={
                    "root": _string(ROOT_DESCRIPTION),
                    "bundle": _string("H-DOCX bundle directory under root."),
                    "id": _string("Target id such as p-000001, r-000001, Normal, or tbl-000001."),
                    "kind": {
                        "type": "string",
                        "enum": ["node", "style", "list", "table", "image"],
                        "description": "Manifest item kind. Defaults to node.",
                    },
                    "report": _string("Optional JSON report path under root."),
                },
            ),
            lambda args: _with_report(
                args,
                inspect_hdocx(_path_arg(args, "bundle", must_exist=True), str(args["id"]), kind=str(args.get("kind", "node"))),
            ),
        ),
        Tool(
            "hdocx_plan",
            "Plan H-DOCX edits without writing a DOCX.",
            _schema(
                required=["bundle"],
                properties={
                    "root": _string(ROOT_DESCRIPTION),
                    "bundle": _string("H-DOCX bundle directory under root."),
                    "report": _string("Optional JSON report path under root."),
                },
            ),
            lambda args: _with_report(args, plan_hdocx(_path_arg(args, "bundle", must_exist=True))),
        ),
        Tool(
            "hdocx_apply",
            "Apply an H-DOCX bundle back to a DOCX file.",
            _schema(
                required=["bundle", "out"],
                properties={
                    "root": _string(ROOT_DESCRIPTION),
                    "bundle": _string("H-DOCX bundle directory under root."),
                    "out": _string("Output DOCX path under root."),
                    "report": _string("Optional JSON report path under root."),
                },
            ),
            lambda args: _with_report(args, apply_hdocx(_path_arg(args, "bundle", must_exist=True), _path_arg(args, "out"))),
        ),
        Tool(
            "hdocx_diff",
            "Compare two DOCX packages with package, semantic, and fragment-level reports.",
            _schema(
                required=["left", "right"],
                properties={
                    "root": _string(ROOT_DESCRIPTION),
                    "left": _string("Left DOCX path under root."),
                    "right": _string("Right DOCX path under root."),
                    "report": _string("Optional JSON report path under root."),
                },
            ),
            lambda args: _with_report(args, diff_docx(_path_arg(args, "left", must_exist=True), _path_arg(args, "right", must_exist=True))),
        ),
        Tool(
            "hdocx_roundtrip",
            "Export and apply a DOCX without edits.",
            _schema(
                required=["input", "work", "out"],
                properties={
                    "root": _string(ROOT_DESCRIPTION),
                    "input": _string("Input DOCX path under root."),
                    "work": _string("Output H-DOCX work directory under root."),
                    "out": _string("Output DOCX path under root."),
                    "force": _boolean("Replace an existing work directory."),
                    "report": _string("Optional JSON report path under root."),
                },
            ),
            lambda args: _with_report(
                args,
                roundtrip_docx(
                    _path_arg(args, "input", must_exist=True),
                    _path_arg(args, "work"),
                    _path_arg(args, "out"),
                    force=bool(args.get("force", False)),
                ),
            ),
        ),
        Tool(
            "hdocx_check",
            "Run export/apply/diff acceptance for one DOCX.",
            _schema(
                required=["input", "work", "out"],
                properties={
                    "root": _string(ROOT_DESCRIPTION),
                    "input": _string("Input DOCX path under root."),
                    "work": _string("Output H-DOCX work directory under root."),
                    "out": _string("Output DOCX path under root."),
                    "force": _boolean("Replace an existing work directory."),
                    "report": _string("Optional JSON report path under root."),
                },
            ),
            lambda args: _with_report(
                args,
                check_docx(
                    _path_arg(args, "input", must_exist=True),
                    _path_arg(args, "work"),
                    _path_arg(args, "out"),
                    force=bool(args.get("force", False)),
                ),
            ),
        ),
        Tool(
            "hdocx_batch_check",
            "Run check over a DOCX file or directory.",
            _schema(
                required=["input", "work", "out"],
                properties={
                    "root": _string(ROOT_DESCRIPTION),
                    "input": _string("Input DOCX file or directory under root."),
                    "work": _string("Output work directory under root."),
                    "out": _string("Output DOCX directory under root."),
                    "force": _boolean("Replace existing generated work/output entries."),
                    "report": _string("Optional JSON report path under root."),
                },
            ),
            lambda args: _with_report(
                args,
                batch_check_docx(
                    _path_arg(args, "input", must_exist=True),
                    _path_arg(args, "work"),
                    _path_arg(args, "out"),
                    force=bool(args.get("force", False)),
                ),
            ),
        ),
        Tool(
            "hdocx_generate_fixtures",
            "Generate local synthetic DOCX pressure fixtures.",
            _schema(
                required=["out"],
                properties={
                    "root": _string(ROOT_DESCRIPTION),
                    "out": _string("Output fixture directory under root."),
                    "force": _boolean("Replace an existing output directory."),
                    "report": _string("Optional JSON report path under root."),
                },
            ),
            lambda args: _with_report(args, generate_pressure_fixtures(_path_arg(args, "out"), force=bool(args.get("force", False)))),
        ),
        Tool(
            "hdocx_render_check",
            "Optionally render a DOCX to PDF with LibreOffice/soffice for visual QA.",
            _schema(
                required=["input", "out"],
                properties={
                    "root": _string(ROOT_DESCRIPTION),
                    "input": _string("Input DOCX path under root."),
                    "out": _string("Output render directory under root."),
                    "force": _boolean("Replace an existing render directory."),
                    "allowMissing": _boolean("Return a structured renderer-missing report when LibreOffice/soffice is unavailable."),
                    "timeoutSeconds": {"type": "integer", "minimum": 1, "description": "Render timeout in seconds. Defaults to 120."},
                    "report": _string("Optional JSON report path under root."),
                },
            ),
            lambda args: _with_report(
                args,
                render_check_docx(
                    _path_arg(args, "input", must_exist=True),
                    _path_arg(args, "out"),
                    force=bool(args.get("force", False)),
                    allow_missing=bool(args.get("allowMissing", False)),
                    timeout_seconds=int(args.get("timeoutSeconds", 120)),
                ),
            ),
        ),
    ]
    return {tool.name: tool for tool in tools}


def _schema(*, properties: dict[str, Any], required: list[str] | None = None) -> dict[str, Any]:
    return {
        "type": "object",
        "properties": properties,
        "required": required or [],
        "additionalProperties": False,
    }


def _string(description: str) -> dict[str, Any]:
    return {"type": "string", "description": description}


def _boolean(description: str) -> dict[str, Any]:
    return {"type": "boolean", "description": description}


def _path_arg(args: dict[str, Any], name: str, *, must_exist: bool = False) -> Path:
    value = args.get(name)
    if not isinstance(value, str) or not value:
        raise HDocxError("MCP_INVALID_ARGUMENT", f"Missing or invalid path argument: {name}.", {"argument": name})
    root = _root_arg(args)
    candidate = Path(value)
    if not candidate.is_absolute():
        candidate = root / candidate
    resolved = candidate.resolve()
    _ensure_under_root(resolved, root, argument=name)
    if must_exist and not resolved.exists():
        raise HDocxError("MCP_PATH_NOT_FOUND", f"Path does not exist: {name}.", {"argument": name, "path": str(resolved)})
    return resolved


def _root_arg(args: dict[str, Any]) -> Path:
    raw_root = args.get("root")
    if raw_root is None:
        raw_root = os.environ.get("HDOCX_MCP_ROOT") or os.environ.get("CLAUDE_PROJECT_DIR")
    if raw_root is None:
        return Path.cwd().resolve()
    if not isinstance(raw_root, str) or not raw_root:
        raise HDocxError("MCP_INVALID_ARGUMENT", "root must be a non-empty string.", {"argument": "root"})
    return Path(raw_root).resolve()


def _ensure_under_root(path: Path, root: Path, *, argument: str) -> None:
    try:
        path.relative_to(root)
    except ValueError as exc:
        raise HDocxError(
            "MCP_PATH_OUTSIDE_ROOT",
            f"Path for {argument} is outside the declared workspace root.",
            {"argument": argument, "path": str(path), "root": str(root)},
        ) from exc


def _with_report(args: dict[str, Any], payload: dict[str, Any]) -> dict[str, Any]:
    report = args.get("report")
    if report:
        report_path = _path_arg(args, "report")
        write_json(report_path, payload)
        payload = dict(payload)
        payload["mcpReportPath"] = str(report_path)
    return payload


def _tool_response(request_id: Any, payload: dict[str, Any], *, is_error: bool) -> dict[str, Any]:
    text = json.dumps(payload, ensure_ascii=False, indent=2, sort_keys=True)
    result: dict[str, Any] = {
        "content": [{"type": "text", "text": text}],
        "structuredContent": payload,
        "isError": is_error,
    }
    return _result_response(request_id, result)


def _result_response(request_id: Any, result: dict[str, Any]) -> dict[str, Any]:
    return {"jsonrpc": "2.0", "id": request_id, "result": result}


def _error_response(request_id: Any, code: int, message: str, data: dict[str, Any] | None = None) -> dict[str, Any]:
    error: dict[str, Any] = {"code": code, "message": message}
    if data is not None:
        error["data"] = data
    return {"jsonrpc": "2.0", "id": request_id, "error": error}


def _write_message(message: dict[str, Any]) -> None:
    sys.stdout.write(json.dumps(message, ensure_ascii=False, separators=(",", ":")))
    sys.stdout.write("\n")
    sys.stdout.flush()


if __name__ == "__main__":
    raise SystemExit(main())
