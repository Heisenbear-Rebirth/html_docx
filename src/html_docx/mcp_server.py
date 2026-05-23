from __future__ import annotations

import json
import os
import sys
import threading
import hashlib
from dataclasses import dataclass
from pathlib import Path
from typing import Any, Callable

from .errors import HDocxError
from .fixtures import generate_pressure_fixtures
from .hdocx import (
    apply_hdocx,
    assert_hdocx,
    audit_docx,
    batch_check_docx,
    check_docx,
    create_docx,
    diff_docx,
    doctor_report,
    export_docx,
    find_hdocx,
    inspect_hdocx,
    plan_hdocx,
    query_hdocx,
    roundtrip_docx,
    SUPPORTED_EDIT_MODES,
    SUPPORTED_HCSS_AT_RULES,
    SUPPORTED_HCSS_PROPERTIES,
    validate_hdocx,
)
from .utils import write_json


PROTOCOL_VERSION = "2024-11-05"
SERVER_INFO = {"name": "html-docx-mcp", "version": "0.1.1"}
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


@dataclass(frozen=True)
class GuideResource:
    uri: str
    name: str
    description: str
    text: str
    mime_type: str = "text/markdown"


@dataclass(frozen=True)
class PromptTemplate:
    name: str
    description: str
    arguments: list[dict[str, Any]]
    handler: Callable[[dict[str, Any]], str]


def main() -> int:
    _configure_stdio_utf8()
    server = HDocxMcpServer()
    return server.run()


class HDocxMcpServer:
    def __init__(self) -> None:
        self.tools = _build_tools()
        self.resources = _build_resources()
        self.prompts = _build_prompts()
        self._tool_lock = threading.Lock()

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
            return _result_response(
                request_id,
                {
                    "resources": [
                        {
                            "uri": resource.uri,
                            "name": resource.name,
                            "description": resource.description,
                            "mimeType": resource.mime_type,
                        }
                        for resource in self.resources.values()
                    ]
                },
            )

        if method == "resources/read":
            return self._read_resource(request_id, params)

        if method == "prompts/list":
            return _result_response(
                request_id,
                {
                    "prompts": [
                        {
                            "name": prompt.name,
                            "description": prompt.description,
                            "arguments": prompt.arguments,
                        }
                        for prompt in self.prompts.values()
                    ]
                },
            )

        if method == "prompts/get":
            return self._get_prompt(request_id, params)

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

        if not self._tool_lock.acquire(blocking=False):
            return _tool_response(
                request_id,
                {
                    "ok": False,
                    "error": {
                        "code": "MCP_SERVER_BUSY",
                        "severity": "error",
                        "message": "html-docx-mcp is already running another tool call; retry this call sequentially.",
                    },
                },
                is_error=True,
            )
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
        finally:
            self._tool_lock.release()

    def _read_resource(self, request_id: Any, params: dict[str, Any]) -> dict[str, Any]:
        uri = params.get("uri")
        resource = self.resources.get(uri)
        if resource is None:
            return _error_response(request_id, -32602, "Unknown resource", {"uri": uri})
        return _result_response(
            request_id,
            {
                "contents": [
                    {
                        "uri": resource.uri,
                        "mimeType": resource.mime_type,
                        "text": resource.text,
                    }
                ]
            },
        )

    def _get_prompt(self, request_id: Any, params: dict[str, Any]) -> dict[str, Any]:
        name = params.get("name")
        arguments = params.get("arguments") or {}
        prompt = self.prompts.get(name)
        if prompt is None:
            return _error_response(request_id, -32602, "Unknown prompt", {"prompt": name})
        if not isinstance(arguments, dict):
            return _error_response(request_id, -32602, "Prompt arguments must be an object", {"prompt": name})
        text = prompt.handler(arguments)
        return _result_response(
            request_id,
            {
                "description": prompt.description,
                "messages": [
                    {
                        "role": "user",
                        "content": {"type": "text", "text": text},
                    }
                ],
            },
        )


def _configure_stdio_utf8() -> None:
    for stream_name in ("stdin", "stdout", "stderr"):
        stream = getattr(sys, stream_name)
        reconfigure = getattr(stream, "reconfigure", None)
        if reconfigure is None:
            continue
        kwargs = {"encoding": "utf-8"}
        if stream_name == "stdin":
            kwargs["errors"] = "surrogateescape"
        else:
            kwargs["errors"] = "strict"
        try:
            reconfigure(**kwargs)
        except (OSError, ValueError):
            pass


def _mcp_doctor_tool(args: dict[str, Any]) -> dict[str, Any]:
    report = doctor_report()
    guidance_info = _guidance_runtime_info()
    report["mcp"] = {
        "serverInfo": SERVER_INFO,
        "loadedModulePath": str(Path(__file__).resolve()),
        "loaded_module_path": str(Path(__file__).resolve()),
        "moduleMtimeUtc": _module_mtime(Path(__file__).resolve()),
        "stdioEncoding": {
            "stdin": getattr(sys.stdin, "encoding", None),
            "stdout": getattr(sys.stdout, "encoding", None),
            "stderr": getattr(sys.stderr, "encoding", None),
        },
        "concurrencyPolicy": {
            "toolCallsSerializedByServerInstance": True,
            "busyErrorCode": "MCP_SERVER_BUSY",
            "note": "Clients may still appear to run quick calls in parallel if they deliver them sequentially to this stdio server.",
        },
        "guidance": guidance_info,
        "guidance_source_path": guidance_info["sourcePath"],
        "guidance_version": guidance_info["sha256ByTopic"],
    }
    report["runtime"]["guidance_source_path"] = guidance_info["sourcePath"]
    report["runtime"]["guidance_version"] = guidance_info["sha256ByTopic"]
    return report


def _module_mtime(path: Path) -> str | None:
    try:
        from datetime import datetime, timezone

        return datetime.fromtimestamp(path.stat().st_mtime, tz=timezone.utc).isoformat()
    except OSError:
        return None


def _guidance_runtime_info() -> dict[str, Any]:
    guidance_map = {
        "workflow": WORKFLOW_GUIDANCE,
        "format": WRITING_FORMAT_GUIDANCE,
        "query": QUERY_GUIDANCE,
        "hcss": f"{HCSS_GUIDANCE}\n\n{_hcss_capability_registry_markdown()}",
        "acceptance": ACCEPTANCE_GUIDANCE,
        "edge-cases": EDGE_CASE_GUIDANCE,
    }
    return {
        "sourcePath": str(Path(__file__).resolve()),
        "topics": sorted(guidance_map),
        "sha256ByTopic": {
            topic: hashlib.sha256(text.encode("utf-8")).hexdigest()
            for topic, text in guidance_map.items()
        },
    }


def _hcss_capability_registry_markdown() -> str:
    lines = [
        "## Runtime H-CSS Capability Registry",
        "",
        "This section is generated from the runtime capability registry used by `hdocx_plan`.",
        "",
        "Supported edit modes:",
        "",
    ]
    for mode in SUPPORTED_EDIT_MODES:
        lines.append(f"- `{mode}`")
    lines.extend(["", "Supported properties by mode:", ""])
    for mode in SUPPORTED_EDIT_MODES:
        properties = SUPPORTED_HCSS_PROPERTIES.get(mode, [])
        lines.append(f"### `{mode}`")
        if not properties:
            lines.append("")
            lines.append("- No direct H-CSS declarations registered.")
        else:
            for prop in properties:
                lines.append(f"- `{prop}`")
        lines.append("")
    lines.extend(["Supported H-CSS at-rules:", ""])
    for at_rule in SUPPORTED_HCSS_AT_RULES:
        lines.append(f"- `{at_rule}`")
    return "\n".join(lines).rstrip()


def _build_tools() -> dict[str, Tool]:
    tools = [
        Tool(
            "hdocx_doctor",
            "Report html_docx runtime capabilities.",
            _schema(properties={}),
            _mcp_doctor_tool,
        ),
        Tool(
            "hdocx_guidance",
            "Return H-DOCX authoring rules, H-CSS format guidance, and required acceptance checks.",
            _schema(
                properties={
                    "topic": {
                        "type": "string",
                        "enum": ["all", "workflow", "format", "query", "hcss", "acceptance", "edge-cases"],
                        "description": "Guidance topic. Defaults to all.",
                    }
                }
            ),
            _guidance_tool,
        ),
        Tool(
            "hdocx_create",
            "Create a new canonical DOCX and optionally export it to an H-DOCX bundle.",
            _schema(
                required=["out"],
                properties={
                    "root": _string(ROOT_DESCRIPTION),
                    "out": _string("Output DOCX path under root."),
                    "title": _string("Optional document title."),
                    "paragraphs": {
                        "type": "array",
                        "items": {"type": "string"},
                        "description": "Optional body paragraphs. Defaults to one empty editable paragraph.",
                    },
                    "template": {
                        "type": "string",
                        "enum": ["blank"],
                        "description": "Creation template. Defaults to blank.",
                    },
                    "force": _boolean("Replace an existing output DOCX and exported bundle."),
                    "exportTo": _string("Optional H-DOCX bundle directory to export after creation."),
                    "report": _string("Optional JSON report path under root."),
                },
            ),
            lambda args: _with_report(
                args,
                create_docx(
                    _path_arg(args, "out"),
                    title=_optional_string_arg(args, "title"),
                    paragraphs=_string_list_arg(args, "paragraphs"),
                    template=_optional_string_arg(args, "template") or "blank",
                    force=bool(args.get("force", False)),
                    export_dir=_path_arg(args, "exportTo") if args.get("exportTo") else None,
                ),
            ),
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
            "hdocx_query",
            "Query projected H-DOCX nodes by text, font, size, paragraph properties, images, or likely level-1 headings.",
            _query_tool_schema(),
            lambda args: _with_report(args, query_hdocx(_path_arg(args, "bundle", must_exist=True), **_query_tool_kwargs(args))),
        ),
        Tool(
            "hdocx_find",
            "Alias for hdocx_query; returns structured JSON so agents do not need to read document.html.",
            _query_tool_schema(),
            lambda args: _with_report(args, find_hdocx(_path_arg(args, "bundle", must_exist=True), **_query_tool_kwargs(args))),
        ),
        Tool(
            "hdocx_assert",
            "Run assertion-style acceptance checks over an H-DOCX bundle.",
            _schema(
                required=["bundle"],
                properties={
                    "root": _string(ROOT_DESCRIPTION),
                    "bundle": _string("H-DOCX bundle directory under root."),
                    "assertions": {
                        "type": "array",
                        "items": {
                            "type": ["object", "string"],
                            "description": "Assertion string or object with a type field and optional parameters such as afterApply, includeRegex, excludeRegex, or paragraphIds.",
                        },
                        "description": "Assertions to run. Defaults to text-payload-unchanged. Set afterApply/plannedOutput on an assertion object to check the planned output state.",
                    },
                    "report": _string("Optional JSON report path under root."),
                },
            ),
            lambda args: _with_report(args, assert_hdocx(_path_arg(args, "bundle", must_exist=True), _assertions_arg(args))),
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
    ]
    return {tool.name: tool for tool in tools}


def _build_resources() -> dict[str, GuideResource]:
    resources = [
        GuideResource(
            "hdocx://guide/workflow",
            "H-DOCX safe editing workflow",
            "Required DOCX editing workflow for agents.",
            WORKFLOW_GUIDANCE,
        ),
        GuideResource(
            "hdocx://guide/writing-format",
            "H-DOCX writing format",
            "How to edit document.html and what must remain read-only.",
            WRITING_FORMAT_GUIDANCE,
        ),
        GuideResource(
            "hdocx://guide/hcss",
            "H-CSS selectors and formatting rules",
            "Supported selector and reusable formatting patterns.",
            f"{HCSS_GUIDANCE}\n\n{_hcss_capability_registry_markdown()}",
        ),
        GuideResource(
            "hdocx://guide/query",
            "H-DOCX query and assertion tools",
            "Structured target discovery and assertion-style acceptance.",
            QUERY_GUIDANCE,
        ),
        GuideResource(
            "hdocx://guide/acceptance",
            "H-DOCX acceptance checks",
            "Checks required before claiming a DOCX edit is safe.",
            ACCEPTANCE_GUIDANCE,
        ),
        GuideResource(
            "hdocx://guide/edge-cases",
            "H-DOCX edge cases",
            "Common edge cases that must be protected or handled conservatively.",
            EDGE_CASE_GUIDANCE,
        ),
    ]
    return {resource.uri: resource for resource in resources}


def _build_prompts() -> dict[str, PromptTemplate]:
    prompts = [
        PromptTemplate(
            "hdocx_create_docx",
            "Create a new DOCX from the canonical H-DOCX blank template.",
            [
                _prompt_arg("root", "Workspace root for the new DOCX and optional bundle.", required=True),
                _prompt_arg("out", "Output DOCX path relative to root.", required=True),
                _prompt_arg("goal", "Requested document content or purpose.", required=True),
                _prompt_arg("title", "Optional document title.", required=False),
                _prompt_arg("work", "Optional H-DOCX bundle directory relative to root.", required=False),
            ],
            _create_docx_prompt,
        ),
        PromptTemplate(
            "hdocx_safe_edit",
            "Safely edit a DOCX through an H-DOCX bundle.",
            [
                _prompt_arg("root", "Workspace root containing all inputs and outputs.", required=True),
                _prompt_arg("input_docx", "Input DOCX path relative to root.", required=True),
                _prompt_arg("goal", "User requested edit.", required=True),
                _prompt_arg("work", "H-DOCX work directory relative to root.", required=False),
                _prompt_arg("out", "Output DOCX path relative to root.", required=False),
            ],
            _safe_edit_prompt,
        ),
        PromptTemplate(
            "hdocx_format_change",
            "Apply a layout or formatting change with H-CSS.",
            [
                _prompt_arg("root", "Workspace root containing all inputs and outputs.", required=True),
                _prompt_arg("input_docx", "Input DOCX path relative to root.", required=True),
                _prompt_arg("goal", "Formatting request.", required=True),
                _prompt_arg("target", "Known id, style id, list level, or narrow selector.", required=False),
                _prompt_arg("work", "H-DOCX work directory relative to root.", required=False),
                _prompt_arg("out", "Output DOCX path relative to root.", required=False),
            ],
            _format_change_prompt,
        ),
        PromptTemplate(
            "hdocx_roundtrip_check",
            "Prove a DOCX family can round-trip without edits.",
            [
                _prompt_arg("root", "Workspace root containing all inputs and outputs.", required=True),
                _prompt_arg("input_docx", "Input DOCX path relative to root.", required=True),
                _prompt_arg("work", "H-DOCX work directory relative to root.", required=False),
                _prompt_arg("out", "Output DOCX path relative to root.", required=False),
            ],
            _roundtrip_prompt,
        ),
    ]
    return {prompt.name: prompt for prompt in prompts}


WORKFLOW_GUIDANCE = """# H-DOCX Safe Editing Workflow

H-DOCX is a strict DOCX <-> H-DOCX bundle <-> DOCX workflow. It is not a
best-effort HTML converter.

Required workflow for agents:

For a new document, call `hdocx_create` first. It creates a canonical DOCX from
the built-in blank template and can optionally export that new DOCX to an
H-DOCX bundle for immediate editing.

For an existing document:

1. Run `hdocx_audit` on the source DOCX.
2. Run `hdocx_export` to create a `.hdocx` bundle.
3. Run `hdocx_query` / `hdocx_find` before reading `document.html` directly.
   Use them to locate text, formatting, images and likely level-1 headings.
4. Run `hdocx_inspect` before targeting styles, lists, tables, images, or
   specific paragraphs/runs.
5. Edit only allowed bundle surfaces: editable run text in `document.html`,
   supported rules in `agent.edits.hcss`, and supported bundle-local assets.
6. Run `hdocx_plan` before writing an output DOCX.
7. Run `hdocx_apply` to produce the output DOCX.
8. Run `hdocx_diff` against the original before claiming success.
9. Run `hdocx_assert` for task-specific invariants such as unchanged text
   payload, structural blank paragraphs, and image host paragraph spacing.

For no-edit reversibility proof, use `hdocx_check` directly. A byte-identical
SHA256 match is stronger than render equality for unedited round-trips.

MCP tool calls should be serialized. If two tool handlers overlap inside the
same server instance, the server returns a structured `MCP_SERVER_BUSY` error
instead of allowing a second operation to interfere with the stdio transport.
Some clients queue quick calls even when the agent asks for them in parallel;
those calls may both succeed because they reached the server sequentially.
"""


WRITING_FORMAT_GUIDANCE = """# H-DOCX Writing Format

The editable projection is `document.html`. It is intentionally HTML-like, but
it is not normal web HTML.

Agent rules:

- Keep `manifest.json` and `original/original.docx` read-only.
- Keep `data-hdocx-id`, `data-hdocx-part`, style ids, numbering ids, and
  protected-kind attributes unchanged.
- Do not remove protected placeholders such as equations, content controls,
  fields, tracked-change islands, unsupported drawings, or unknown XML.
- Paragraphs map to paragraph-level projected nodes. Runs inside a paragraph
  may have different text properties such as font size, bold, italic, language,
  highlight, or script-specific fonts.
- Change literal text only inside editable run text. If a requested text edit
  crosses multiple runs, preserve the existing run boundaries unless there is a
  dedicated supported operation that proves the merge is safe.
- Do not invent ordinary HTML/CSS behavior. Use `agent.edits.hcss` for
  supported H-CSS formatting changes.

When unsure, inspect the relevant node and fail with a report instead of
guessing.
"""


QUERY_GUIDANCE = """# H-DOCX Query And Assert Tools

Use `hdocx_query` / `hdocx_find` before reading `document.html` directly. They
return structured JSON from the H-DOCX projection manifest and avoid terminal
encoding problems with large or multilingual HTML.

Common target discovery:

```json
{"bundle":"work.hdocx","text":"Keywords"}
{"bundle":"work.hdocx","align":"center","fontSize":"14pt","fontFamily":"SimHei","suspectedHeadingLevel1":true}
{"bundle":"work.hdocx","kind":"image"}
```

Query filters include `text`, `textRegex`, `styleId`, `fontFamily`,
`eastAsiaFont`, `fontSize`, `bold`, `italic`, `color`, `align`, `lineSpacing`,
`spaceBefore`, `spaceAfter`, `hasImage`, and `suspectedHeadingLevel1`.

Use `hdocx_assert` after edits for explicit acceptance checks:

```json
{
  "bundle": "work.hdocx",
  "assertions": [
    "text-payload-unchanged",
    "images-host-paragraph-not-exact-line-spacing",
    {"type":"paragraphs-have-empty-before","paragraphIds":["p-000006"]},
    {"type":"paragraphs-have-empty-before","paragraphIds":["p-000006"],"afterApply":true},
    {"type":"level1-headings-have-empty-paragraph-before","minScore":3,"includeRegex":"^\\\\d+","excludeRegex":"Appendix"}
  ]
}
```

Supported assertions:

- `text-payload-unchanged`
- `paragraphs-have-empty-before`
- `paragraphs-have-empty-after`
- `images-host-paragraph-not-exact-line-spacing`
- `level1-headings-have-empty-paragraph-before`

Assertion scope:

- Without `afterApply`, structure and formatting assertions inspect the current
  exported bundle. `text-payload-unchanged` checks the planned patch list.
- With `afterApply: true` or `plannedOutput: true`, the server applies the
  current plan to a bundle-local scratch DOCX, exports it again, maps original
  paragraph ids to the planned output when possible, and checks that output
  state.
- `level1-headings-have-empty-paragraph-before` is heuristic. It supports
  `includeRegex`, `excludeRegex`, `paragraphIds`, `minScore`,
  `allowFirstInPart`, and `useDefaultExcludes`. By default it excludes
  front-matter labels such as abstract, contents/table of contents, and
  keywords.
"""


HCSS_GUIDANCE = """# H-CSS Selector And Formatting Guidance

Use `agent.edits.hcss` for controlled formatting edits. Prefer inspected ids or
named sets over broad selectors.

Recommended targeting patterns:

```css
@hdocx-set target-paragraph {
  select: id(p-000001);
}

@hdocx-set selected-body {
  select: id(p-000007), id(p-000008), id(p-000009);
}

@hdocx-set body-style {
  select: style(BodyText);
}

@hdocx-set first-level-list {
  select: list(1, 0);
}

@hdocx-set header-paragraphs {
  select: part(/word/header1.xml, paragraph);
}
```

If a selector is allowed to match nothing, say so explicitly:

```css
@hdocx-set optional-notes {
  select: .maybe-note;
  allow-empty: true;
}
```

Selector support is intentionally small: ids, classes, exact attribute
selectors, class+attribute compounds such as
`.hdocx-r[data-hdocx-id="r-000001"]`, selector lists separated by commas, and
the H-DOCX functions above.

Do not add custom grouping classes or other projection metadata to
`document.html`; that file's read-only metadata is part of the reversible
projection. Put reusable groups in `agent.edits.hcss` with `@hdocx-set`, or use
comma selector lists directly:

```css
.role-body,
.role-reference {
  hdocx-font-size: 10.5pt;
}
```

Do not use broad selectors when the user asked for a narrow change. Do not
write normal CSS properties unless H-DOCX explicitly supports them.

## Supported Formatting Declarations

All declarations must use the `hdocx-` prefix. `hdocx_plan` reports the source
line, selector matches, normalized value, OOXML mapping, support status, and
patch ids.

Run formatting, used with `@hdocx-edit mode(all-runs);`:

| H-CSS declaration | Value | OOXML |
| --- | --- | --- |
| `hdocx-font-family` | quoted or bare font name | `w:rFonts @w:ascii` and `@w:hAnsi` |
| `hdocx-eastAsia-font` or `hdocx-east-asia-font` | quoted or bare font name | `w:rFonts @w:eastAsia` |
| `hdocx-ascii-font` | quoted or bare font name | `w:rFonts @w:ascii` |
| `hdocx-hansi-font` | quoted or bare font name | `w:rFonts @w:hAnsi` |
| `hdocx-cs-font` | quoted or bare font name | `w:rFonts @w:cs` |
| `hdocx-font-size` | positive `pt`, such as `10.5pt` | `w:sz` half-points |
| `hdocx-bold` | `true` or `false` | `w:b` |
| `hdocx-italic` | `true` or `false` | `w:i` |
| `hdocx-color` | `#RRGGBB` | `w:color @w:val` |

Paragraph formatting, used with `@hdocx-edit mode(paragraph-formatting);`:

| H-CSS declaration | Value | OOXML |
| --- | --- | --- |
| `hdocx-text-align` or `hdocx-align` | `left`, `center`, `right`, `justify`/`both` | `w:jc @w:val` |
| `hdocx-first-line-indent` | non-negative `char` or `pt`, such as `2char` | `w:ind @w:firstLineChars` or `@w:firstLine` |
| `hdocx-line-spacing` | positive multiple or exact `pt` | `w:spacing @w:line` and `@w:lineRule` |
| `hdocx-line-spacing-exact` | positive `pt`, such as `18pt` | `w:spacing @w:lineRule="exact"` |
| `hdocx-space-before` | `0`, non-negative `pt`, or `line`, such as `0.5line` | `w:spacing @w:before` or `@w:beforeLines` |
| `hdocx-space-after` | `0`, non-negative `pt`, or `line`, such as `0.5line` | `w:spacing @w:after` or `@w:afterLines` |
| `hdocx-manual-page-break-before` | `true` or `false` | insert an idempotent `<w:br w:type="page"/>` paragraph before the target |

Image formatting, used with `@hdocx-edit mode(image-formatting);`:

Use this for existing projected images. It targets drawing runs, or paragraphs
containing drawing runs. Host paragraph declarations are included because
academic body styles with exact line spacing can clip inline images.

| H-CSS declaration | Value | OOXML |
| --- | --- | --- |
| `hdocx-alt` | quoted or bare text | `wp:docPr @descr` |
| `hdocx-width-emu` | positive EMU integer | `wp:extent @cx` |
| `hdocx-height-emu` | positive EMU integer | `wp:extent @cy` |
| `hdocx-paragraph-line-spacing` | positive multiple or exact `pt` | host paragraph `w:spacing` |
| `hdocx-paragraph-line-spacing-exact` | positive `pt` | host paragraph exact `w:spacing` |
| `hdocx-paragraph-space-before` | `0`, non-negative `pt`, or `line` | host paragraph spacing before |
| `hdocx-paragraph-space-after` | `0`, non-negative `pt`, or `line` | host paragraph spacing after |
| `hdocx-paragraph-text-align` or `hdocx-paragraph-align` | `left`, `center`, `right`, `justify`/`both` | host paragraph alignment |

Paragraph structure, used with `@hdocx-edit mode(paragraph-structure);`:

Use this for real blank lines: H-DOCX inserts an empty Word paragraph, not just
spacing on a neighboring paragraph. `hdocx_diff` reports these changes under
`emptyParagraphDiff` and aligns later semantic nodes.

| H-CSS declaration | Value | OOXML |
| --- | --- | --- |
| `hdocx-insert-empty-paragraph-before` | `true` or `false` | idempotent empty `<w:p>` before the target |
| `hdocx-insert-empty-paragraph-after` | `true` or `false` | idempotent empty `<w:p>` after the target |
| `hdocx-empty-paragraph-style-id` | existing simple style id | inserted empty paragraph `w:pStyle` |
| `hdocx-empty-paragraph-line-spacing` | positive multiple or exact `pt` | inserted empty paragraph `w:spacing` |
| `hdocx-empty-paragraph-line-spacing-exact` | positive `pt`, such as `12pt` | inserted empty paragraph exact `w:spacing` |
| `hdocx-empty-paragraph-space-before` | `0`, non-negative `pt`, or `line` | inserted empty paragraph spacing before |
| `hdocx-empty-paragraph-space-after` | `0`, non-negative `pt`, or `line` | inserted empty paragraph spacing after |

Example for paper body text:

```css
@hdocx-set body {
  select: style(BodyText);
}

@hdocx-edit mode(paragraph-formatting);

body {
  hdocx-text-align: justify;
  hdocx-first-line-indent: 2char;
  hdocx-line-spacing-exact: 18pt;
  hdocx-space-before: 0;
  hdocx-space-after: 0;
}

@hdocx-edit mode(all-runs);

body {
  hdocx-font-family: "Times New Roman";
  hdocx-eastAsia-font: "SimSun";
  hdocx-font-size: 10.5pt;
}
```

Manual page break before a first-level heading. `hdocx_diff` reports these
changes under `manualPageBreakDiff` and aligns later semantic nodes so inserted
break paragraphs do not look like text edits:

```css
@hdocx-edit mode(paragraph-formatting);

.hdocx-p[data-hdocx-id="p-000006"] {
  hdocx-manual-page-break-before: true;
}
```

Real blank line after a paragraph:

```css
@hdocx-edit mode(paragraph-structure);

#p-000010 {
  hdocx-insert-empty-paragraph-after: true;
  hdocx-empty-paragraph-line-spacing-exact: 12pt;
  hdocx-empty-paragraph-space-before: 0;
  hdocx-empty-paragraph-space-after: 0;
}
```

Existing image that should not be clipped by body fixed line spacing:

```css
@hdocx-edit mode(image-formatting);

#r-000001 {
  hdocx-width-emu: 1828800;
  hdocx-height-emu: 914400;
  hdocx-paragraph-line-spacing: 1;
  hdocx-paragraph-space-before: 0;
  hdocx-paragraph-space-after: 0;
}
```

Always run `hdocx_plan` and inspect the resulting target set, declaration
diagnostics, and patch list before applying.
"""


ACCEPTANCE_GUIDANCE = """# H-DOCX Acceptance Checks

Minimum acceptance before claiming success:

- Source changes: run the unit tests.
- New or unknown DOCX family: run `hdocx_audit` and `hdocx_check`.
- Edited DOCX: run `hdocx_plan`, `hdocx_apply`, and `hdocx_diff`.
- Format-only tasks: run `hdocx_assert` with relevant assertions, especially
  `text-payload-unchanged` and image/blank-paragraph assertions.
- Batch or conversion-logic changes: run generated fixture pressure checks.

For unedited round-trips, success means byte-identical package output. For
edited round-trips, untouched OOXML must remain unchanged and every controlled
change must be reported.
"""


EDGE_CASE_GUIDANCE = """# H-DOCX Edge Cases

Treat these structures conservatively unless a dedicated operation supports the
requested change:

- tracked changes, comments, bookmarks, cross references, fields, citations,
  content controls, headers, footers, footnotes, endnotes, equations, drawings,
  charts, SmartArt, embedded objects, VML, macros, and custom XML.
- mixed formatting inside one paragraph, including multiple font sizes,
  East Asian and Latin font splits, superscript/subscript, hyperlinks, and
  field-code runs.
- table cell merges, nested tables, floating images, anchored drawings, and
  list numbering inherited from styles.

If the requested edit touches an unsupported or protected island, preserve it
exactly and return a clear report instead of approximating.
"""


def _guidance_tool(args: dict[str, Any]) -> dict[str, Any]:
    topic = str(args.get("topic", "all"))
    topic_resources = {
        "workflow": ["hdocx://guide/workflow"],
        "format": ["hdocx://guide/writing-format"],
        "query": ["hdocx://guide/query"],
        "hcss": ["hdocx://guide/hcss"],
        "acceptance": ["hdocx://guide/acceptance"],
        "edge-cases": ["hdocx://guide/edge-cases"],
        "all": [
            "hdocx://guide/workflow",
            "hdocx://guide/writing-format",
            "hdocx://guide/query",
            "hdocx://guide/hcss",
            "hdocx://guide/acceptance",
            "hdocx://guide/edge-cases",
        ],
    }
    uris = topic_resources.get(topic)
    if uris is None:
        raise HDocxError("MCP_INVALID_ARGUMENT", "Unknown guidance topic.", {"topic": topic})
    resources = _build_resources()
    selected = [resources[uri] for uri in uris]
    return {
        "ok": True,
        "topic": topic,
        "guidance": "\n\n".join(resource.text for resource in selected),
        "guidanceRuntime": _guidance_runtime_info(),
        "resources": [
            {
                "uri": resource.uri,
                "name": resource.name,
                "description": resource.description,
            }
            for resource in selected
        ],
        "prompts": ["hdocx_create_docx", "hdocx_safe_edit", "hdocx_format_change", "hdocx_roundtrip_check"],
    }


def _prompt_arg(name: str, description: str, *, required: bool) -> dict[str, Any]:
    return {"name": name, "description": description, "required": required}


def _prompt_value(args: dict[str, Any], name: str, default: str) -> str:
    value = args.get(name)
    if value is None or value == "":
        return default
    return str(value)


def _safe_edit_prompt(args: dict[str, Any]) -> str:
    root = _prompt_value(args, "root", "<workspace-root>")
    input_docx = _prompt_value(args, "input_docx", "<input.docx>")
    goal = _prompt_value(args, "goal", "<edit goal>")
    work = _prompt_value(args, "work", "work.hdocx")
    out = _prompt_value(args, "out", "output.docx")
    return f"""Use H-DOCX to safely edit this DOCX.

Workspace root: {root}
Input DOCX: {input_docx}
Work bundle: {work}
Output DOCX: {out}
Goal: {goal}

Before editing, read these MCP resources:
- hdocx://guide/workflow
- hdocx://guide/writing-format
- hdocx://guide/query
- hdocx://guide/hcss
- hdocx://guide/acceptance

Required calls:
1. hdocx_audit(root, input={input_docx})
2. hdocx_export(root, input={input_docx}, out={work}, force=true)
3. hdocx_query(root, bundle={work}, ...) or hdocx_find(...) to discover targets
4. hdocx_inspect(root, bundle={work}, ...) for every nontrivial exact target
5. Edit only supported H-DOCX surfaces.
6. hdocx_plan(root, bundle={work})
7. hdocx_apply(root, bundle={work}, out={out})
8. hdocx_diff(root, left={input_docx}, right={out})
9. hdocx_assert(root, bundle={work}, assertions=[...]) for task-specific invariants

Never change read-only ids, manifest metadata, original/original.docx, or
protected placeholders. If the edit cannot be proven safe, stop with a report.
"""


def _create_docx_prompt(args: dict[str, Any]) -> str:
    root = _prompt_value(args, "root", "<workspace-root>")
    out = _prompt_value(args, "out", "created.docx")
    goal = _prompt_value(args, "goal", "<new document goal>")
    title = _prompt_value(args, "title", "<optional title>")
    work = _prompt_value(args, "work", "created.hdocx")
    return f"""Create a new DOCX with H-DOCX.

Workspace root: {root}
Output DOCX: {out}
Optional H-DOCX bundle: {work}
Title: {title}
Goal: {goal}

Use `hdocx_create`, not an ad hoc ZIP writer. The tool creates a canonical DOCX
from the built-in blank template. Provide body text as `paragraphs` when the
content is known. Set `exportTo={work}` when you want to continue editing through
`document.html` and `agent.edits.hcss`.

Recommended calls:
1. hdocx_create(root, out={out}, title=..., paragraphs=[...], exportTo={work}, force=true)
2. hdocx_plan(root, bundle={work}) if exported and further H-DOCX edits were made
3. hdocx_apply(root, bundle={work}, out={out}) after H-DOCX edits
4. hdocx_check(root, input={out}, work=<check bundle>, out=<checked docx>, force=true)

Keep all paths inside the workspace root.
"""


def _format_change_prompt(args: dict[str, Any]) -> str:
    root = _prompt_value(args, "root", "<workspace-root>")
    input_docx = _prompt_value(args, "input_docx", "<input.docx>")
    goal = _prompt_value(args, "goal", "<formatting goal>")
    target = _prompt_value(args, "target", "<inspect first; prefer id/style/list/part selector>")
    work = _prompt_value(args, "work", "format-work.hdocx")
    out = _prompt_value(args, "out", "format-output.docx")
    return f"""Apply a formatting change with H-CSS, not ordinary CSS.

Workspace root: {root}
Input DOCX: {input_docx}
Work bundle: {work}
Output DOCX: {out}
Target: {target}
Goal: {goal}

Read hdocx://guide/query, hdocx://guide/hcss and
hdocx://guide/writing-format first.

Workflow:
1. hdocx_audit(root, input={input_docx})
2. hdocx_export(root, input={input_docx}, out={work}, force=true)
3. hdocx_query(root, bundle={work}, ...) to find candidate targets
4. hdocx_inspect(root, bundle={work}, kind=node/style/list/table/image, id=...)
5. Write the narrowest supported H-CSS rule in {work}/agent.edits.hcss.
6. hdocx_plan(root, bundle={work}) and verify target counts.
7. hdocx_apply(root, bundle={work}, out={out})
8. hdocx_diff(root, left={input_docx}, right={out})
9. hdocx_assert(root, bundle={work}, assertions=["text-payload-unchanged", ...])

Use inspected ids or @hdocx-set definitions. If a selector may match nothing,
declare allow-empty: true. Do not approximate unsupported formatting.
"""


def _roundtrip_prompt(args: dict[str, Any]) -> str:
    root = _prompt_value(args, "root", "<workspace-root>")
    input_docx = _prompt_value(args, "input_docx", "<input.docx>")
    work = _prompt_value(args, "work", "check.hdocx")
    out = _prompt_value(args, "out", "checked.docx")
    return f"""Prove this DOCX can round-trip through H-DOCX without edits.

Workspace root: {root}
Input DOCX: {input_docx}
Work bundle: {work}
Output DOCX: {out}

Call:
hdocx_check(root, input={input_docx}, work={work}, out={out}, force=true)

Success requires byteIdentical=true, contentIdentical=true, semanticIdentical=true,
and zero changed/left-only/right-only entries. Report SHA256 values in the final
answer.
"""


def _schema(*, properties: dict[str, Any], required: list[str] | None = None) -> dict[str, Any]:
    return {
        "type": "object",
        "properties": properties,
        "required": required or [],
        "additionalProperties": False,
    }


def _query_tool_schema() -> dict[str, Any]:
    return _schema(
        required=["bundle"],
        properties={
            "root": _string(ROOT_DESCRIPTION),
            "bundle": _string("H-DOCX bundle directory under root."),
            "kind": {
                "type": "string",
                "enum": ["paragraph", "run", "image", "all"],
                "description": "Node kind to query. Defaults to paragraph.",
            },
            "text": _string("Substring to find in visible run text or paragraph text."),
            "textRegex": _string("Regular expression to match visible text."),
            "styleId": _string("Exact paragraph style id."),
            "fontFamily": _string("Exact run font family; checks latin and east Asia font slots."),
            "eastAsiaFont": _string("Exact run east Asia font."),
            "fontSize": _string("Exact run font size, e.g. 14pt."),
            "bold": _boolean("Require bold true/false."),
            "italic": _boolean("Require italic true/false."),
            "color": _string("Exact run color, e.g. #ff0000."),
            "align": _string("Exact paragraph alignment such as center, left, right, justify."),
            "lineSpacing": _string("Exact paragraph line spacing, e.g. 1.5 or 18pt."),
            "spaceBefore": _string("Exact paragraph space before, e.g. 0pt or 0.5line."),
            "spaceAfter": _string("Exact paragraph space after, e.g. 0pt or 0.5line."),
            "hasImage": _boolean("Require paragraph/run to contain an image drawing."),
            "suspectedHeadingLevel1": _boolean("Filter by heuristic likely level-1 headings."),
            "includeRuns": _boolean("Include run summaries for paragraph matches. Defaults true."),
            "limit": {"type": "integer", "minimum": 1, "description": "Maximum matches to return. Defaults to 100."},
            "report": _string("Optional JSON report path under root."),
        },
    )


def _query_tool_kwargs(args: dict[str, Any]) -> dict[str, Any]:
    return {
        "kind": str(args.get("kind", "paragraph")),
        "text": args.get("text"),
        "text_regex": args.get("textRegex"),
        "style_id": args.get("styleId"),
        "font_family": args.get("fontFamily"),
        "east_asia_font": args.get("eastAsiaFont"),
        "font_size": args.get("fontSize"),
        "bold": args.get("bold"),
        "italic": args.get("italic"),
        "color": args.get("color"),
        "align": args.get("align"),
        "line_spacing": args.get("lineSpacing"),
        "space_before": args.get("spaceBefore"),
        "space_after": args.get("spaceAfter"),
        "has_image": args.get("hasImage"),
        "suspected_heading_level1": args.get("suspectedHeadingLevel1"),
        "include_runs": bool(args.get("includeRuns", True)),
        "limit": int(args.get("limit", 100)),
    }


def _assertions_arg(args: dict[str, Any]) -> list[Any] | None:
    value = args.get("assertions")
    if value is None:
        return None
    if not isinstance(value, list):
        raise HDocxError("MCP_INVALID_ARGUMENT", "assertions must be an array.", {"argument": "assertions"})
    for item in value:
        if not isinstance(item, (str, dict)):
            raise HDocxError("MCP_INVALID_ARGUMENT", "Each assertion must be a string or object.", {"argument": "assertions"})
    return value


def _string(description: str) -> dict[str, Any]:
    return {"type": "string", "description": description}


def _boolean(description: str) -> dict[str, Any]:
    return {"type": "boolean", "description": description}


def _path_arg(args: dict[str, Any], name: str, *, must_exist: bool = False) -> Path:
    value = args.get(name)
    if not isinstance(value, str) or not value:
        raise HDocxError("MCP_INVALID_ARGUMENT", f"Missing or invalid path argument: {name}.", {"argument": name})
    _validate_path_text(value, name)
    root = _root_arg(args)
    candidate = Path(value)
    if not candidate.is_absolute():
        candidate = root / candidate
    resolved = candidate.resolve()
    _ensure_under_root(resolved, root, argument=name)
    if must_exist and not resolved.exists():
        raise HDocxError("MCP_PATH_NOT_FOUND", f"Path does not exist: {name}.", {"argument": name, "path": str(resolved)})
    return resolved


def _string_list_arg(args: dict[str, Any], name: str) -> list[str] | None:
    value = args.get(name)
    if value is None:
        return None
    if not isinstance(value, list) or any(not isinstance(item, str) for item in value):
        raise HDocxError("MCP_INVALID_ARGUMENT", f"{name} must be an array of strings.", {"argument": name})
    return value


def _optional_string_arg(args: dict[str, Any], name: str) -> str | None:
    value = args.get(name)
    if value is None:
        return None
    if not isinstance(value, str):
        raise HDocxError("MCP_INVALID_ARGUMENT", f"{name} must be a string.", {"argument": name})
    return value


def _root_arg(args: dict[str, Any]) -> Path:
    raw_root = args.get("root")
    if raw_root is None:
        raw_root = os.environ.get("HDOCX_MCP_ROOT") or os.environ.get("CLAUDE_PROJECT_DIR")
    if raw_root is None:
        return Path.cwd().resolve()
    if not isinstance(raw_root, str) or not raw_root:
        raise HDocxError("MCP_INVALID_ARGUMENT", "root must be a non-empty string.", {"argument": "root"})
    _validate_path_text(raw_root, "root")
    return Path(raw_root).resolve()


def _validate_path_text(value: str, argument: str) -> None:
    try:
        value.encode("utf-8")
    except UnicodeEncodeError as exc:
        raise HDocxError(
            "PATH_ENCODING_ERROR",
            "Path contains characters that cannot be encoded safely.",
            {"argument": argument, "pathRepr": repr(value)},
        ) from exc


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
    text = json.dumps(payload, ensure_ascii=True, indent=2, sort_keys=True)
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
    sys.stdout.write(json.dumps(message, ensure_ascii=True, separators=(",", ":")))
    sys.stdout.write("\n")
    sys.stdout.flush()


if __name__ == "__main__":
    raise SystemExit(main())
