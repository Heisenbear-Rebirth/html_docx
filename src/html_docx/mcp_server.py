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
    server = HDocxMcpServer()
    return server.run()


class HDocxMcpServer:
    def __init__(self) -> None:
        self.tools = _build_tools()
        self.resources = _build_resources()
        self.prompts = _build_prompts()

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


def _build_tools() -> dict[str, Tool]:
    tools = [
        Tool(
            "hdocx_doctor",
            "Report html_docx runtime capabilities and optional renderer availability.",
            _schema(properties={}),
            lambda args: doctor_report(),
        ),
        Tool(
            "hdocx_guidance",
            "Return H-DOCX authoring rules, H-CSS format guidance, and required acceptance checks.",
            _schema(
                properties={
                    "topic": {
                        "type": "string",
                        "enum": ["all", "workflow", "format", "hcss", "acceptance", "edge-cases"],
                        "description": "Guidance topic. Defaults to all.",
                    }
                }
            ),
            _guidance_tool,
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
            HCSS_GUIDANCE,
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

1. Run `hdocx_audit` on the source DOCX.
2. Run `hdocx_export` to create a `.hdocx` bundle.
3. Run `hdocx_inspect` before targeting styles, lists, tables, images, or
   specific paragraphs/runs.
4. Edit only allowed bundle surfaces: editable run text in `document.html`,
   supported rules in `agent.edits.hcss`, and supported bundle-local assets.
5. Run `hdocx_plan` before writing an output DOCX.
6. Run `hdocx_apply` to produce the output DOCX.
7. Run `hdocx_diff` against the original before claiming success.

For no-edit reversibility proof, use `hdocx_check` directly. A byte-identical
SHA256 match is stronger than render equality for unedited round-trips.
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


HCSS_GUIDANCE = """# H-CSS Selector And Formatting Guidance

Use `agent.edits.hcss` for controlled formatting edits. Prefer inspected ids or
named sets over broad selectors.

Recommended targeting patterns:

```css
@hdocx-set target-paragraph {
  select: id(p-000001);
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

Do not use broad selectors when the user asked for a narrow change. Do not
write normal CSS properties unless H-DOCX explicitly supports them. Always run
`hdocx_plan` and inspect the resulting target set before applying.
"""


ACCEPTANCE_GUIDANCE = """# H-DOCX Acceptance Checks

Minimum acceptance before claiming success:

- Source changes: run the unit tests.
- New or unknown DOCX family: run `hdocx_audit` and `hdocx_check`.
- Edited DOCX: run `hdocx_plan`, `hdocx_apply`, and `hdocx_diff`.
- Formatting-sensitive work: run `hdocx_render_check` when LibreOffice/soffice
  is available, or report that render QA was unavailable.
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
        "hcss": ["hdocx://guide/hcss"],
        "acceptance": ["hdocx://guide/acceptance"],
        "edge-cases": ["hdocx://guide/edge-cases"],
        "all": [
            "hdocx://guide/workflow",
            "hdocx://guide/writing-format",
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
        "resources": [
            {
                "uri": resource.uri,
                "name": resource.name,
                "description": resource.description,
            }
            for resource in selected
        ],
        "prompts": ["hdocx_safe_edit", "hdocx_format_change", "hdocx_roundtrip_check"],
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
- hdocx://guide/hcss
- hdocx://guide/acceptance

Required calls:
1. hdocx_audit(root, input={input_docx})
2. hdocx_export(root, input={input_docx}, out={work}, force=true)
3. hdocx_inspect(root, bundle={work}, ...) for every nontrivial target
4. Edit only supported H-DOCX surfaces.
5. hdocx_plan(root, bundle={work})
6. hdocx_apply(root, bundle={work}, out={out})
7. hdocx_diff(root, left={input_docx}, right={out})

Never change read-only ids, manifest metadata, original/original.docx, or
protected placeholders. If the edit cannot be proven safe, stop with a report.
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

Read hdocx://guide/hcss and hdocx://guide/writing-format first.

Workflow:
1. hdocx_audit(root, input={input_docx})
2. hdocx_export(root, input={input_docx}, out={work}, force=true)
3. hdocx_inspect(root, bundle={work}, kind=node/style/list/table/image, id=...)
4. Write the narrowest supported H-CSS rule in {work}/agent.edits.hcss.
5. hdocx_plan(root, bundle={work}) and verify target counts.
6. hdocx_apply(root, bundle={work}, out={out})
7. hdocx_diff(root, left={input_docx}, right={out})

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
