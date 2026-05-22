from __future__ import annotations

import json
import os
import shutil
import subprocess
import sys
import unittest
from contextlib import redirect_stdout
from io import StringIO
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "src"))
sys.path.insert(0, str(ROOT / "tests"))

from html_docx.cli import main
from html_docx.mcp_server import HDocxMcpServer
from test_roundtrip import (
    _make_audit_feature_docx,
    _make_complex_academic_docx,
    _make_image_docx,
    _make_minimal_docx,
    _make_numbered_docx,
    _make_styled_docx,
    _make_two_cell_table_docx,
)


TMP = ROOT / "tests" / "_tmp_cli"


class CLITests(unittest.TestCase):
    def setUp(self) -> None:
        if TMP.exists():
            shutil.rmtree(TMP)
        TMP.mkdir(parents=True)

    def tearDown(self) -> None:
        if TMP.exists():
            shutil.rmtree(TMP)

    def test_plan_report_file(self) -> None:
        input_docx = TMP / "input.docx"
        work = TMP / "work.hdocx"
        report_path = TMP / "plan.json"
        _make_minimal_docx(input_docx)

        with redirect_stdout(StringIO()):
            self.assertEqual(main(["export", str(input_docx), "--out", str(work)]), 0)
        with redirect_stdout(StringIO()):
            self.assertEqual(main(["plan", str(work), "--report", str(report_path)]), 0)

        report = json.loads(report_path.read_text(encoding="utf-8"))
        self.assertTrue(report["ok"], report)
        self.assertEqual(report["command"], "plan")

    def test_diff_report_file_contains_entry_details(self) -> None:
        input_docx = TMP / "input.docx"
        work = TMP / "work.hdocx"
        output_docx = TMP / "output.docx"
        report_path = TMP / "diff.json"
        _make_minimal_docx(input_docx)

        with redirect_stdout(StringIO()):
            self.assertEqual(main(["export", str(input_docx), "--out", str(work)]), 0)
        html_path = work / "document.html"
        html_path.write_text(
            html_path.read_text(encoding="utf-8").replace("Hello", "Changed"),
            encoding="utf-8",
            newline="\n",
        )
        with redirect_stdout(StringIO()):
            self.assertEqual(main(["apply", str(work), "--out", str(output_docx)]), 0)
        with redirect_stdout(StringIO()):
            self.assertEqual(main(["diff", str(input_docx), str(output_docx), "--report", str(report_path)]), 0)

        report = json.loads(report_path.read_text(encoding="utf-8"))
        self.assertTrue(report["ok"], report)
        self.assertEqual(report["changed"], ["word/document.xml"])
        self.assertEqual(report["entryCounts"]["changed"], 1)
        self.assertEqual(report["changedEntries"][0]["path"], "word/document.xml")
        self.assertEqual(report["changedEntries"][0]["kind"], "main-document")
        self.assertTrue(report["semanticDiff"]["available"])
        self.assertIn("r-000001", report["semanticDiff"]["changed"])
        self.assertTrue(report["fragmentDiff"]["available"])
        self.assertEqual(report["fragmentDiff"]["entries"][0]["path"], "word/document.xml")
        self.assertIn("r-000001", [item["nodeId"] for item in report["fragmentDiff"]["entries"][0]["linkedNodes"]])

    def test_check_command_runs_acceptance_chain(self) -> None:
        input_docx = TMP / "input.docx"
        work = TMP / "check.hdocx"
        output_docx = TMP / "checked.docx"
        report_path = TMP / "check.json"
        _make_minimal_docx(input_docx)

        with redirect_stdout(StringIO()):
            self.assertEqual(
                main(["check", str(input_docx), "--work", str(work), "--out", str(output_docx), "--force", "--report", str(report_path)]),
                0,
            )

        report = json.loads(report_path.read_text(encoding="utf-8"))
        self.assertTrue(report["ok"], report)
        self.assertEqual(report["command"], "check")
        self.assertTrue(report["acceptance"]["byteIdentical"])
        self.assertTrue(report["acceptance"]["contentIdentical"])
        self.assertTrue(report["acceptance"]["semanticIdentical"])
        self.assertEqual(report["roundtrip"]["command"], "roundtrip")
        self.assertEqual(report["diff"]["command"], "diff")

    def test_doctor_reports_runtime_without_external_dependencies(self) -> None:
        report_path = TMP / "doctor.json"

        with redirect_stdout(StringIO()):
            self.assertEqual(main(["doctor", "--report", str(report_path)]), 0)

        report = json.loads(report_path.read_text(encoding="utf-8"))
        self.assertTrue(report["ok"], report)
        self.assertEqual(report["command"], "doctor")
        self.assertEqual(report["package"]["dependencies"], [])
        self.assertFalse(report["capabilities"]["requiresExternalRuntime"])
        self.assertIn("rendering", report["capabilities"])
        self.assertIn("available", report["capabilities"]["rendering"])

    def test_create_command_can_export_new_docx(self) -> None:
        output_docx = TMP / "created.docx"
        work = TMP / "created.hdocx"
        report_path = TMP / "created.json"

        with redirect_stdout(StringIO()):
            self.assertEqual(
                main(
                    [
                        "create",
                        "--out",
                        str(output_docx),
                        "--title",
                        "Created From CLI",
                        "--paragraph",
                        "Body paragraph.",
                        "--export-to",
                        str(work),
                        "--force",
                        "--report",
                        str(report_path),
                    ]
                ),
                0,
            )

        report = json.loads(report_path.read_text(encoding="utf-8"))
        self.assertTrue(report["ok"], report)
        self.assertEqual(report["command"], "create")
        self.assertTrue(output_docx.exists())
        self.assertTrue((work / "manifest.json").exists())
        self.assertEqual(report["export"]["command"], "export")

    def test_mcp_returns_busy_instead_of_breaking_transport(self) -> None:
        server = HDocxMcpServer()
        self.assertTrue(server._tool_lock.acquire(blocking=False))
        try:
            response = server.handle_message(
                {
                    "jsonrpc": "2.0",
                    "id": 1,
                    "method": "tools/call",
                    "params": {"name": "hdocx_doctor", "arguments": {}},
                }
            )
        finally:
            server._tool_lock.release()

        result = response["result"]
        self.assertTrue(result["isError"], response)
        self.assertEqual(result["structuredContent"]["error"]["code"], "MCP_SERVER_BUSY")

    def test_plan_command_handles_bom_parse_error_without_stdout_crash(self) -> None:
        input_docx = TMP / "input.docx"
        work = TMP / "work.hdocx"
        _make_minimal_docx(input_docx)

        with redirect_stdout(StringIO()):
            self.assertEqual(main(["export", str(input_docx), "--out", str(work)]), 0)
        (work / "agent.edits.hcss").write_bytes(
            "\ufeff@hdocx-edit mode(all-runs);\n\n.hdocx-r, .other { hdocx-bold: true; }\n".encode("utf-8")
        )
        stdout = StringIO()
        with redirect_stdout(stdout):
            code = main(["plan", str(work)])

        self.assertEqual(code, 2)
        report = json.loads(stdout.getvalue())
        self.assertFalse(report["ok"], report)
        self.assertEqual(report["errors"][0]["code"], "HCSS_PARSE_UNSUPPORTED_SYNTAX")

    def test_audit_report_file_contains_advanced_object_summary(self) -> None:
        input_docx = TMP / "audit-features.docx"
        report_path = TMP / "audit.json"
        _make_audit_feature_docx(input_docx)

        with redirect_stdout(StringIO()):
            self.assertEqual(main(["audit", str(input_docx), "--report", str(report_path)]), 0)

        report = json.loads(report_path.read_text(encoding="utf-8"))
        self.assertTrue(report["ok"], report)
        self.assertEqual(report["command"], "audit")
        self.assertTrue(report["features"]["chart"]["present"])
        self.assertTrue(report["features"]["alternateContent"]["present"])
        self.assertTrue(report["summary"]["hasNonEditableAdvancedObjects"])

    def test_generate_fixtures_then_batch_check(self) -> None:
        fixture_dir = TMP / "generated-fixtures"
        fixture_report_path = TMP / "fixtures.json"
        batch_report_path = TMP / "generated-batch.json"

        with redirect_stdout(StringIO()):
            self.assertEqual(main(["generate-fixtures", "--out", str(fixture_dir), "--force", "--report", str(fixture_report_path)]), 0)
        with redirect_stdout(StringIO()):
            self.assertEqual(
                main(
                    [
                        "batch-check",
                        str(fixture_dir),
                        "--work",
                        str(TMP / "generated-work"),
                        "--out",
                        str(TMP / "generated-out"),
                        "--force",
                        "--report",
                        str(batch_report_path),
                    ]
                ),
                0,
            )

        fixture_report = json.loads(fixture_report_path.read_text(encoding="utf-8"))
        batch_report = json.loads(batch_report_path.read_text(encoding="utf-8"))
        self.assertEqual(fixture_report["count"], 6)
        self.assertEqual(batch_report["summary"], {"failed": 0, "passed": 6, "total": 6})

    def test_render_check_allow_missing_is_structured(self) -> None:
        input_docx = TMP / "input.docx"
        report_path = TMP / "render.json"
        old_path = os.environ.get("PATH")
        _make_minimal_docx(input_docx)
        try:
            os.environ["PATH"] = ""
            with redirect_stdout(StringIO()):
                self.assertEqual(
                    main(
                        [
                            "render-check",
                            str(input_docx),
                            "--out",
                            str(TMP / "render-out"),
                            "--force",
                            "--allow-missing",
                            "--report",
                            str(report_path),
                        ]
                    ),
                    0,
                )
        finally:
            if old_path is None:
                os.environ.pop("PATH", None)
            else:
                os.environ["PATH"] = old_path

        report = json.loads(report_path.read_text(encoding="utf-8"))
        self.assertTrue(report["ok"], report)
        self.assertEqual(report["command"], "render-check")
        self.assertFalse(report["available"])
        self.assertEqual(report["status"], "renderer-missing")

    def test_batch_check_runs_directory_acceptance(self) -> None:
        input_dir = TMP / "inputs"
        input_dir.mkdir()
        _make_minimal_docx(input_dir / "minimal.docx")
        _make_styled_docx(input_dir / "styled.docx")
        _make_complex_academic_docx(input_dir / "complex-academic.docx")
        report_path = TMP / "batch.json"

        with redirect_stdout(StringIO()):
            self.assertEqual(
                main(
                    [
                        "batch-check",
                        str(input_dir),
                        "--work",
                        str(TMP / "batch-work"),
                        "--out",
                        str(TMP / "batch-out"),
                        "--force",
                        "--report",
                        str(report_path),
                    ]
                ),
                0,
            )

        report = json.loads(report_path.read_text(encoding="utf-8"))
        self.assertTrue(report["ok"], report)
        self.assertEqual(report["command"], "batch-check")
        self.assertEqual(report["summary"], {"failed": 0, "passed": 3, "total": 3})
        self.assertEqual(
            [item["relativePath"] for item in report["results"]],
            ["complex-academic.docx", "minimal.docx", "styled.docx"],
        )
        self.assertTrue(all(item["acceptance"]["byteIdentical"] for item in report["results"]))

    def test_batch_check_can_keep_generated_dirs_inside_input_dir(self) -> None:
        input_dir = TMP / "fixtures"
        input_dir.mkdir()
        _make_minimal_docx(input_dir / "minimal.docx")
        args = [
            "batch-check",
            str(input_dir),
            "--work",
            str(input_dir / "_work.hdocx"),
            "--out",
            str(input_dir / "_out"),
            "--force",
            "--report",
            str(input_dir / "pressure.json"),
        ]

        with redirect_stdout(StringIO()):
            self.assertEqual(main(args), 0)
        with redirect_stdout(StringIO()):
            self.assertEqual(main(args), 0)
        report = json.loads((input_dir / "pressure.json").read_text(encoding="utf-8"))
        self.assertTrue(report["ok"], report)
        self.assertEqual(report["summary"], {"failed": 0, "passed": 1, "total": 1})
        self.assertEqual(report["results"][0]["relativePath"], "minimal.docx")

    def test_mcp_stdio_lists_and_calls_tools(self) -> None:
        env = os.environ.copy()
        env["PYTHONPATH"] = str(ROOT / "src")
        messages = [
            {"jsonrpc": "2.0", "id": 1, "method": "initialize", "params": {"protocolVersion": "2024-11-05", "capabilities": {}, "clientInfo": {"name": "tests", "version": "0"}}},
            {"jsonrpc": "2.0", "method": "notifications/initialized", "params": {}},
            {"jsonrpc": "2.0", "id": 2, "method": "tools/list", "params": {}},
            {"jsonrpc": "2.0", "id": 3, "method": "tools/call", "params": {"name": "hdocx_doctor", "arguments": {}}},
        ]
        stdin = "\n".join(json.dumps(item) for item in messages) + "\n"

        proc = subprocess.run(
            [sys.executable, "-m", "html_docx.mcp_server"],
            input=stdin,
            text=True,
            capture_output=True,
            env=env,
            cwd=ROOT,
            timeout=20,
        )

        self.assertEqual(proc.returncode, 0, proc.stderr)
        responses = [json.loads(line) for line in proc.stdout.splitlines() if line.strip()]
        self.assertEqual([item["id"] for item in responses], [1, 2, 3])
        self.assertEqual(responses[0]["result"]["capabilities"]["tools"]["listChanged"], False)
        tool_names = {tool["name"] for tool in responses[1]["result"]["tools"]}
        self.assertIn("hdocx_check", tool_names)
        self.assertIn("hdocx_create", tool_names)
        self.assertIn("hdocx_export", tool_names)
        self.assertIn("hdocx_guidance", tool_names)
        self.assertFalse(responses[2]["result"]["isError"])
        self.assertTrue(responses[2]["result"]["structuredContent"]["ok"])

    def test_mcp_create_tool_creates_and_exports_docx(self) -> None:
        from html_docx.mcp_server import HDocxMcpServer

        server = HDocxMcpServer()
        root = TMP / "mcp-create"
        root.mkdir()

        response = server.handle_message(
            {
                "jsonrpc": "2.0",
                "id": 20,
                "method": "tools/call",
                "params": {
                    "name": "hdocx_create",
                    "arguments": {
                        "root": str(root),
                        "out": "created.docx",
                        "title": "MCP Created",
                        "paragraphs": ["Created by MCP."],
                        "exportTo": "created.hdocx",
                        "force": True,
                    },
                },
            }
        )

        self.assertIsNotNone(response)
        result = response["result"]
        self.assertFalse(result["isError"], result)
        payload = result["structuredContent"]
        self.assertTrue(payload["ok"], payload)
        self.assertEqual(payload["command"], "create")
        self.assertEqual(payload["export"]["command"], "export")
        self.assertTrue((root / "created.docx").exists())
        self.assertTrue((root / "created.hdocx" / "manifest.json").exists())

    def test_mcp_exposes_guidance_resources_prompts_and_tool(self) -> None:
        from html_docx.mcp_server import HDocxMcpServer

        server = HDocxMcpServer()
        resources = server.handle_message({"jsonrpc": "2.0", "id": 1, "method": "resources/list", "params": {}})
        self.assertIsNotNone(resources)
        resource_items = resources["result"]["resources"]
        resource_uris = {item["uri"] for item in resource_items}
        self.assertIn("hdocx://guide/writing-format", resource_uris)
        self.assertIn("hdocx://guide/hcss", resource_uris)

        read = server.handle_message(
            {
                "jsonrpc": "2.0",
                "id": 2,
                "method": "resources/read",
                "params": {"uri": "hdocx://guide/writing-format"},
            }
        )
        self.assertIsNotNone(read)
        text = read["result"]["contents"][0]["text"]
        self.assertIn("document.html", text)
        self.assertIn("agent.edits.hcss", text)

        prompts = server.handle_message({"jsonrpc": "2.0", "id": 3, "method": "prompts/list", "params": {}})
        self.assertIsNotNone(prompts)
        prompt_names = {item["name"] for item in prompts["result"]["prompts"]}
        self.assertIn("hdocx_create_docx", prompt_names)
        self.assertIn("hdocx_safe_edit", prompt_names)
        self.assertIn("hdocx_format_change", prompt_names)

        prompt = server.handle_message(
            {
                "jsonrpc": "2.0",
                "id": 4,
                "method": "prompts/get",
                "params": {
                    "name": "hdocx_format_change",
                    "arguments": {
                        "root": str(TMP),
                        "input_docx": "input.docx",
                        "goal": "Set body text to 12 pt.",
                    },
                },
            }
        )
        self.assertIsNotNone(prompt)
        prompt_text = prompt["result"]["messages"][0]["content"]["text"]
        self.assertIn("hdocx_plan", prompt_text)
        self.assertIn("H-CSS", prompt_text)

        guidance = server.handle_message(
            {
                "jsonrpc": "2.0",
                "id": 5,
                "method": "tools/call",
                "params": {"name": "hdocx_guidance", "arguments": {"topic": "hcss"}},
            }
        )
        self.assertIsNotNone(guidance)
        payload = guidance["result"]["structuredContent"]
        self.assertTrue(payload["ok"])
        self.assertIn("@hdocx-set", payload["guidance"])

    def test_mcp_rejects_paths_outside_root(self) -> None:
        from html_docx.mcp_server import HDocxMcpServer

        server = HDocxMcpServer()
        root = TMP / "mcp-root"
        root.mkdir()
        outside = ROOT / "outside.docx"

        response = server.handle_message(
            {
                "jsonrpc": "2.0",
                "id": 10,
                "method": "tools/call",
                "params": {
                    "name": "hdocx_audit",
                    "arguments": {"root": str(root), "input": str(outside)},
                },
            }
        )

        self.assertIsNotNone(response)
        result = response["result"]
        self.assertTrue(result["isError"])
        payload = result["structuredContent"]
        self.assertEqual(payload["error"]["code"], "MCP_PATH_OUTSIDE_ROOT")

    def test_mcp_uses_environment_root_when_root_is_omitted(self) -> None:
        from html_docx.mcp_server import HDocxMcpServer

        server = HDocxMcpServer()
        root = TMP / "mcp-env-root"
        root.mkdir()
        _make_minimal_docx(root / "input.docx")
        previous = os.environ.get("HDOCX_MCP_ROOT")
        os.environ["HDOCX_MCP_ROOT"] = str(root)
        try:
            response = server.handle_message(
                {
                    "jsonrpc": "2.0",
                    "id": 11,
                    "method": "tools/call",
                    "params": {
                        "name": "hdocx_audit",
                        "arguments": {"input": "input.docx"},
                    },
                }
            )
        finally:
            if previous is None:
                os.environ.pop("HDOCX_MCP_ROOT", None)
            else:
                os.environ["HDOCX_MCP_ROOT"] = previous

        self.assertIsNotNone(response)
        result = response["result"]
        self.assertFalse(result["isError"], result)
        payload = result["structuredContent"]
        self.assertTrue(payload["ok"])
        self.assertEqual(payload["input"], str((root / "input.docx").resolve()))

    def test_inspect_style_list_table_and_image(self) -> None:
        styled_docx = TMP / "styled.docx"
        styled_work = TMP / "styled.hdocx"
        styled_report = TMP / "styled-inspect.json"
        _make_styled_docx(styled_docx)
        with redirect_stdout(StringIO()):
            self.assertEqual(main(["export", str(styled_docx), "--out", str(styled_work)]), 0)
        with redirect_stdout(StringIO()):
            self.assertEqual(
                main(["inspect", str(styled_work), "--kind", "style", "--id", "Normal", "--report", str(styled_report)]),
                0,
            )
        style_report = json.loads(styled_report.read_text(encoding="utf-8"))
        self.assertEqual(style_report["kind"], "style")
        self.assertEqual(style_report["style"]["styleId"], "Normal")
        self.assertEqual(style_report["usageCount"], 1)

        numbered_docx = TMP / "numbered.docx"
        numbered_work = TMP / "numbered.hdocx"
        numbered_report = TMP / "numbered-inspect.json"
        _make_numbered_docx(numbered_docx)
        with redirect_stdout(StringIO()):
            self.assertEqual(main(["export", str(numbered_docx), "--out", str(numbered_work)]), 0)
        with redirect_stdout(StringIO()):
            self.assertEqual(
                main(["inspect", str(numbered_work), "--kind", "list", "--id", "1", "--report", str(numbered_report)]),
                0,
            )
        list_report = json.loads(numbered_report.read_text(encoding="utf-8"))
        self.assertEqual(list_report["kind"], "list")
        self.assertEqual(list_report["list"]["num"]["numId"], "1")
        self.assertEqual(list_report["list"]["abstractNum"]["levels"]["0"]["numFmt"], "decimal")

        table_docx = TMP / "table.docx"
        table_work = TMP / "table.hdocx"
        table_report_path = TMP / "table-inspect.json"
        _make_two_cell_table_docx(table_docx)
        with redirect_stdout(StringIO()):
            self.assertEqual(main(["export", str(table_docx), "--out", str(table_work)]), 0)
        with redirect_stdout(StringIO()):
            self.assertEqual(
                main(["inspect", str(table_work), "--kind", "table", "--id", "tbl-000001", "--report", str(table_report_path)]),
                0,
            )
        table_report = json.loads(table_report_path.read_text(encoding="utf-8"))
        self.assertEqual(table_report["kind"], "table")
        self.assertEqual(table_report["rowCount"], 1)
        self.assertEqual(len(table_report["rows"][0]["cells"]), 2)

        image_docx = TMP / "image.docx"
        image_work = TMP / "image.hdocx"
        image_report_path = TMP / "image-inspect.json"
        _make_image_docx(image_docx)
        with redirect_stdout(StringIO()):
            self.assertEqual(main(["export", str(image_docx), "--out", str(image_work)]), 0)
        with redirect_stdout(StringIO()):
            self.assertEqual(
                main(["inspect", str(image_work), "--kind", "image", "--id", "r-000001", "--report", str(image_report_path)]),
                0,
            )
        image_report = json.loads(image_report_path.read_text(encoding="utf-8"))
        self.assertEqual(image_report["kind"], "image")
        self.assertEqual(image_report["image"]["objectKind"], "drawing")
        self.assertIn("/word/media/image1.png", image_report["mediaParts"])


if __name__ == "__main__":
    unittest.main()
