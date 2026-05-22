from __future__ import annotations

import argparse
import json
import sys
from pathlib import Path
from typing import Any, Callable

from .errors import HDocxError
from .fixtures import generate_pressure_fixtures
from .hdocx import (
    apply_hdocx,
    audit_docx,
    batch_check_docx,
    check_docx,
    create_docx,
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


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(prog="html-docx", description="H-DOCX reversible DOCX editing CLI")
    subparsers = parser.add_subparsers(dest="command", required=True)

    create_parser = subparsers.add_parser("create", help="Create a new canonical DOCX")
    create_parser.add_argument("--out", required=True, type=Path)
    create_parser.add_argument("--title")
    create_parser.add_argument("--paragraph", action="append", dest="paragraphs")
    create_parser.add_argument("--template", default="blank", choices=["blank"])
    create_parser.add_argument("--force", action="store_true")
    create_parser.add_argument("--export-to", type=Path)
    create_parser.add_argument("--report", type=Path)
    create_parser.set_defaults(func=_cmd_create)

    export_parser = subparsers.add_parser("export", help="Export DOCX to an H-DOCX bundle")
    export_parser.add_argument("input", type=Path)
    export_parser.add_argument("--out", required=True, type=Path)
    export_parser.add_argument("--force", action="store_true")
    export_parser.add_argument("--report", type=Path)
    export_parser.set_defaults(func=_cmd_export)

    validate_parser = subparsers.add_parser("validate", help="Validate an H-DOCX bundle")
    validate_parser.add_argument("bundle", type=Path)
    validate_parser.add_argument("--report", type=Path)
    validate_parser.set_defaults(func=_cmd_validate)

    plan_parser = subparsers.add_parser("plan", help="Plan H-DOCX edits without writing DOCX")
    plan_parser.add_argument("bundle", type=Path)
    plan_parser.add_argument("--report", type=Path)
    plan_parser.set_defaults(func=_cmd_plan)

    apply_parser = subparsers.add_parser("apply", help="Apply an H-DOCX bundle to DOCX")
    apply_parser.add_argument("bundle", type=Path)
    apply_parser.add_argument("--out", required=True, type=Path)
    apply_parser.add_argument("--report", type=Path)
    apply_parser.set_defaults(func=_cmd_apply)

    roundtrip_parser = subparsers.add_parser("roundtrip", help="Export and apply without edits")
    roundtrip_parser.add_argument("input", type=Path)
    roundtrip_parser.add_argument("--work", required=True, type=Path)
    roundtrip_parser.add_argument("--out", required=True, type=Path)
    roundtrip_parser.add_argument("--force", action="store_true")
    roundtrip_parser.add_argument("--report", type=Path)
    roundtrip_parser.set_defaults(func=_cmd_roundtrip)

    check_parser = subparsers.add_parser("check", help="Run export/apply/diff acceptance for a DOCX")
    check_parser.add_argument("input", type=Path)
    check_parser.add_argument("--work", required=True, type=Path)
    check_parser.add_argument("--out", required=True, type=Path)
    check_parser.add_argument("--force", action="store_true")
    check_parser.add_argument("--report", type=Path)
    check_parser.set_defaults(func=_cmd_check)

    batch_check_parser = subparsers.add_parser("batch-check", help="Run check over a DOCX file or directory")
    batch_check_parser.add_argument("input", type=Path)
    batch_check_parser.add_argument("--work", required=True, type=Path)
    batch_check_parser.add_argument("--out", required=True, type=Path)
    batch_check_parser.add_argument("--force", action="store_true")
    batch_check_parser.add_argument("--report", type=Path)
    batch_check_parser.set_defaults(func=_cmd_batch_check)

    generate_fixtures_parser = subparsers.add_parser("generate-fixtures", help="Generate local DOCX pressure fixtures")
    generate_fixtures_parser.add_argument("--out", required=True, type=Path)
    generate_fixtures_parser.add_argument("--force", action="store_true")
    generate_fixtures_parser.add_argument("--report", type=Path)
    generate_fixtures_parser.set_defaults(func=_cmd_generate_fixtures)

    diff_parser = subparsers.add_parser("diff", help="Compare DOCX package entries")
    diff_parser.add_argument("left", type=Path)
    diff_parser.add_argument("right", type=Path)
    diff_parser.add_argument("--report", type=Path)
    diff_parser.set_defaults(func=_cmd_diff)

    audit_parser = subparsers.add_parser("audit", help="Audit high-risk DOCX structures")
    audit_parser.add_argument("input", type=Path)
    audit_parser.add_argument("--report", type=Path)
    audit_parser.set_defaults(func=_cmd_audit)

    render_parser = subparsers.add_parser("render-check", help="Optionally render DOCX to PDF with LibreOffice/soffice")
    render_parser.add_argument("input", type=Path)
    render_parser.add_argument("--out", required=True, type=Path)
    render_parser.add_argument("--force", action="store_true")
    render_parser.add_argument("--allow-missing", action="store_true")
    render_parser.add_argument("--timeout", type=int, default=120)
    render_parser.add_argument("--report", type=Path)
    render_parser.set_defaults(func=_cmd_render_check)

    inspect_parser = subparsers.add_parser("inspect", help="Inspect an H-DOCX manifest item")
    inspect_parser.add_argument("bundle", type=Path)
    inspect_parser.add_argument("--kind", choices=["node", "style", "list", "table", "image"], default="node")
    inspect_parser.add_argument("--id", required=True, dest="target_id")
    inspect_parser.add_argument("--report", type=Path)
    inspect_parser.set_defaults(func=_cmd_inspect)

    doctor_parser = subparsers.add_parser("doctor", help="Report local html-docx runtime capabilities")
    doctor_parser.add_argument("--report", type=Path)
    doctor_parser.set_defaults(func=_cmd_doctor)

    mcp_parser = subparsers.add_parser("mcp", help="Run the H-DOCX MCP stdio server")
    mcp_parser.set_defaults(func=_cmd_mcp)

    return parser


def main(argv: list[str] | None = None) -> int:
    parser = build_parser()
    args = parser.parse_args(argv)
    try:
        report = args.func(args)
    except HDocxError as exc:
        _print_json({"ok": False, "error": exc.to_dict()})
        return 2
    except Exception as exc:
        _print_json(
            {
                "ok": False,
                "error": {
                    "code": "UNEXPECTED_ERROR",
                    "severity": "error",
                    "message": str(exc),
                },
            }
        )
        return 1
    if getattr(args, "report", None):
        write_json(args.report, report)
    _print_json(report)
    return 0 if report.get("ok", False) else 2


def _cmd_export(args: argparse.Namespace) -> dict[str, Any]:
    return export_docx(args.input, args.out, force=args.force)


def _cmd_create(args: argparse.Namespace) -> dict[str, Any]:
    return create_docx(
        args.out,
        title=args.title,
        paragraphs=args.paragraphs,
        template=args.template,
        force=args.force,
        export_dir=args.export_to,
    )


def _cmd_validate(args: argparse.Namespace) -> dict[str, Any]:
    return validate_hdocx(args.bundle)


def _cmd_plan(args: argparse.Namespace) -> dict[str, Any]:
    return plan_hdocx(args.bundle)


def _cmd_apply(args: argparse.Namespace) -> dict[str, Any]:
    return apply_hdocx(args.bundle, args.out)


def _cmd_roundtrip(args: argparse.Namespace) -> dict[str, Any]:
    return roundtrip_docx(args.input, args.work, args.out, force=args.force)


def _cmd_check(args: argparse.Namespace) -> dict[str, Any]:
    return check_docx(args.input, args.work, args.out, force=args.force)


def _cmd_batch_check(args: argparse.Namespace) -> dict[str, Any]:
    return batch_check_docx(args.input, args.work, args.out, force=args.force)


def _cmd_generate_fixtures(args: argparse.Namespace) -> dict[str, Any]:
    return generate_pressure_fixtures(args.out, force=args.force)


def _cmd_diff(args: argparse.Namespace) -> dict[str, Any]:
    return diff_docx(args.left, args.right)


def _cmd_audit(args: argparse.Namespace) -> dict[str, Any]:
    return audit_docx(args.input)


def _cmd_render_check(args: argparse.Namespace) -> dict[str, Any]:
    return render_check_docx(args.input, args.out, force=args.force, allow_missing=args.allow_missing, timeout_seconds=args.timeout)


def _cmd_inspect(args: argparse.Namespace) -> dict[str, Any]:
    return inspect_hdocx(args.bundle, args.target_id, kind=args.kind)


def _cmd_doctor(args: argparse.Namespace) -> dict[str, Any]:
    return doctor_report()


def _cmd_mcp(args: argparse.Namespace) -> dict[str, Any]:
    from .mcp_server import main as mcp_main

    raise SystemExit(mcp_main())


def _print_json(payload: dict[str, Any]) -> None:
    text = json.dumps(payload, ensure_ascii=False, indent=2, sort_keys=True) + "\n"
    stdout_buffer = getattr(sys.stdout, "buffer", None)
    if stdout_buffer is not None:
        stdout_buffer.write(text.encode("utf-8"))
        stdout_buffer.flush()
        return
    sys.stdout.write(text)
