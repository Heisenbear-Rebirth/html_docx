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
    validate_hdocx,
)
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

    inspect_parser = subparsers.add_parser("inspect", help="Inspect an H-DOCX manifest item")
    inspect_parser.add_argument("bundle", type=Path)
    inspect_parser.add_argument("--kind", choices=["node", "style", "list", "table", "image"], default="node")
    inspect_parser.add_argument("--id", required=True, dest="target_id")
    inspect_parser.add_argument("--report", type=Path)
    inspect_parser.set_defaults(func=_cmd_inspect)

    query_parser = subparsers.add_parser("query", help="Query projected H-DOCX nodes by text and formatting")
    _add_query_arguments(query_parser)
    query_parser.set_defaults(func=_cmd_query)

    find_parser = subparsers.add_parser("find", help="Alias for query")
    _add_query_arguments(find_parser)
    find_parser.set_defaults(func=_cmd_find)

    assert_parser = subparsers.add_parser("assert", help="Run assertion-style H-DOCX acceptance checks")
    assert_parser.add_argument("bundle", type=Path)
    assert_parser.add_argument("--assertion", action="append", dest="assertions", help="Assertion type; may be repeated.")
    assert_parser.add_argument("--assertions-json", help="JSON array of assertion strings or objects.")
    assert_parser.add_argument("--report", type=Path)
    assert_parser.set_defaults(func=_cmd_assert)

    doctor_parser = subparsers.add_parser("doctor", help="Report local html-docx runtime capabilities")
    doctor_parser.add_argument("--report", type=Path)
    doctor_parser.set_defaults(func=_cmd_doctor)

    mcp_parser = subparsers.add_parser("mcp", help="Run the H-DOCX MCP stdio server")
    mcp_parser.set_defaults(func=_cmd_mcp)

    return parser


def _add_query_arguments(parser: argparse.ArgumentParser) -> None:
    parser.add_argument("bundle", type=Path)
    parser.add_argument("--kind", choices=["paragraph", "run", "image", "all"], default="paragraph")
    parser.add_argument("--text")
    parser.add_argument("--text-regex")
    parser.add_argument("--style-id")
    parser.add_argument("--font-family")
    parser.add_argument("--east-asia-font")
    parser.add_argument("--font-size")
    parser.add_argument("--bold", action=argparse.BooleanOptionalAction)
    parser.add_argument("--italic", action=argparse.BooleanOptionalAction)
    parser.add_argument("--color")
    parser.add_argument("--align")
    parser.add_argument("--line-spacing")
    parser.add_argument("--space-before")
    parser.add_argument("--space-after")
    parser.add_argument("--has-image", action=argparse.BooleanOptionalAction)
    parser.add_argument("--suspected-heading-level1", action=argparse.BooleanOptionalAction)
    parser.add_argument("--include-runs", action=argparse.BooleanOptionalAction, default=True)
    parser.add_argument("--limit", type=int, default=100)
    parser.add_argument("--report", type=Path)


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


def _cmd_inspect(args: argparse.Namespace) -> dict[str, Any]:
    return inspect_hdocx(args.bundle, args.target_id, kind=args.kind)


def _cmd_query(args: argparse.Namespace) -> dict[str, Any]:
    return query_hdocx(args.bundle, **_query_kwargs(args))


def _cmd_find(args: argparse.Namespace) -> dict[str, Any]:
    return find_hdocx(args.bundle, **_query_kwargs(args))


def _cmd_assert(args: argparse.Namespace) -> dict[str, Any]:
    assertions: list[Any] = []
    if args.assertions_json:
        loaded = json.loads(args.assertions_json)
        if not isinstance(loaded, list):
            raise HDocxError("ASSERTIONS_JSON_INVALID", "--assertions-json must be a JSON array.", {})
        assertions.extend(loaded)
    if args.assertions:
        assertions.extend(args.assertions)
    return assert_hdocx(args.bundle, assertions or None)


def _cmd_doctor(args: argparse.Namespace) -> dict[str, Any]:
    return doctor_report()


def _cmd_mcp(args: argparse.Namespace) -> dict[str, Any]:
    from .mcp_server import main as mcp_main

    raise SystemExit(mcp_main())


def _query_kwargs(args: argparse.Namespace) -> dict[str, Any]:
    return {
        "kind": args.kind,
        "text": args.text,
        "text_regex": args.text_regex,
        "style_id": args.style_id,
        "font_family": args.font_family,
        "east_asia_font": args.east_asia_font,
        "font_size": args.font_size,
        "bold": args.bold,
        "italic": args.italic,
        "color": args.color,
        "align": args.align,
        "line_spacing": args.line_spacing,
        "space_before": args.space_before,
        "space_after": args.space_after,
        "has_image": args.has_image,
        "suspected_heading_level1": args.suspected_heading_level1,
        "include_runs": args.include_runs,
        "limit": args.limit,
    }


def _print_json(payload: dict[str, Any]) -> None:
    text = json.dumps(payload, ensure_ascii=True, indent=2, sort_keys=True) + "\n"
    stdout_buffer = getattr(sys.stdout, "buffer", None)
    if stdout_buffer is not None:
        stdout_buffer.write(text.encode("utf-8"))
        stdout_buffer.flush()
        return
    sys.stdout.write(text)
