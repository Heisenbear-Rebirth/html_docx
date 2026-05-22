from __future__ import annotations

from pathlib import Path
from typing import Any
import shutil
import subprocess

from .errors import HDocxError
from .utils import ensure_new_dir, sha256_file


def render_capabilities() -> dict[str, Any]:
    executable = _find_renderer()
    return {
        "available": executable is not None,
        "engine": "libreoffice" if executable else None,
        "executable": executable,
        "requiresExternalRuntime": True,
    }


def render_check_docx(
    input_docx: Path,
    output_dir: Path,
    *,
    force: bool = False,
    allow_missing: bool = False,
    timeout_seconds: int = 120,
) -> dict[str, Any]:
    input_docx = input_docx.resolve()
    output_dir = output_dir.resolve()
    if not input_docx.exists():
        raise HDocxError("RENDER_INPUT_NOT_FOUND", "Input DOCX does not exist.", {"path": str(input_docx)})
    if input_docx.suffix.lower() != ".docx":
        raise HDocxError("RENDER_INPUT_NOT_DOCX", "Render input must be a .docx file.", {"path": str(input_docx)})

    ensure_new_dir(output_dir, force=force)
    capabilities = render_capabilities()
    if not capabilities["available"]:
        return {
            "ok": bool(allow_missing),
            "command": "render-check",
            "input": str(input_docx),
            "output": str(output_dir),
            "available": False,
            "status": "renderer-missing",
            "message": "LibreOffice/soffice was not found on PATH. No render output was produced.",
            "capabilities": capabilities,
        }

    profile_dir = output_dir / "lo-profile"
    profile_dir.mkdir(parents=True, exist_ok=True)
    command = [
        capabilities["executable"],
        "--headless",
        "--norestore",
        "--nodefault",
        "--nofirststartwizard",
        f"-env:UserInstallation={profile_dir.as_uri()}",
        "--convert-to",
        "pdf",
        "--outdir",
        str(output_dir),
        str(input_docx),
    ]
    try:
        completed = subprocess.run(
            command,
            cwd=str(output_dir),
            capture_output=True,
            text=True,
            timeout=timeout_seconds,
            check=False,
        )
    except subprocess.TimeoutExpired as exc:
        return {
            "ok": False,
            "command": "render-check",
            "input": str(input_docx),
            "output": str(output_dir),
            "available": True,
            "status": "timeout",
            "timeoutSeconds": timeout_seconds,
            "message": str(exc),
            "capabilities": capabilities,
        }

    pdf_path = output_dir / f"{input_docx.stem}.pdf"
    ok = completed.returncode == 0 and pdf_path.exists()
    report: dict[str, Any] = {
        "ok": ok,
        "command": "render-check",
        "input": str(input_docx),
        "output": str(output_dir),
        "available": True,
        "status": "rendered" if ok else "render-failed",
        "returnCode": completed.returncode,
        "stdout": completed.stdout[-2000:],
        "stderr": completed.stderr[-2000:],
        "capabilities": capabilities,
    }
    if pdf_path.exists():
        report["pdf"] = {
            "path": str(pdf_path),
            "sha256": sha256_file(pdf_path),
            "size": pdf_path.stat().st_size,
        }
    return report


def _find_renderer() -> str | None:
    for name in ("soffice", "libreoffice"):
        path = shutil.which(name)
        if path:
            return path
    return None
