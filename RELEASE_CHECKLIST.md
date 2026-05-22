# Release Checklist

Use this checklist before calling the project deliverable.

## Local Boundary

- All generated files are inside the repository directory.
- No global Python packages or system configuration were changed.
- Any future third-party dependency must be installed into a project-local virtual environment.

## Automated Tests

```powershell
python -m unittest discover -s tests
```

Expected current result:

```text
66 tests passing
```

## CLI Smoke Tests

```powershell
$env:PYTHONPATH = "src"
python -m html_docx doctor
python -m html_docx --help
python -m html_docx audit sample.docx --report sample-audit.json
python -m html_docx render-check sample.docx --out sample-render --force --allow-missing --report sample-render.json
```

For a sample DOCX:

```powershell
python -m html_docx check sample.docx --work sample-check.hdocx --out sample-checked.docx --force --report sample-check.json
```

For a local fixture directory:

```powershell
python -m html_docx generate-fixtures --out pressure-fixtures --force --report pressure-fixtures.json
python -m html_docx batch-check fixtures --work pressure-work --out pressure-out --force --report pressure.json
```

## Functional Acceptance

- Unmodified DOCX roundtrip is byte-identical.
- Plain editable text edits patch only intended nodes.
- Run formatting and paragraph formatting produce explicit patch reports.
- Style and numbering edits are represented as dedicated H-CSS modes.
- Missing style and numbering parts can be created through controlled H-CSS operations.
- Unused style deletion is allowed; used/default style deletion is rejected.
- Tables support simple row/column operations and reject merged or unsafe structures.
- Headers, footnotes, comments, revisions, images, and equations have explicit support or explicit protection.
- Directory pressure checks can be run through `batch-check` using `PRESSURE_FIXTURES.md`.
- Built-in pressure fixtures can be generated through `generate-fixtures`.
- Optional render QA records renderer availability and PDF output when LibreOffice/soffice exists.
- Unknown or unsupported structures are preserved or rejected, never silently simplified.

## Report Acceptance

- `plan` explains intended patches.
- `apply` reports patch count, operation summary, and package diff.
- `apply` reports patch risk classes.
- `diff` reports byte identity, entry identity, entry details, semantic node diff, and fragment byte ranges linked to node ids.
- `audit` reports high-risk/protected DOCX structures and policies.
- `check` and `batch-check` expose pass/fail acceptance fields.

## Documentation Acceptance

- `README.md` lists current commands and core examples.
- `AGENT_GUIDE.md` explains safe agent workflow.
- `IMPLEMENTATION_STATUS.md` reflects the current test count and limits.
- `PRESSURE_FIXTURES.md` explains local real-DOCX fixture validation.
- Design documents remain available for architecture and edge-case reasoning.
