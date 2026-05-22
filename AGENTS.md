# Agent Rules for H-DOCX

This repository is designed for Codex, Claude Code, and similar agents to edit
DOCX files through a strict reversible HTML-like bundle.

## Non-Negotiable Boundary

All operations must stay inside the current repository directory.

- Do not write to parent directories, user profile folders, global caches, or
  system locations.
- Do not install packages globally.
- If Python dependencies are needed, create and use the project-local `.venv`.
- Keep temporary files, build caches, render outputs, and generated reports
  inside this repository.
- If a task truly requires touching anything outside this repository, stop and
  ask the user first.
- Do not run `scripts/install-hdocx.ps1` unless the user explicitly asks for
  user-level installation. That script is intentionally for configuring
  user-level CLI and agent skills outside this repository.

## Core Contract

H-DOCX is not a best-effort HTML converter.

The contract is:

```text
DOCX <-> H-DOCX bundle <-> DOCX
```

Unedited round-trips must be byte-identical. Edited round-trips must preserve
all untouched OOXML exactly and must report every controlled change.

If an edit cannot be proven safe, fail with a report instead of guessing.

## Standard Agent Workflow

Use the local source tree while developing:

```powershell
$env:PYTHONPATH = "src"
python -m html_docx audit input.docx --report audit.json
python -m html_docx export input.docx --out work.hdocx --force
python -m html_docx inspect work.hdocx --kind node --id p-000001
python -m html_docx plan work.hdocx --report plan.json
python -m html_docx apply work.hdocx --out output.docx --report apply.json
python -m html_docx diff input.docx output.docx --report diff.json
```

For a no-edit reversibility proof:

```powershell
python -m html_docx check input.docx --work check.hdocx --out checked.docx --force --report check.json
```

For pressure testing:

```powershell
python -m html_docx generate-fixtures --out pressure-fixtures --force --report pressure-fixtures.json
python -m html_docx batch-check pressure-fixtures --work pressure-work --out pressure-out --force --report pressure.json
```

For optional visual QA when LibreOffice/soffice is available:

```powershell
python -m html_docx render-check output.docx --out render-out --force --allow-missing --report render.json
```

## What Agents May Edit

Agents may edit:

- Editable run text in `document.html`.
- Supported H-CSS rules in `agent.edits.hcss`.
- Bundle-local assets referenced by supported H-CSS operations.
- Existing files under `parts/word/media/` only for controlled media
  replacement.

Agents must not edit:

- `manifest.json`.
- `original/original.docx`.
- Protected placeholders in `document.html`.
- Read-only metadata attributes such as `data-hdocx-id`,
  `data-hdocx-part`, style ids, numbering ids, or protected-kind attributes.
- Non-media files under `parts/` unless a dedicated operation explicitly
  supports that change.

## Targeting Rules

Prefer inspected ids or named sets over broad selectors.

Supported convenience selectors include:

```css
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

If a selector may legitimately match nothing, declare it explicitly:

```css
@hdocx-set optional-notes {
  select: .maybe-note;
  allow-empty: true;
}
```

## Acceptance Gates

Before claiming success, run the relevant checks:

- `python -m unittest discover -s tests`
- `python -m html_docx check ...` for the target DOCX
- `python -m html_docx diff ...` after edited output
- `python -m html_docx batch-check ...` for pressure fixtures when changing
  conversion logic
- `python -m html_docx render-check ...` when visual QA is available or when
  edited layout needs visual confirmation

Passing SHA256 byte identity is stronger than render equality for unedited
round-trips.

## Reference Documents

- `skills/hdocx-agent/SKILL.md`: installable/local Codex skill for H-DOCX
  agent workflows.
- `AGENT_GUIDE.md`: detailed agent workflow and editing examples.
- `FUNCTIONAL_SPEC.md`: product-level functional boundary.
- `HDOCX_HTML_DESIGN.md`: HTML/H-DOCX representation design.
- `SELECTOR_AND_REUSE_DESIGN.md`: H-CSS selector and reuse model.
- `SELECTOR_EDGE_CASES_AND_GUARDS.md`: selector edge cases and safeguards.
- `EDITING_EDGE_CASES.md`: editing edge cases.
- `EDGE_CASE_TEST_MATRIX.md`: edge and pressure test matrix.
- `PRESSURE_FIXTURES.md`: pressure fixture coverage.
- `RELEASE_CHECKLIST.md`: release verification checklist.
