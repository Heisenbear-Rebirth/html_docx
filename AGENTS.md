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
- Do not modify user `PATH`, MCP client config, shell profiles, or other
  user-level integration files unless the user explicitly asks for that
  outside-repository setup.

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
python -m html_docx create --out new.docx --title "Draft Title" --paragraph "First paragraph." --export-to new.hdocx --force
python -m html_docx audit input.docx --report audit.json
python -m html_docx export input.docx --out work.hdocx --force
python -m html_docx query work.hdocx --text "Keywords"
python -m html_docx find work.hdocx --kind image
python -m html_docx inspect work.hdocx --kind node --id p-000001
python -m html_docx plan work.hdocx --report plan.json
python -m html_docx apply work.hdocx --out output.docx --report apply.json
python -m html_docx diff input.docx output.docx --report diff.json
python -m html_docx assert work.hdocx --assertion text-payload-unchanged
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

## What Agents May Edit

Agents may edit:

- Editable run text in `document.html`.
- Supported H-CSS rules in `agent.edits.hcss`.
- Bundle-local assets referenced by supported H-CSS operations.
- Existing files under `parts/word/media/` only for controlled media
  replacement.

Agents should use `query` / `find` / `hdocx_query` / `hdocx_find` for target
discovery before reading `document.html` directly. These tools return
structured JSON for text, font, font size, paragraph formatting, image runs and
host paragraphs, and likely level-1 headings.

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

Supported selectors are ids, classes, exact attributes, class+attribute
compounds such as `.hdocx-r[data-hdocx-id="r-000001"]`, and the functions above.
Selector lists separated by commas are supported both in rules and inside
`@hdocx-set`.

Keep agent-defined groups in `agent.edits.hcss`; do not add custom classes or
other projection metadata to `document.html`. Use selector lists or named sets
instead:

```css
@hdocx-set body {
  select: id(p-000007), id(p-000008), id(p-000009);
}
```

If a selector may legitimately match nothing, declare it explicitly:

```css
@hdocx-set optional-notes {
  select: .maybe-note;
  allow-empty: true;
}
```

## Formatting Declaration Contract

H-CSS is not browser CSS. Formatting declarations must use the `hdocx-`
prefix. `hdocx_plan` reports selector matches, per-declaration support,
normalized values, OOXML mappings, line numbers, errors, and patch ids.

Use `@hdocx-edit mode(paragraph-formatting);` for paragraph declarations:
`hdocx-text-align`/`hdocx-align`, `hdocx-first-line-indent`,
`hdocx-line-spacing`, `hdocx-line-spacing-exact`, `hdocx-space-before`, and
`hdocx-space-after`. Use `hdocx-manual-page-break-before: true` for explicit
manual page breaks before target paragraphs; `diff` reports them under
`manualPageBreakDiff`.

Use `@hdocx-edit mode(paragraph-structure);` when a blank line must be a real
Word empty paragraph, not just visual spacing. Supported declarations are
`hdocx-insert-empty-paragraph-before`, `hdocx-insert-empty-paragraph-after`,
`hdocx-empty-paragraph-style-id`, `hdocx-empty-paragraph-line-spacing`,
`hdocx-empty-paragraph-line-spacing-exact`, `hdocx-empty-paragraph-space-before`,
and `hdocx-empty-paragraph-space-after`. `plan` reports
`insert-empty-paragraph-before` / `insert-empty-paragraph-after`; `apply` is
idempotent for an already-adjacent blank paragraph; `diff` reports changes under
`emptyParagraphDiff`.

Use `@hdocx-edit mode(all-runs);` for run declarations:
`hdocx-font-family`, `hdocx-eastAsia-font`/`hdocx-east-asia-font`,
`hdocx-ascii-font`, `hdocx-hansi-font`, `hdocx-cs-font`, `hdocx-font-size`,
`hdocx-bold`, `hdocx-italic`, and `hdocx-color`.

Use `@hdocx-edit mode(image-formatting);` for existing projected images.
Supported declarations are `hdocx-alt`, `hdocx-width-emu`, `hdocx-height-emu`,
and host paragraph declarations prefixed with `hdocx-paragraph-`:
`hdocx-paragraph-line-spacing`, `hdocx-paragraph-line-spacing-exact`,
`hdocx-paragraph-space-before`, `hdocx-paragraph-space-after`, and
`hdocx-paragraph-text-align`/`hdocx-paragraph-align`. If an inline picture is
cropped by inherited fixed body line spacing, set its host paragraph spacing
through this mode instead of changing unrelated body paragraphs.

## Acceptance Gates

Before claiming success, run the relevant checks:

- `python -m unittest discover -s tests`
- `python -m html_docx check ...` for the target DOCX
- `python -m html_docx diff ...` after edited output
- `python -m html_docx assert ...` for task-specific invariants such as
  unchanged text payload, required structural blank paragraphs, and image host
  paragraphs that must not use exact line spacing
  - For structure/format assertions that should inspect the planned result,
    set `afterApply: true` or `plannedOutput: true` in the assertion object.
  - `level1-headings-have-empty-paragraph-before` supports `includeRegex`,
    `excludeRegex`, and default front-matter excludes for abstract, contents,
    and keywords labels.
- `python -m html_docx batch-check ...` for pressure fixtures when changing
  conversion logic
Passing SHA256 byte identity is stronger than render equality for unedited
round-trips.

## Reference Documents

- `AGENT_GUIDE.md`: detailed agent workflow and editing examples.
- `FUNCTIONAL_SPEC.md`: product-level functional boundary.
- `HDOCX_HTML_DESIGN.md`: HTML/H-DOCX representation design.
- `SELECTOR_AND_REUSE_DESIGN.md`: H-CSS selector and reuse model.
- `SELECTOR_EDGE_CASES_AND_GUARDS.md`: selector edge cases and safeguards.
- `EDITING_EDGE_CASES.md`: editing edge cases.
- `EDGE_CASE_TEST_MATRIX.md`: edge and pressure test matrix.
- `PRESSURE_FIXTURES.md`: pressure fixture coverage.
- `RELEASE_CHECKLIST.md`: release verification checklist.
