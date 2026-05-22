# H-DOCX Agent Guide

This guide is for Codex, Claude Code, and other code agents editing DOCX files through this project.

## Operating Principle

H-DOCX is not a best-effort HTML converter. It is a strict reversible editing bundle:

```text
DOCX <-> H-DOCX bundle <-> DOCX
```

Editable content must patch exactly. Protected content must be preserved. If an edit cannot be proven safe, the tool must fail.

## Standard Workflow

```powershell
$env:PYTHONPATH = "src"
python -m html_docx export input.docx --out work.hdocx --force
python -m html_docx audit input.docx --report audit.json
python -m html_docx inspect work.hdocx --kind node --id p-000001
python -m html_docx plan work.hdocx --report plan.json
python -m html_docx apply work.hdocx --out output.docx --report apply.json
python -m html_docx diff input.docx output.docx --report diff.json
python -m html_docx render-check output.docx --out render-out --force --allow-missing --report render.json
```

For an unedited reversibility check:

```powershell
python -m html_docx check input.docx --work check.hdocx --out checked.docx --force --report check.json
```

For a local fixture directory:

```powershell
python -m html_docx generate-fixtures --out pressure-fixtures --force --report pressure-fixtures.json
python -m html_docx batch-check fixtures --work pressure-work --out pressure-out --force --report pressure.json
```

## What Agents May Edit

Agents may edit:

- Plain editable run text in `document.html`.
- Supported H-CSS formatting rules in `agent.edits.hcss`.
- Bundle-local source assets referenced by supported H-CSS operations, such as inserted images or OMML replacement files.
- Existing files under `parts/word/media/` for controlled media replacement.

Agents must not edit:

- `manifest.json`.
- `original/original.docx`.
- Protected placeholders in `document.html`.
- Read-only metadata attributes such as `data-hdocx-id`, style ids, numbering ids, or protected-kind attributes.
- Non-media files under `parts/` unless a dedicated operation supports the change.

## Inspection

Use `inspect` before writing broad H-CSS rules:

```powershell
python -m html_docx inspect work.hdocx --kind node --id p-000001
python -m html_docx inspect work.hdocx --kind style --id Normal
python -m html_docx inspect work.hdocx --kind list --id 1
python -m html_docx inspect work.hdocx --kind table --id tbl-000001
python -m html_docx inspect work.hdocx --kind image --id r-000001
```

Use `audit` before editing unfamiliar real-world DOCX files. Treat reported high-risk structures such as customXml, charts, SmartArt, OLE, AlternateContent, fields, equations, revisions, and comments as preserved/protected unless a dedicated H-CSS mode supports the exact requested change.

## H-CSS Safety Pattern

Prefer named sets before broad edits:

```css
@hdocx-set body {
  select: [data-hdocx-type="paragraph"];
}

@hdocx-edit mode(paragraph-formatting);

body {
  hdocx-line-spacing: 1.5;
  hdocx-first-line-indent: 2char;
}
```

Function selectors keep common targeting concise:

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

If a selector may legitimately match nothing, make that explicit:

```css
@hdocx-set optional-notes {
  select: .maybe-note;
  allow-empty: true;
}
```

## Equation Replacement

Equations are protected by default. To intentionally replace a whole equation, place a valid OMML fragment inside the bundle and reference it:

```css
@hdocx-edit mode(equation-omml);

[data-hdocx-protected-kind="equation"] {
  hdocx-omml-source: equations/replacement.omml;
}
```

The OMML source must be a safe relative path inside the bundle, encoded as UTF-8, and its root must be `m:oMath` or `m:oMathPara`.

## Style And List Creation

`@hdocx-style` and `@hdocx-list` can create missing `word/styles.xml` or `word/numbering.xml` parts when needed. `@hdocx-delete-style(StyleId);` is allowed only for unused styles; deleting a style still referenced by projected paragraphs is rejected during `plan`.

## Acceptance Rules

Before delivering an edited DOCX:

1. Run `plan`.
2. Run `apply`.
3. Run `diff`.
4. Confirm changed entries, `semanticDiff` nodes, and `fragmentDiff` byte ranges match the requested edit.
5. For unedited or fixture validation, run `check` or `batch-check`.
6. When a renderer is available, run `render-check`; when it is not available, use `--allow-missing` so the absence is recorded explicitly.

Do not call a document safe merely because a DOCX file was produced. The diff report is part of the deliverable.
