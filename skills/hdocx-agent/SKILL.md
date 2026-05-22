---
name: hdocx-agent
description: Use when an agent needs to inspect, edit, round-trip, validate, or pressure-test DOCX files through the strict H-DOCX/html_docx reversible DOCX <-> HTML bundle, especially for preserving Word layout while making controlled text or formatting edits.
---

# H-DOCX Agent

Use this skill for DOCX editing tasks where layout preservation and strict
reversibility matter. H-DOCX is not a best-effort HTML converter. Treat it as a
controlled editing bundle:

```text
DOCX <-> H-DOCX bundle <-> DOCX
```

Unedited round-trips must be byte-identical. Edited round-trips must preserve
untouched OOXML exactly and report every controlled change.

## Boundary

Keep all work inside the current workspace unless the user explicitly approves
otherwise. Do not install packages globally. Put temp files, reports, caches,
render output, and generated bundles inside the workspace. If dependencies are
missing, use a workspace-local virtual environment.

## Pick the Command Form

Prefer the installed CLI if available:

```powershell
html-docx doctor
```

When working inside the `html_docx` source repository, use:

```powershell
$env:PYTHONPATH = "src"
python -m html_docx doctor
```

If neither works, ask the user where the H-DOCX tool/repository is. Do not
guess by installing or searching outside the approved workspace.

## Required Workflow

1. Audit the input before editing unfamiliar DOCX files.
2. Export to a `.hdocx` bundle.
3. Inspect target nodes/styles/lists/tables/images before broad edits.
4. Edit only `document.html`, `agent.edits.hcss`, or supported bundle-local
   assets.
5. Run `plan`, then `apply`, then `diff`.
6. For conversion logic changes, run unit tests and pressure fixtures.
7. For layout-sensitive edited output, run `render-check` when a renderer is
   available.

See `references/workflow.md` for exact commands.

## Allowed Edits

Agents may edit:

- Editable run text in `document.html`.
- Supported H-CSS formatting rules in `agent.edits.hcss`.
- Bundle-local assets referenced by supported H-CSS operations.
- Existing files under `parts/word/media/` only for controlled media
  replacement.

Agents must not edit:

- `manifest.json`.
- `original/original.docx`.
- Protected placeholders in `document.html`.
- Read-only metadata such as `data-hdocx-id`, `data-hdocx-part`, style ids,
  numbering ids, or protected-kind attributes.
- Non-media files under `parts/` unless a dedicated operation explicitly
  supports that change.

## H-CSS Targeting

Prefer inspected ids, named sets, and narrow selectors. Use convenience
functions for common targets:

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

If a selector may match nothing, declare it:

```css
@hdocx-set optional-notes {
  select: .maybe-note;
  allow-empty: true;
}
```

See `references/hcss.md` for compact examples.

## Acceptance

Never claim success from visual inspection alone. Use reports.

- For no-edit round-trips, require `byteIdentical: true`.
- For edited output, require the diff report to show only intended changes.
- SHA256 identity is stronger than render equality for unedited round-trips.
- If an edit cannot be proven safe, fail with a report instead of guessing.

See `references/acceptance.md` for gates and interpretation.
