# html_docx Implementation Status

Updated: 2026-05-22

## Implemented

### Package And Bundle

- Reads `.docx` ZIP/OPC packages.
- Stores the original DOCX as `original/original.docx`.
- Exports an H-DOCX bundle with:
  - `manifest.json`
  - `document.html`
  - `styles.generated.css`
  - `agent.edits.hcss`
  - `original/original.docx`
  - `parts/`
- Unmodified `apply` copies original DOCX bytes directly.
- Unmodified `roundtrip` is byte-identical.

### HTML Projection

- Projects `/word/document.xml`.
- Projects `/word/header*.xml`.
- Projects `/word/footer*.xml`.
- Projects `/word/footnotes.xml`.
- Projects `/word/endnotes.xml`.
- Projects `/word/comments.xml` as read-only comment articles.
- Projects basic tables as `<table>/<tr>/<td>`.
- Projects paragraphs as `<p>` and runs as `<span>`.
- Detects simple editable runs.
- Marks field, object, footnote reference, endnote reference, comment reference, and similar run-level structures as protected.
- Projects comment ranges, bookmarks, revision wrappers, content controls, equations, and related paragraph-level structures as protected nodes.
- Projects numbered/list paragraphs with `data-hdocx-num-id` and `data-hdocx-ilvl` instead of flattening them into HTML lists.
- Parses `word/numbering.xml` into the manifest and exposes resolved list metadata such as abstract numbering id, number format, level text, start value, suffix, and indentation.
- Treats paragraph style id and numbering/list metadata as read-only; tampering is rejected even when the same paragraph also contains supported text edits.
- Emits `data-hdocx-part` on projected editable/protected nodes so selectors can target specific Word parts.

### Patch Support

- Plain run text edits.
- Fragment-preserving patching for simple `w:t` text edits when no whitespace-preservation attribute change is needed.
- Fragment-preserving patching for simple run-format edits.
- Run-level formatting:
  - bold
  - italic
  - font-size
  - color
- Manual `run-segment` split for local formatting inside a run.
- Paragraph-level formatting:
  - align
  - line-spacing
  - first-line-indent
- Table cell text edits.
- Simple table row insertion/deletion through H-CSS.
- Simple table column insertion/deletion through H-CSS.
- Header text edits.
- Footnote text edits.
- Fragment-preserving paragraph property patching for supported paragraph formatting in simple XML cases.
- Image alt text edits via `wp:docPr descr`.
- Image size edits via existing DrawingML `wp:extent` EMU values.
- Controlled replacement of existing `word/media/*` package entries from the extracted `parts/` tree.
- Controlled new image insertion through `@hdocx-insert-image after(...)` and `before(...)`, including media entry creation, the target part `.rels`, and `[Content_Types].xml` updates.
- Existing image formatting through `@hdocx-edit mode(image-formatting)`, including alt text, EMU size, and host paragraph spacing/alignment.
- Non-media `parts/` modifications are rejected instead of silently applied.
- Numbering/list definition edits through `@hdocx-edit mode(numbering-definition)`.
- New single-level and multi-level list definition creation through `@hdocx-list`, including controlled creation of `word/numbering.xml` when missing.
- Paragraph list assignment through `@hdocx-edit mode(paragraph-numbering)`.
- Existing paragraph style assignment through `@hdocx-edit mode(paragraph-style)`.
- New paragraph style creation through `@hdocx-style`, including controlled creation of `word/styles.xml` when missing.
- Safe unused-style deletion through `@hdocx-delete-style(...)`; styles used by projected paragraphs are rejected.
- Controlled comment body text edits through `@hdocx-edit mode(comment-text)`.
- Controlled revision accept/reject actions through `@hdocx-edit mode(revision-action)`.
- Controlled whole-equation replacement through `@hdocx-edit mode(equation-omml)` using bundle-local OMML source files.
- Controlled structural blank-line insertion through `@hdocx-edit mode(paragraph-structure)`, with per-empty-paragraph spacing.
- Grouped part writeback: only modified ZIP entries are replaced.

### H-CSS V1

Supports:

- `@hdocx-token`
- `@hdocx-set`
- `@hdocx-format`
- `@hdocx-include`
- `@hdocx-edit mode(paragraph-formatting)`
- `@hdocx-edit mode(all-runs)`
- `@hdocx-edit mode(image-formatting)`
- `@hdocx-edit mode(direct-formatting)`
- `@hdocx-edit mode(style-definition)`
- `@hdocx-edit mode(numbering-definition)`
- `@hdocx-edit mode(paragraph-style)`
- `@hdocx-edit mode(paragraph-numbering)`
- `@hdocx-edit mode(comment-text)`
- `@hdocx-edit mode(revision-action)`
- `@hdocx-edit mode(equation-omml)`
- `@hdocx-edit mode(paragraph-structure)`
- `@hdocx-style StyleId`
- `@hdocx-delete-style(StyleId)`
- `@hdocx-list ListId`
- `@hdocx-insert-image after(selector)`
- `@hdocx-insert-image before(selector)`
- `@hdocx-insert-table-row after(selector)`
- `@hdocx-delete-table-row(selector)`
- `@hdocx-insert-table-column after(selector)`
- `@hdocx-delete-table-column(selector)`

Current H-CSS behavior:

- Batch paragraph formatting.
- Batch run formatting.
- Explicit formatting declaration contract for common academic layout:
  - run: font family, Latin script fonts, East Asian fonts, complex-script fonts, font size, bold, italic, color.
  - paragraph: alignment, first-line indent, exact or multiple line spacing, space before, space after.
  - paragraph structure: manual page breaks and real empty paragraphs with optional style, line spacing, space before, and space after.
- `plan` returns H-CSS diagnostics with selector matches, declaration support, normalized values, OOXML mappings, line numbers, per-rule errors, and generated patch ids.
- Direct `#node-id` selectors.
- `.class` selectors.
- Function selectors:
  - `type(paragraph)`
  - `style(BodyText)`
  - `list(1)` or `list(1, 0)`
  - `part(/word/header1.xml, paragraph)`
- Patching existing Word style definitions in `word/styles.xml`.
- Creating paragraph styles in existing or newly created `word/styles.xml`.
- Deleting unused paragraph styles.
- Applying paragraph styles.
- Patching existing Word numbering levels in `word/numbering.xml`.
- Creating single-level or multi-level list definitions in existing or newly created `word/numbering.xml`.
- Applying paragraph numbering.
- Editing comment body text through a dedicated protected-node mode.
- Accepting/rejecting insert/delete revision wrappers through a dedicated protected-node mode.
- Replacing a protected equation with validated OMML from a bundle-local file.
- Inserting images into projected Word XML parts, including headers/footers/notes where a paragraph target is available.
- Formatting existing projected images and their host paragraphs to avoid inline-image clipping from exact line spacing.
- Inserting idempotent structural empty paragraphs before or after target paragraphs.
- Zero-match selectors fail.
- `allow-empty: true` sets do not fail when they match nothing.
- Protected matches fail.

### CLI

Supported commands:

```powershell
python -m html_docx export input.docx --out work.hdocx
python -m html_docx validate work.hdocx
python -m html_docx plan work.hdocx --report plan.json
python -m html_docx apply work.hdocx --out output.docx --report apply.json
python -m html_docx diff input.docx output.docx --report diff.json
python -m html_docx audit input.docx --report audit.json
python -m html_docx inspect work.hdocx --kind node --id p-000001
python -m html_docx inspect work.hdocx --kind style --id Normal
python -m html_docx query work.hdocx --text "Keywords"
python -m html_docx find work.hdocx --kind image
python -m html_docx assert work.hdocx --assertion text-payload-unchanged
python -m html_docx roundtrip input.docx --work work.hdocx --out roundtrip.docx
python -m html_docx check input.docx --work check.hdocx --out checked.docx --force --report check.json
python -m html_docx batch-check fixtures --work pressure-work --out pressure-out --force --report pressure.json
python -m html_docx generate-fixtures --out pressure-fixtures --force --report pressure-fixtures.json
python -m html_docx doctor
```

Inspection supports `node`, `style`, `list`, `table`, and `image`.
`query` / `find` locate nodes by text, style id, run font properties,
paragraph properties, image presence, and likely level-1 heading heuristics.
`assert` runs assertion-style checks such as unchanged text payload, required
empty paragraphs, and image host paragraphs that must not use exact line
spacing. Assertion objects can set `afterApply` / `plannedOutput` to validate a
bundle-local post-apply re-export, and level-1 heading assertions support
`includeRegex`, `excludeRegex`, and default front-matter exclusions.
`check` runs export/apply/diff acceptance for one DOCX.
`batch-check` runs the same acceptance chain over a single DOCX or a directory of DOCX files.
`audit` reports protected/high-risk DOCX structures and the project policy for each detected feature.
`generate-fixtures` creates a local synthetic pressure DOCX suite.

### Report And Diff

- `diff` reports whole-package byte identity separately from DOCX entry-content identity.
- `diff` reports entry counts for changed, unchanged, left-only, right-only, and ZIP-metadata-only changes.
- `diff` includes changed entry details:
  - part kind
  - SHA-256
  - compressed and uncompressed size
  - CRC
  - compression type
  - ZIP order
  - ZIP timestamp
- `diff` includes `semanticDiff`, an in-memory H-DOCX projection comparison with node counts, changed node ids, added/removed node ids, and field-level node changes.
- `semanticDiff` distinguishes document semantic changes from media-only binary replacement.
- `diff` includes `manualPageBreakDiff` and `emptyParagraphDiff` for explicit structural insertions, while semantic alignment prevents following text from being misreported as changed.
- `diff` includes `fragmentDiff`, with byte-range change windows per changed/added/removed package entry and linked H-DOCX node ids when the entry is projected.
- `apply` includes a patch summary grouped by entry and operation.
- `apply` patch summaries include risk classes such as `fragment-preserving-eligible`, `xml-entry-reserialize`, `package-metadata`, `structural-insert`, and `binary-package-entry`.
- `apply` includes a package diff between the original DOCX and the produced DOCX.
- `audit` classifies customXml, chart, SmartArt, OLE, AlternateContent, VML/text boxes, fields, equations, revisions, comments, footnotes, endnotes, headers/footers, and images.

### Tests

Current automated tests:

```text
101 tests passing
```

Coverage includes:

- Byte-identical unmodified roundtrip.
- Simple text patch preserving non-target XML bytes in the modified part.
- Export and validate.
- Plain run text patch.
- Run formatting patch.
- Run split.
- Paragraph formatting patch.
- H-CSS paragraph formatting.
- H-CSS structural empty paragraph insertion with per-empty-paragraph spacing and idempotence.
- Structured `query` / `find` for text, font/size/alignment, likely level-1 headings, images, and host paragraphs.
- Assertion checks for unchanged text payload, required empty paragraphs, and image host exact line-spacing failures.
- Assertion `afterApply` planned-output checks with original-to-planned paragraph id mapping.
- Level-1 heading assertion front-matter excludes plus include/exclude regex filters.
- H-CSS token/format/include/all-runs.
- H-CSS paper-format declaration contract and OOXML mapping.
- H-CSS unsupported declaration diagnostics with line numbers.
- H-CSS run-level class+attribute selectors, UTF-8 BOM handling, and set aliases.
- H-CSS comma selector lists in rules and `@hdocx-set` aliases.
- H-CSS manual page-break insertion before target paragraphs with semantic diff alignment.
- MCP/CLI structured path-encoding errors for unsafe output paths.
- H-CSS `#node-id`, `.class`, and `allow-empty` selector behavior.
- H-CSS function selector behavior for type, style, list, and part targeting.
- H-CSS style-definition patch.
- H-CSS paragraph-style patch.
- H-CSS style creation and immediate application.
- H-CSS style/list creation when `word/styles.xml` and `word/numbering.xml` are missing.
- H-CSS unused-style deletion and used-style deletion rejection.
- H-CSS numbering-definition patch.
- H-CSS list creation and paragraph numbering assignment.
- H-CSS multi-level list creation and second-level assignment.
- H-CSS zero-match failure.
- Chinese and mixed special-character DOCX paths for audit/export/apply/diff.
- MCP busy response instead of transport-breaking concurrent tool execution.
- Protected run text/format failure.
- Protected node text modification failure.
- ID tamper failure.
- Table cell text patch.
- Simple table row insertion/deletion.
- Simple table column insertion/deletion.
- Header text patch.
- Footnote text patch.
- Comment range and comment body protection.
- Revision wrapper protection.
- Numbered paragraph metadata projection.
- Numbering definition manifest analysis and read-only metadata tamper rejection.
- Image alt text patch.
- Image size patch and invalid size-removal rejection.
- Existing media part replacement and non-media part tamper rejection.
- New image insertion and unsafe image source rejection.
- New image insertion before a paragraph.
- New image insertion into a header part with a new header relationship part.
- Controlled comment text patch.
- Controlled revision accept action.
- Controlled equation OMML replacement and invalid OMML rejection.
- Complex academic synthetic fixture byte-identical roundtrip.
- CLI report file output.
- CLI classified inspect output.
- CLI doctor, check, and batch-check output.
- CLI generate-fixtures output.
- Structured DOCX diff report output.
- Semantic node diff report output.
- Fragment-level diff report output.
- Advanced-object audit report output.

## Current Limits

- Simple text, run-format, drawing metadata, and paragraph-format XML patches can preserve non-target XML bytes in simple cases; split-run, style, numbering, table, equation, and some drawing fallbacks still use ElementTree serialization.
- Diff reports include package/entry-level, projection node-level, and fragment byte-range changes. Fragment reports identify the changed byte window for an entry and link projected XML changes to node ids; they do not expose full changed XML payloads.
- `styles.xml` is indexed into the manifest; existing style IDs can be patched, new paragraph styles can be created, and unused styles can be deleted. Deleting in-use or default styles is rejected.
- Numbering/list support projects and resolves core numbering definitions; existing abstract numbering levels can be edited and new single-level or multi-level lists can be created. Bulk operations are available through selector-based H-CSS, including `list(numId, ilvl)`.
- Simple table row/column insertion and deletion are implemented; merged or irregular table structure editing is still rejected.
- Image support covers alt text, existing `wp:extent` size metadata, existing media entry replacement, and before/after new image insertion into projected Word XML parts.
- Comments, insert/delete revision wrappers, and equations have limited controlled edit actions; charts, SmartArt, and OLE are audited and primarily protected/preserved, not semantically editable.
- Built-in synthetic pressure fixtures can be generated locally; real-world private DOCX corpora still need to be supplied by the user when broader validation is required.
- H-CSS parser is a small V1 parser, not a full CSS parser.
- Large real-world DOCX pressure fixtures are not bundled yet; `batch-check` and `PRESSURE_FIXTURES.md` are available for local fixture directories.

## Next Priorities

1. Real academic DOCX fixture pressure tests with local private fixture directories.
2. Broader semantic editing for selected advanced objects only where strict preservation can be proven.
3. Broader assertion coverage for paper-specific layout invariants.
