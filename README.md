# html_docx

Agent-friendly reversible DOCX editing through H-DOCX bundles.

The project goal is a strict reversible workflow:

```text
DOCX <-> H-DOCX bundle <-> DOCX
```

`document.html` is the agent-editable projection. `agent.edits.hcss` is a small H-DOCX edit DSL for batch formatting and controlled structural operations. The original DOCX package, OOXML parts, and `manifest.json` remain the source of truth.

Core rule:

```text
Editable content must patch exactly.
Non-editable content must be preserved exactly.
Unsafe edits must fail.
```

For the full usage and design explanation, read
[`USAGE_AND_PRINCIPLES.md`](USAGE_AND_PRINCIPLES.md).

## Current Capabilities

- Byte-identical unmodified roundtrip.
- H-DOCX export with original package, extracted parts, HTML projection, H-CSS file, and manifest.
- Editable run text, run formatting, run split, paragraph formatting, table-cell text, headers, footnotes.
- Protected fields, references, comments, revisions, equations, and complex structures.
- Image alt text, image size, existing media replacement, and controlled new image insertion.
- Numbering/list metadata projection, controlled numbering-level edits, new single-level or multi-level list creation, paragraph list assignment, and automatic creation of missing `word/numbering.xml`.
- Existing style definition edits, new paragraph style creation, safe unused-style deletion, paragraph style assignment, style manifest indexing, and automatic creation of missing `word/styles.xml`.
- Simple table row/column insertion and deletion for non-merged tables.
- Controlled comment text edits and revision accept/reject actions.
- Controlled whole-equation OMML replacement from bundle-local source files.
- Package diff, semantic node diff, and fragment-level byte change ranges linked to H-DOCX node ids.
- Single-file and directory-level acceptance checks.
- Built-in synthetic pressure fixture generation.
- DOCX advanced-object audit for customXml, charts, SmartArt, OLE, AlternateContent, fields, equations, revisions, comments, notes, headers/footers, and images.
- Optional render check through LibreOffice/soffice when available.
- Classified inspection for nodes, styles, lists, tables, and images.
- Fragment-preserving patching for simple text and run-format edits.

## CLI

Run from the repository root:

```powershell
$env:PYTHONPATH = "src"
python -m html_docx export input.docx --out work.hdocx
python -m html_docx validate work.hdocx
python -m html_docx plan work.hdocx --report plan.json
python -m html_docx apply work.hdocx --out output.docx --report apply.json
python -m html_docx diff input.docx output.docx --report diff.json
python -m html_docx audit input.docx --report audit.json
python -m html_docx inspect work.hdocx --kind node --id p-000001
python -m html_docx inspect work.hdocx --kind style --id Normal
python -m html_docx roundtrip input.docx --work work.hdocx --out roundtrip.docx --force
python -m html_docx check input.docx --work check.hdocx --out checked.docx --force --report check.json
python -m html_docx batch-check fixtures --work pressure-work --out pressure-out --force --report pressure.json
python -m html_docx generate-fixtures --out pressure-fixtures --force --report pressure-fixtures.json
python -m html_docx render-check input.docx --out render-out --force --allow-missing --report render.json
python -m html_docx doctor
```

## H-CSS Examples

Batch paragraph formatting:

```css
@hdocx-set body {
  select: [data-hdocx-type="paragraph"];
}

@hdocx-edit mode(paragraph-formatting);

body {
  hdocx-align: center;
  hdocx-line-spacing: 1.5;
}
```

Function selectors are available for common agent workflows:

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

Insert a new image after a paragraph:

```css
@hdocx-insert-image after(#p-000001) {
  source: assets/figure.png;
  alt: "Figure 1";
  width-emu: 914400;
  height-emu: 457200;
}
```

The target paragraph may be in the main document or another projected Word XML part such as a header, footer, footnote, or endnote.

Insert a table row:

```css
@hdocx-insert-table-row after(#tr-000001) {
  cells: "New A|New B";
}
```

Create and apply a paragraph style:

```css
@hdocx-style AgentBody {
  type: paragraph;
  name: "Agent Body";
  based-on: Normal;
  hdocx-font-size: 13pt;
}

@hdocx-edit mode(paragraph-style);

#p-000001 {
  hdocx-style-id: AgentBody;
}
```

Delete an unused style:

```css
@hdocx-delete-style(Heading1);
```

Patch a numbering level:

```css
@hdocx-set list-items {
  select: [data-hdocx-num-id="1"][data-hdocx-ilvl="0"];
}

@hdocx-edit mode(numbering-definition);

list-items {
  hdocx-num-format: upperLetter;
  hdocx-level-text: "Appendix %1)";
  hdocx-start: 3;
}
```

Create a two-level list:

```css
@hdocx-list AgentMulti {
  hdocx-num-format: decimal;
  hdocx-level-text: "%1.";
  hdocx-level-1-num-format: lowerLetter;
  hdocx-level-1-level-text: "%2)";
}
```

Replace a protected equation with an OMML fragment stored inside the bundle:

```css
@hdocx-edit mode(equation-omml);

[data-hdocx-protected-kind="equation"] {
  hdocx-omml-source: equations/replacement.omml;
}
```

## Tests

```powershell
python -m unittest discover -s tests
```

Current local status is tracked in `IMPLEMENTATION_STATUS.md`.

## Design Documents

- `FUNCTIONAL_SPEC.md`
- `HDOCX_HTML_DESIGN.md`
- `EDITING_EDGE_CASES.md`
- `SELECTOR_AND_REUSE_DESIGN.md`
- `SELECTOR_EDGE_CASES_AND_GUARDS.md`
- `EDGE_CASE_TEST_MATRIX.md`
- `SOFTWARE_ARCHITECTURE.md`
- `GLOBAL_DELIVERY_PLAN.md`
- `COMPLETION_PLAN.md`
- `IMPLEMENTATION_STATUS.md`
- `AGENT_GUIDE.md`
- `RELEASE_CHECKLIST.md`
- `PRESSURE_FIXTURES.md`
