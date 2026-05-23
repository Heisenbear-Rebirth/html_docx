# html_docx

[中文 README](README.zh-CN.md)

Agent-oriented, reversible DOCX editing through H-DOCX bundles.

`html_docx` lets coding agents such as Codex and Claude Code inspect, edit, and
validate Word documents without turning DOCX into lossy browser HTML. It exposes
the parts agents should edit as an HTML-like projection and keeps the original
OOXML package as the source of truth.

```text
DOCX <-> H-DOCX bundle <-> DOCX
```

The core contract is deliberately strict:

- Unedited round-trips must be byte-identical.
- Edited round-trips must preserve every untouched OOXML part exactly.
- Unsupported or unsafe edits must fail with a report instead of guessing.
- Validation reports, not visual inspection alone, are the source of truth.

## Why H-DOCX

Normal DOCX-to-HTML conversion is useful for display, but it is not enough for
agentic editing. Academic and professional Word documents contain styles,
numbering, footnotes, endnotes, headers, fields, equations, comments, revisions,
relationships, media, package metadata, and other OOXML structures that ordinary
HTML cannot represent without loss.

H-DOCX takes a different approach:

- Project editable content into `document.html`.
- Express formatting and structural operations in `agent.edits.hcss`.
- Preserve the original DOCX, raw package parts, relationships, and metadata.
- Patch only controlled fragments back into the original package.
- Diff and audit the result before declaring success.

This makes the format practical for agents: they can read and edit familiar
HTML-like files, while the tool enforces DOCX safety.

## Status

This is an early but usable implementation focused on strict reversibility,
controlled edits, and agent workflows.

Current local validation at the time of this release:

- `101` unit tests passing.
- Built-in pressure fixture round-trip: `6/6` byte-identical.
- Real sample round-trip: byte-identical and semantically identical.
- MCP stdio smoke test passing.

## Requirements

- Windows PowerShell for the bundled installer script.
- Python `>=3.11`.
- No Python runtime dependencies are required by the package itself.

The project is designed to avoid global Python package installation. Use a local
workspace `.venv`, or the user-level isolated installer described below.

## Quick Start

### Clone

```powershell
git clone https://github.com/Heisenbear-Rebirth/html_docx.git
cd html_docx
```

### Verify From Source

```powershell
$env:PYTHONPATH = "src"
python -m html_docx doctor
python -m unittest discover -s tests
python -m html_docx generate-fixtures --out pressure-fixtures --force --report pressure-fixtures.json
python -m html_docx batch-check pressure-fixtures --work pressure-work --out pressure-out --force --report pressure.json
```

These commands do not need network access and do not modify global Python state.

### Put The Commands On PATH

H-DOCX does not edit your MCP client configuration automatically. Install the
package into a Python environment you control, then add that environment's
command directory to `PATH` so every workspace can run:

```text
html-docx
html-docx-mcp
```

For example, after installing into a venv, add its `Scripts` directory on
Windows:

```powershell
$env:PATH = "C:\Tools\hdocx\.venv\Scripts;$env:PATH"
```

Use your own install location. For a persistent setup, add the same command
directory to your user `PATH` through your shell profile or Windows environment
variable settings.

Open a new terminal, then verify:

```powershell
html-docx doctor
'{"jsonrpc":"2.0","id":1,"method":"tools/list","params":{}}' | html-docx-mcp
```

Use this MCP JSON in clients that accept the common `mcpServers` shape:

```json
{
  "mcpServers": {
    "hdocx": {
      "command": "html-docx-mcp",
      "args": []
    }
  }
}
```

If your MCP client uses a different outer key, keep the `hdocx` server object
and adapt only the wrapper. Agents should pass each tool's `root` argument for
the active workspace. `HDOCX_MCP_ROOT` is available only as an optional fallback
when a client cannot pass `root`; when omitted, the server falls back to
`CLAUDE_PROJECT_DIR` and then its current directory.

Restart your MCP client after changing its configuration.

## Editing Workflow

### Create A New DOCX

Use the built-in canonical blank template when the task starts from no existing
document:

```powershell
html-docx create --out new.docx --title "Draft Title" --paragraph "First paragraph." --export-to new.hdocx --force
html-docx check new.docx --work new-check.hdocx --out new-checked.docx --force --report new-check.json
```

`create` writes a valid DOCX package first. If `--export-to` is provided, it
also creates an H-DOCX bundle so agents can continue through
`document.html` and `agent.edits.hcss`.

### No-Edit Reversibility Check

Use this before trusting a new document family:

```powershell
html-docx check input.docx --work check.hdocx --out checked.docx --force --report check.json
```

Success means:

- `ok: true`
- `acceptance.byteIdentical: true`
- `acceptance.semanticIdentical: true`
- input and output SHA256 values match

If SHA256 matches, the DOCX byte stream is identical. That is stronger evidence
than render equality for an unedited round-trip.

### Controlled Edit

```powershell
html-docx audit input.docx --report audit.json
html-docx export input.docx --out work.hdocx --force
html-docx inspect work.hdocx --kind node --id p-000001

# Edit work.hdocx/document.html or work.hdocx/agent.edits.hcss.

html-docx plan work.hdocx --report plan.json
html-docx apply work.hdocx --out output.docx --report apply.json
html-docx diff input.docx output.docx --report diff.json
```

When running directly from source instead of an installed CLI, prefix commands
with:

```powershell
$env:PYTHONPATH = "src"
python -m html_docx ...
```

## H-DOCX Bundle Layout

An exported bundle is a normal directory:

```text
work.hdocx/
  manifest.json
  document.html
  agent.edits.hcss
  styles.generated.css
  audit.log.jsonl
  original/
    original.docx
    entries.json
  parts/
    ...
```

Important rules:

- Edit `document.html` only where content is projected as editable.
- Edit `agent.edits.hcss` for supported formatting and structural operations.
- Do not edit `manifest.json`.
- Do not edit `original/original.docx`.
- Do not edit protected placeholders.
- Do not rewrite arbitrary OOXML under `parts/` unless a dedicated operation
  explicitly supports that change.

## CLI Reference

| Command | Purpose |
| --- | --- |
| `doctor` | Report runtime capabilities. |
| `create` | Create a new canonical DOCX, optionally exporting it to H-DOCX. |
| `audit` | Detect high-risk DOCX structures and preservation policies. |
| `export` | Convert a DOCX into an H-DOCX bundle. |
| `validate` | Validate an H-DOCX bundle before applying it. |
| `inspect` | Inspect nodes, styles, lists, tables, or images by id. |
| `query` / `find` | Return structured matches by text, style, font, paragraph formatting, image presence, or likely level-1 headings. |
| `assert` | Run assertion-style acceptance checks over a bundle. |
| `plan` | Plan edits without writing a DOCX. |
| `apply` | Apply a bundle back to DOCX. |
| `diff` | Compare two DOCX packages with package, semantic, and fragment reports. |
| `roundtrip` | Export and apply without edits. |
| `check` | Run export/apply/diff acceptance for one DOCX. |
| `batch-check` | Run `check` over a file or directory. |
| `generate-fixtures` | Generate synthetic pressure DOCX fixtures. |
| `mcp` | Run the stdio MCP server. |

The package also installs a dedicated `html-docx-mcp` command for MCP clients.

## Structured Query And Assertions

Prefer `query`/`find` over reading `document.html` when locating targets. It
returns JSON from the projection manifest, including host paragraphs for image
runs and effective style-inherited formatting where available.

```powershell
html-docx query work.hdocx --text "Keywords"
html-docx query work.hdocx --align center --font-size 14pt --font-family SimHei --suspected-heading-level1
html-docx find work.hdocx --kind image
```

Use `assert` for post-edit checks that agents otherwise tend to script by hand:

```powershell
html-docx assert work.hdocx --assertion text-payload-unchanged
html-docx assert work.hdocx --assertion images-host-paragraph-not-exact-line-spacing
html-docx assert work.hdocx --assertions-json "[{\"type\":\"paragraphs-have-empty-before\",\"paragraphIds\":[\"p-000006\"]}]"
html-docx assert work.hdocx --assertions-json "[{\"type\":\"paragraphs-have-empty-before\",\"paragraphIds\":[\"p-000006\"],\"afterApply\":true}]"
```

Supported built-in assertions are:

- `text-payload-unchanged`
- `paragraphs-have-empty-before`
- `paragraphs-have-empty-after`
- `images-host-paragraph-not-exact-line-spacing`
- `level1-headings-have-empty-paragraph-before`

By default, structure and formatting assertions inspect the current exported
bundle. Set `afterApply: true` or `plannedOutput: true` on an assertion object
to apply the planned edits to a bundle-local scratch DOCX, export it again, and
check that planned output state. `text-payload-unchanged` normally checks the
planned patch list; with `afterApply: true`, it compares the original and
planned-output text payloads.

`level1-headings-have-empty-paragraph-before` uses a heading heuristic. It
supports `includeRegex` and `excludeRegex`, and by default excludes front-matter
labels such as abstract, contents/table of contents, and keywords. Set
`useDefaultExcludes: false` when those labels are intentional targets.

## H-CSS Examples

H-CSS is a small project-specific edit language. It is not browser CSS. Every
formatting declaration must use the `hdocx-` prefix; normal CSS declarations are
reported as unsupported by `plan`.

`hdocx_plan` is the contract boundary for agents. For each H-CSS rule it reports
the rule line, matched H-DOCX node ids, each declaration's support status,
normalized value, OOXML mapping, and generated patch ids.

### Paragraph Formatting

```css
@hdocx-set body {
  select: [data-hdocx-type="paragraph"];
}

@hdocx-edit mode(paragraph-formatting);

body {
  hdocx-text-align: justify;
  hdocx-line-spacing-exact: 18pt;
  hdocx-first-line-indent: 2char;
  hdocx-space-before: 0;
  hdocx-space-after: 0;
}
```

Supported paragraph declarations:

| Declaration | Value | OOXML mapping |
| --- | --- | --- |
| `hdocx-text-align` / `hdocx-align` | `left`, `center`, `right`, `justify`/`both` | `w:pPr/w:jc @w:val` |
| `hdocx-first-line-indent` | non-negative `char` or `pt` | `w:pPr/w:ind @w:firstLineChars` or `@w:firstLine` |
| `hdocx-line-spacing` | positive multiple or exact `pt` | `w:pPr/w:spacing @w:line` and `@w:lineRule` |
| `hdocx-line-spacing-exact` | positive `pt` | `w:pPr/w:spacing @w:lineRule="exact"` |
| `hdocx-space-before` | `0`, non-negative `pt`, or `line` | `w:pPr/w:spacing @w:before` or `@w:beforeLines` |
| `hdocx-space-after` | `0`, non-negative `pt`, or `line` | `w:pPr/w:spacing @w:after` or `@w:afterLines` |
| `hdocx-manual-page-break-before` | `true` or `false` | inserts an idempotent manual page-break paragraph before the target paragraph |

### Paragraph Structure

Use `@hdocx-edit mode(paragraph-structure);` when the required output is a real
Word structural paragraph, not just visual spacing on an existing paragraph.
This is the mode for blank lines. `plan` reports structural operations such as
`insert-empty-paragraph-after`; `apply` skips insertion if the neighboring blank
paragraph already exists; `diff` reports blank-line changes under
`emptyParagraphDiff` and aligns later semantic nodes so unchanged text is not
misreported as edited.

```css
@hdocx-edit mode(paragraph-structure);

#p-000010 {
  hdocx-insert-empty-paragraph-after: true;
  hdocx-empty-paragraph-line-spacing-exact: 12pt;
  hdocx-empty-paragraph-space-before: 0;
  hdocx-empty-paragraph-space-after: 0;
}
```

Supported paragraph-structure declarations:

| Declaration | Value | OOXML mapping |
| --- | --- | --- |
| `hdocx-insert-empty-paragraph-before` | `true` or `false` | inserts an idempotent empty `<w:p>` before the target paragraph |
| `hdocx-insert-empty-paragraph-after` | `true` or `false` | inserts an idempotent empty `<w:p>` after the target paragraph |
| `hdocx-empty-paragraph-style-id` | existing simple style id | inserted empty paragraph `w:pPr/w:pStyle @w:val` |
| `hdocx-empty-paragraph-line-spacing` | positive multiple or exact `pt` | inserted empty paragraph `w:pPr/w:spacing` |
| `hdocx-empty-paragraph-line-spacing-exact` | positive `pt` | inserted empty paragraph `w:pPr/w:spacing @w:lineRule="exact"` |
| `hdocx-empty-paragraph-space-before` | `0`, non-negative `pt`, or `line` | inserted empty paragraph `w:pPr/w:spacing @w:before` or `@w:beforeLines` |
| `hdocx-empty-paragraph-space-after` | `0`, non-negative `pt`, or `line` | inserted empty paragraph `w:pPr/w:spacing @w:after` or `@w:afterLines` |

Use paragraph spacing (`hdocx-space-before` / `hdocx-space-after`) when the
document only needs visual separation. Use paragraph structure when Word should
contain a real blank paragraph that can have its own line spacing and paragraph
spacing.

### Function Selectors

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

Selector support is intentionally small: ids, classes, exact attribute
selectors, class+attribute compounds such as
`.hdocx-r[data-hdocx-id="r-000001"]`, selector lists separated by commas,
and the H-DOCX functions above.

For safe grouping, keep groups in `agent.edits.hcss`; do not add custom classes
or other projection metadata to `document.html`. A selector list may be used
directly in a rule or inside `@hdocx-set`:

```css
@hdocx-set body {
  select: id(p-000007), id(p-000008), id(p-000009);
}

.role-body,
.role-reference {
  hdocx-font-size: 10.5pt;
}
```

### Run Formatting

```css
@hdocx-edit mode(all-runs);

#r-000001 {
  hdocx-font-family: "Times New Roman";
  hdocx-eastAsia-font: "SimSun";
  hdocx-font-size: 12pt;
  hdocx-bold: true;
}
```

Supported run declarations:

| Declaration | Value | OOXML mapping |
| --- | --- | --- |
| `hdocx-font-family` | quoted or bare font name | `w:rFonts @w:ascii` and `@w:hAnsi` |
| `hdocx-eastAsia-font` / `hdocx-east-asia-font` | quoted or bare font name | `w:rFonts @w:eastAsia` |
| `hdocx-ascii-font` | quoted or bare font name | `w:rFonts @w:ascii` |
| `hdocx-hansi-font` | quoted or bare font name | `w:rFonts @w:hAnsi` |
| `hdocx-cs-font` | quoted or bare font name | `w:rFonts @w:cs` |
| `hdocx-font-size` | positive `pt`, including `10.5pt` | `w:sz` half-points |
| `hdocx-bold` | `true` or `false` | `w:b` |
| `hdocx-italic` | `true` or `false` | `w:i` |
| `hdocx-color` | `#RRGGBB` | `w:color @w:val` |

### Image Formatting

Use `@hdocx-edit mode(image-formatting);` for existing projected images. It can
change DrawingML metadata and size, and it can also patch the host paragraph's
spacing so inline images are not clipped by an inherited exact line spacing.

```css
@hdocx-edit mode(image-formatting);

#r-000001 {
  hdocx-width-emu: 1828800;
  hdocx-height-emu: 914400;
  hdocx-alt: "Scaled figure";
  hdocx-paragraph-line-spacing: 1;
  hdocx-paragraph-space-before: 0;
  hdocx-paragraph-space-after: 0;
}
```

Supported image declarations:

| Declaration | Value | OOXML mapping |
| --- | --- | --- |
| `hdocx-alt` | quoted or bare text | `wp:docPr @descr` |
| `hdocx-width-emu` | positive EMU integer | `wp:extent @cx` |
| `hdocx-height-emu` | positive EMU integer | `wp:extent @cy` |
| `hdocx-paragraph-line-spacing` | positive multiple or exact `pt` | host paragraph `w:pPr/w:spacing` |
| `hdocx-paragraph-line-spacing-exact` | positive `pt` | host paragraph exact line spacing |
| `hdocx-paragraph-space-before` | `0`, non-negative `pt`, or `line` | host paragraph spacing before |
| `hdocx-paragraph-space-after` | `0`, non-negative `pt`, or `line` | host paragraph spacing after |
| `hdocx-paragraph-text-align` / `hdocx-paragraph-align` | `left`, `center`, `right`, `justify`/`both` | host paragraph alignment |

When a picture appears hidden or cropped after applying academic body line
spacing, target the image run and set `hdocx-paragraph-line-spacing: 1` (auto
single spacing) or another non-clipping host paragraph spacing.

### Paper Body Formatting

```css
@hdocx-set body {
  select: style(BodyText);
}

@hdocx-edit mode(paragraph-formatting);

body {
  hdocx-text-align: justify;
  hdocx-first-line-indent: 2char;
  hdocx-line-spacing-exact: 18pt;
  hdocx-space-before: 0;
  hdocx-space-after: 0;
}

@hdocx-edit mode(all-runs);

body {
  hdocx-font-family: "Times New Roman";
  hdocx-eastAsia-font: "SimSun";
  hdocx-font-size: 10.5pt;
}
```

### Manual Page Breaks

Use `hdocx-manual-page-break-before: true` when the required output is an
explicit manual page break, not Word's automatic `pageBreakBefore` paragraph
property. `plan` reports this as `insert-manual-page-break-before`; `apply`
inserts a standalone page-break paragraph before the target and skips insertion
when the same break already exists. `diff` reports these changes under
`manualPageBreakDiff` and aligns later semantic nodes so the target text is not
misreported as changed merely because a break paragraph was inserted.

```css
@hdocx-edit mode(paragraph-formatting);

.hdocx-p[data-hdocx-id="p-000006"] {
  hdocx-manual-page-break-before: true;
}
```

### Image Insertion

```css
@hdocx-insert-image after(#p-000001) {
  source: assets/figure.png;
  alt: "Figure 1";
  width-emu: 914400;
  height-emu: 457200;
}
```

### Table Row Insertion

```css
@hdocx-insert-table-row after(#tr-000001) {
  cells: "New A|New B";
}
```

### Style Creation And Assignment

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

### Numbering Definition Edit

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

### Equation Replacement

```css
@hdocx-edit mode(equation-omml);

[data-hdocx-protected-kind="equation"] {
  hdocx-omml-source: equations/replacement.omml;
}
```

## MCP Agent Integration

The repository exposes the DOCX workflow as a local stdio MCP server:

```text
html-docx-mcp
```

Add `html-docx-mcp` to `PATH`, then configure your MCP client with JSON:

```json
{
  "mcpServers": {
    "hdocx": {
      "command": "html-docx-mcp",
      "args": []
    }
  }
}
```

The MCP server provides tools such as `hdocx_create`, `hdocx_audit`, `hdocx_export`,
`hdocx_query`, `hdocx_find`, `hdocx_inspect`, `hdocx_plan`, `hdocx_apply`,
`hdocx_diff`, `hdocx_assert`, `hdocx_check`, `hdocx_batch_check`,
and `hdocx_guidance`.

It also exposes agent-facing resources and prompts:

```text
hdocx://guide/workflow
hdocx://guide/writing-format
hdocx://guide/query
hdocx://guide/hcss
hdocx://guide/acceptance
hdocx://guide/edge-cases

hdocx_create_docx
hdocx_safe_edit
hdocx_format_change
hdocx_roundtrip_check
```

Agents should read the relevant resources before editing. Clients that do not
surface MCP resources or prompts can call `hdocx_guidance` to retrieve the same
authoring rules through a normal tool response.

Every file-oriented tool accepts an optional `root` argument. All file paths must
resolve inside that root. If `root` is omitted, the server uses
`HDOCX_MCP_ROOT`, then `CLAUDE_PROJECT_DIR`, then the MCP server current
directory.

Tool calls are intended to be serialized. If two tool handlers overlap inside
one server instance, the server returns structured `MCP_SERVER_BUSY` instead of
risking a broken stdio transport. Some clients queue quick calls even when an
agent asks for them in parallel; those calls may both succeed because they
reached the server sequentially.

`hdocx_doctor` reports the loaded module path, package version, git commit or
module timestamp, supported H-CSS edit modes, supported H-CSS properties, stdio
encoding, and guidance source hash. If README and MCP behavior disagree, run
`hdocx_doctor` first to verify which code the MCP process actually loaded.
After upgrading this project, restart the MCP client or server process; a
long-running MCP process keeps using the code it loaded at startup.

Agent-facing policy:

- Inspect before broad edits.
- For new documents, call `hdocx_create` instead of writing OOXML by hand.
- Prefer ids, named sets, and narrow selectors.
- Treat advanced structures as protected unless a dedicated mode supports the
  requested operation.
- Always run `plan`, `apply`, and `diff` before claiming success.

## Validation And QA

Use validation according to the risk of the task:

- Source changes: `python -m unittest discover -s tests`
- New DOCX from scratch: `html-docx create` followed by `html-docx check`
- New or unknown DOCX: `html-docx audit` and `html-docx check`
- Edited DOCX: `html-docx apply` followed by `html-docx diff`
- Conversion logic changes: `generate-fixtures` followed by `batch-check`

## Supported Editing Surface

Currently supported areas include:

- Canonical new DOCX creation from the built-in blank template.
- Editable run text.
- Run formatting and run split operations.
- Paragraph formatting.
- Paragraph style assignment, style creation, and safe unused-style deletion.
- Numbering/list metadata projection and controlled numbering-level edits.
- New single-level and multi-level list creation.
- Table-cell text and simple non-merged table row/column operations.
- Headers, footnotes, and endnotes as projected secondary parts.
- Image alt text, size metadata, controlled media replacement, and image
  insertion.
- Controlled comment text edits.
- Revision accept/reject actions.
- Whole-equation OMML replacement from bundle-local files.
- Package, semantic, and fragment-level diffs.

## Preservation Policy

High-risk structures are preserved and protected unless a dedicated operation
supports the exact requested change:

- custom XML
- charts
- SmartArt
- OLE
- AlternateContent
- fields
- equations
- comments
- revisions
- text boxes
- VML

This is intentional. Strict reversibility is more important than pretending all
Word features are ordinary HTML.

## Repository Layout

```text
src/html_docx/                  # CLI and library implementation
tests/                          # unit and round-trip tests
AGENTS.md                       # repository-level agent rules
CLAUDE.md                       # Claude Code entry point
USAGE_AND_PRINCIPLES.md         # full usage and design explanation
```

## Development

Run tests:

```powershell
$env:PYTHONPATH = "src"
python -m unittest discover -s tests
```

Build a wheel in a local environment:

```powershell
.\.venv\Scripts\python.exe -m pip wheel . --no-deps --no-build-isolation -w dist
```

Generated fixtures, render outputs, temporary bundles, local DOCX samples, and
local virtual environments are ignored by git.

## Documentation Map

- `USAGE_AND_PRINCIPLES.md`: usage and implementation principles.
- `FUNCTIONAL_SPEC.md`: functional boundary.
- `HDOCX_HTML_DESIGN.md`: H-DOCX/HTML representation design.
- `SELECTOR_AND_REUSE_DESIGN.md`: selector and reuse model.
- `SELECTOR_EDGE_CASES_AND_GUARDS.md`: selector edge cases and safeguards.
- `EDITING_EDGE_CASES.md`: editing edge cases.
- `EDGE_CASE_TEST_MATRIX.md`: edge and pressure test matrix.
- `SOFTWARE_ARCHITECTURE.md`: architecture.
- `GLOBAL_DELIVERY_PLAN.md`: delivery plan.
- `COMPLETION_PLAN.md`: completion gates.
- `IMPLEMENTATION_STATUS.md`: implementation status.
- `AGENT_GUIDE.md`: detailed agent workflow.
- `PRESSURE_FIXTURES.md`: pressure fixture coverage.
- `RELEASE_CHECKLIST.md`: release verification checklist.
