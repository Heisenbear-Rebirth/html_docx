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

- `70` unit tests passing.
- Built-in pressure fixture round-trip: `6/6` byte-identical.
- Real sample round-trip: byte-identical and semantically identical.
- MCP stdio smoke test passing.
- Optional render QA is supported when LibreOffice/soffice is available.

## Requirements

- Windows PowerShell for the bundled installer script.
- Python `>=3.11`.
- No Python runtime dependencies are required by the package itself.
- Optional: LibreOffice/soffice on `PATH` for render-based QA.

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
| `doctor` | Report runtime capabilities and optional renderer availability. |
| `audit` | Detect high-risk DOCX structures and preservation policies. |
| `export` | Convert a DOCX into an H-DOCX bundle. |
| `validate` | Validate an H-DOCX bundle before applying it. |
| `inspect` | Inspect nodes, styles, lists, tables, or images by id. |
| `plan` | Plan edits without writing a DOCX. |
| `apply` | Apply a bundle back to DOCX. |
| `diff` | Compare two DOCX packages with package, semantic, and fragment reports. |
| `roundtrip` | Export and apply without edits. |
| `check` | Run export/apply/diff acceptance for one DOCX. |
| `batch-check` | Run `check` over a file or directory. |
| `generate-fixtures` | Generate synthetic pressure DOCX fixtures. |
| `render-check` | Optionally render DOCX through LibreOffice/soffice. |
| `mcp` | Run the stdio MCP server. |

The package also installs a dedicated `html-docx-mcp` command for MCP clients.

## H-CSS Examples

H-CSS is a small project-specific edit language. It is not browser CSS.

### Paragraph Formatting

```css
@hdocx-set body {
  select: [data-hdocx-type="paragraph"];
}

@hdocx-edit mode(paragraph-formatting);

body {
  hdocx-align: center;
  hdocx-line-spacing: 1.5;
  hdocx-first-line-indent: 2char;
}
```

### Function Selectors

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

### Run Formatting

```css
@hdocx-edit mode(run-formatting);

#r-000001 {
  hdocx-font-size: 12pt;
  hdocx-bold: true;
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

The MCP server provides tools such as `hdocx_audit`, `hdocx_export`,
`hdocx_plan`, `hdocx_apply`, `hdocx_diff`, `hdocx_check`,
`hdocx_batch_check`, `hdocx_inspect`, `hdocx_render_check`, and
`hdocx_guidance`.

It also exposes agent-facing resources and prompts:

```text
hdocx://guide/workflow
hdocx://guide/writing-format
hdocx://guide/hcss
hdocx://guide/acceptance
hdocx://guide/edge-cases

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

Agent-facing policy:

- Inspect before broad edits.
- Prefer ids, named sets, and narrow selectors.
- Treat advanced structures as protected unless a dedicated mode supports the
  requested operation.
- Always run `plan`, `apply`, and `diff` before claiming success.

## Validation And QA

Use validation according to the risk of the task:

- Source changes: `python -m unittest discover -s tests`
- New or unknown DOCX: `html-docx audit` and `html-docx check`
- Edited DOCX: `html-docx apply` followed by `html-docx diff`
- Conversion logic changes: `generate-fixtures` followed by `batch-check`
- Layout-sensitive edited output: `render-check` when LibreOffice/soffice is
  available

`render-check` is optional because it depends on an external renderer. A
`renderer-missing` report means the renderer was not available; it does not
invalidate byte-identical no-edit checks.

## Supported Editing Surface

Currently supported areas include:

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
