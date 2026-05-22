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

## Quick Start

### 1. Clone

```powershell
git clone https://github.com/Heisenbear-Rebirth/html_docx.git
cd html_docx
```

### 2. Verify The Tool

From the repository root:

```powershell
$env:PYTHONPATH = "src"
python -m html_docx doctor
python -m unittest discover -s tests
python -m html_docx generate-fixtures --out pressure-fixtures --force --report pressure-fixtures.json
python -m html_docx batch-check pressure-fixtures --work pressure-work --out pressure-out --force --report pressure.json
```

These commands require no network and no global Python installation changes.

### 3. Install Once For All Workspaces

If you want `html-docx`, the Codex skill, and the Claude Code skill available
outside this repository, run the installer from the repository root:

```powershell
.\scripts\install-hdocx.ps1 -All -AddToUserPath
```

This writes only to user-level locations:

```text
%USERPROFILE%\.hdocx\                 # isolated CLI venv and shims
%CODEX_HOME%\skills\hdocx-agent       # Codex skill, default %USERPROFILE%\.codex\skills\hdocx-agent
%USERPROFILE%\.claude\skills\hdocx-agent
```

Open a new terminal after `-AddToUserPath`, then verify from any directory:

```powershell
html-docx doctor
```

To preview the install without writing files:

```powershell
.\scripts\install-hdocx.ps1 -All -DryRun
```

To install only the agent skills and leave the CLI alone:

```powershell
.\scripts\install-hdocx.ps1 -Codex -Claude
```

Use `-Force` to replace an existing installed skill. Restart Codex and Claude
Code after installing skills so they discover the new user-level workflow.

### 4. Use With Codex

Open Codex Desktop or Codex CLI in the repository root. Codex should treat
`AGENTS.md` as the canonical repository instruction file.

For this repository, no extra Codex configuration is required:

```text
AGENTS.md
```

The reusable Codex skill package is also checked in here:

```text
skills/hdocx-agent/SKILL.md
```

Recommended first prompt in Codex:

```text
Read AGENTS.md and use skills/hdocx-agent/SKILL.md as the workflow for strict DOCX editing. Then run the verification commands from README Quick Start.
```

For a DOCX task, ask Codex naturally:

```text
Use H-DOCX to audit input.docx, export it to a work.hdocx bundle, make only the requested edits, apply it back to DOCX, and prove the diff contains only intended changes.
```

Optional user-level Codex skill installation:

```powershell
.\scripts\install-hdocx.ps1 -Codex
```

Only do this when you intentionally want the skill available outside this
repository. This changes user-level Codex configuration, so this project does
not do it automatically.

After running `.\scripts\install-hdocx.ps1 -Codex`, the same skill is available
outside this checkout from:

```text
%CODEX_HOME%\skills\hdocx-agent
```

Then from any Codex workspace, ask:

```text
Use the hdocx-agent skill and html-docx CLI to edit this DOCX safely.
```

### 5. Use With Claude Code

This repository is already configured for Claude Code.

Claude Code reads `CLAUDE.md`, not `AGENTS.md`, so this repo keeps a small
`CLAUDE.md` file that imports the canonical agent rules:

```markdown
@AGENTS.md
```

The project-level Claude Code skill is checked in at:

```text
.claude/skills/hdocx-agent/SKILL.md
```

Start Claude Code from the repository root:

```powershell
claude
```

Inside Claude Code, invoke the skill directly when you want strict DOCX editing:

```text
/hdocx-agent
```

You can also ask naturally:

```text
Use H-DOCX to audit input.docx and prove whether it round-trips byte-identically.
```

Claude Code should load the project skill automatically for relevant
DOCX/H-DOCX tasks, because the skill description names DOCX inspection,
editing, round-trip, validation, pressure testing, and layout preservation.

After running `.\scripts\install-hdocx.ps1 -Claude`, the same skill is available
outside this checkout from:

```text
%USERPROFILE%\.claude\skills\hdocx-agent
```

### 6. Edit A DOCX

Use the same core workflow in Codex, Claude Code, or a normal terminal:

```powershell
$env:PYTHONPATH = "src"
python -m html_docx audit input.docx --report audit.json
python -m html_docx export input.docx --out work.hdocx --force
python -m html_docx inspect work.hdocx --kind node --id p-000001
# The agent edits work.hdocx/document.html or work.hdocx/agent.edits.hcss
python -m html_docx plan work.hdocx --report plan.json
python -m html_docx apply work.hdocx --out output.docx --report apply.json
python -m html_docx diff input.docx output.docx --report diff.json
```

If you installed the CLI with `-AddToUserPath`, the same workflow can use
`html-docx` directly:

```powershell
html-docx audit input.docx --report audit.json
html-docx export input.docx --out work.hdocx --force
html-docx plan work.hdocx --report plan.json
html-docx apply work.hdocx --out output.docx --report apply.json
html-docx diff input.docx output.docx --report diff.json
```

### 7. Reuse From Another Workspace

Prefer workspace-local approaches:

- Best for general use: run `.\scripts\install-hdocx.ps1 -All -AddToUserPath`
  once, then use `html-docx` plus the user-level Codex/Claude skills anywhere.
- Best for strict workspace isolation: clone or vendor this repository inside
  that workspace and run the agent from the repository root.
- For Codex without user-level install, explicitly point the agent to
  `AGENTS.md` and `skills/hdocx-agent/SKILL.md`.
- For Claude Code without user-level install, add this repository as an
  approved extra directory with `--add-dir` only when you intentionally want
  Claude Code to read the H-DOCX project skill from this checkout.
- For Python-only reuse, install the package into that workspace's own `.venv`;
  do not install it globally.

Do not put private preferences or machine-specific permissions in tracked files.
Use `CLAUDE.local.md` or `.claude/settings.local.json`; both are ignored by git.

References:

- [CLAUDE.md project memory](https://code.claude.com/docs/en/memory)
- [Claude Code skills](https://code.claude.com/docs/en/skills)

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
