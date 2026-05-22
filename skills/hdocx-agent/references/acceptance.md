# Acceptance Rules

## No-Edit Round-Trips

For `input.docx -> work.hdocx -> checked.docx`, success requires:

- `ok: true`
- `byteIdentical: true`
- `semanticIdentical: true`
- Matching SHA256 values

If SHA256 is identical, the DOCX byte stream is identical. For the same renderer
and environment, rendered output should also be identical.

## Edited Output

Edited output does not need byte identity, but the report must prove that every
change is intended:

- `apply.json` must show the expected patch count and target ids.
- `patchSummary.byRiskClass` should be reviewed.
- `diff.json` must show no unrelated changed entries.
- Fragment/semantic diffs must match the requested edit.

## High-Risk Content

Treat these as protected unless a dedicated mode supports the exact operation:

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

Preserve protected content exactly. Do not rewrite it through normal HTML.

## Failure Policy

If safety cannot be proven, stop with the generated report and explain the
blocker. Do not guess, normalize, reserialize arbitrary OOXML, or silently drop
markup.
