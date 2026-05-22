# Completion Plan

This project is delivered as a strict H-DOCX alpha: unedited DOCX packages must roundtrip byte-identically, supported edits must produce explicit patch reports, and unsupported or unsafe edits must fail before writing.

## Completion Gates

1. Core reversible package workflow.
   - Status: complete.
   - Evidence: `export`, `validate`, `plan`, `apply`, `roundtrip`, `diff`, `check`, and `batch-check`.

2. Agent editing surface for common academic writing.
   - Status: complete for V0.1.
   - Covered: run text/formatting, paragraph formatting, styles, numbering/list definitions, paragraph style/list assignment, simple tables, headers/footers/notes/comments, images, revisions, and equation replacement.

3. Strict protection for high-risk Word features.
   - Status: complete for V0.1.
   - Covered: protected projection/audit for customXml, charts, SmartArt, OLE, AlternateContent, VML/text boxes, fields, comments, revisions, equations, notes, headers/footers, and images.

4. Verification and pressure testing.
   - Status: complete for local synthetic pressure plus user sample.
   - Covered: `generate-fixtures`, `batch-check`, `audit`, `fragmentDiff`, optional `render-check`, and `doctor` renderer capability detection.

5. Package delivery.
   - Status: complete for local wheel delivery.
   - Artifact: `dist/html_docx-0.1.0-py3-none-any.whl`.

## Remaining External Gates

These cannot be completed without additional user-provided or environment-provided inputs:

- A larger private corpus of real academic DOCX files from Word, WPS, templates, journal formats, and thesis formats.
- A system renderer such as LibreOffice/soffice or Microsoft Word for visual PDF/page-image comparison.
- Explicit product decisions for semantically editing charts, SmartArt, OLE, and text boxes. Until then they remain audited/protected.

## Recommended Final Validation Loop

```powershell
$env:PYTHONPATH = "src"
python -m unittest discover -s tests
python -m html_docx generate-fixtures --out pressure-fixtures --force --report pressure-fixtures.json
python -m html_docx batch-check pressure-fixtures --work pressure-work --out pressure-out --force --report pressure.json
python -m html_docx audit docx/test.docx --report docx/test-audit.json
python -m html_docx check docx/test.docx --work docx/test-check.hdocx --out docx/test-checked.docx --force --report docx/test-check.json
python -m html_docx render-check docx/test.docx --out render-test --force --allow-missing --report render-test.json
```

If all commands pass and `render-check` is either rendered or reports `renderer-missing` with `--allow-missing`, the project is locally deliverable.
