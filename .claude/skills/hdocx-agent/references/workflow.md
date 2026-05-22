# H-DOCX Workflow

Use these commands from the workspace that contains the DOCX files. If working
inside the H-DOCX source repo, first set:

```powershell
$env:PYTHONPATH = "src"
```

## No-Edit Reversibility Check

```powershell
python -m html_docx check input.docx --work check.hdocx --out checked.docx --force --report check.json
```

Success requires:

- `ok: true`
- `acceptance.byteIdentical: true`
- `acceptance.semanticIdentical: true`

## Controlled Editing

```powershell
python -m html_docx audit input.docx --report audit.json
python -m html_docx export input.docx --out work.hdocx --force
python -m html_docx inspect work.hdocx --kind node --id p-000001
python -m html_docx inspect work.hdocx --kind style --id Normal
python -m html_docx plan work.hdocx --report plan.json
python -m html_docx apply work.hdocx --out output.docx --report apply.json
python -m html_docx diff input.docx output.docx --report diff.json
```

Inspect before broad H-CSS edits:

```powershell
python -m html_docx inspect work.hdocx --kind list --id 1
python -m html_docx inspect work.hdocx --kind table --id tbl-000001
python -m html_docx inspect work.hdocx --kind image --id r-000001
```

## Pressure Fixtures

Run after conversion logic changes:

```powershell
python -m html_docx generate-fixtures --out pressure-fixtures --force --report pressure-fixtures.json
python -m html_docx batch-check pressure-fixtures --work pressure-work --out pressure-out --force --report pressure.json
```

## Optional Render QA

Use when LibreOffice/soffice is available, or when edited layout needs visual
confirmation:

```powershell
python -m html_docx render-check output.docx --out render-out --force --allow-missing --report render.json
```

`renderer-missing` means the external renderer was not available. It does not
invalidate byte-identical no-edit checks.
