# Pressure Fixture Workflow

The repository does not bundle real user DOCX files. Keep private or licensed fixtures outside version control, but place local copies under a workspace directory when running validation, for example:

```text
fixtures/
  01-basic-body.docx
  02-mixed-runs.docx
  03-headings-numbering.docx
  04-toc-crossrefs.docx
  05-footnotes-endnotes.docx
  06-comments.docx
  07-revisions.docx
  08-tables.docx
  09-images.docx
  10-equations.docx
  11-headers-footers.docx
  12-customxml-alternatecontent.docx
  13-east-asian-layout.docx
```

Run the strict unedited reversibility chain:

```powershell
$env:PYTHONPATH = "src"
python -m html_docx generate-fixtures --out pressure-fixtures --force --report pressure-fixtures.json
python -m html_docx batch-check fixtures --work pressure-work --out pressure-out --force --report pressure.json
```

`generate-fixtures` creates a synthetic local pressure suite. Use your private real-world fixture directory in place of `fixtures` when validating journal, thesis, WPS, or Word-generated documents.

For unfamiliar fixtures, run `audit` first or after export to classify high-risk structures:

```powershell
python -m html_docx audit fixtures/12-customxml-alternatecontent.docx --report audit.json
```

Acceptance for each fixture:

- `acceptance.byteIdentical` is `true`.
- `acceptance.contentIdentical` is `true`.
- `acceptance.semanticIdentical` is `true`.
- `changedEntries`, `leftOnlyEntries`, and `rightOnlyEntries` are all `0`.

When a fixture fails, keep the generated `.hdocx` bundle and inspect:

```powershell
python -m html_docx validate pressure-work/0001-name.hdocx --report validate.json
python -m html_docx diff fixtures/name.docx pressure-out/0001-name.docx --report diff.json
```

Use `diff.fragmentDiff` to locate the changed byte window for each changed package entry and to see the projected H-DOCX node ids linked to that entry.

Optional visual render QA:

```powershell
python -m html_docx render-check fixtures/name.docx --out render-name --force --allow-missing --report render-name.json
```

When LibreOffice/soffice is unavailable, `--allow-missing` records `renderer-missing` instead of changing the system environment.

Do not broaden edit support from a failing fixture until the failure is classified:

- Projection failure: the DOCX contains a structure the projector cannot parse.
- Preservation failure: unedited roundtrip changed bytes or package entries.
- Semantic-diff limitation: bytes are identical but semantic projection comparison failed.
- Unsupported edit: the structure is preserved but intentionally not editable.
