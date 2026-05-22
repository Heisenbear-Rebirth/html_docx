from __future__ import annotations

from pathlib import Path
from typing import Any
import zipfile

from .utils import ensure_new_dir, sha256_file


def generate_pressure_fixtures(output_dir: Path, *, force: bool = False) -> dict[str, Any]:
    output_dir = output_dir.resolve()
    ensure_new_dir(output_dir, force=force)
    fixtures = [
        ("01-minimal.docx", _minimal_fixture),
        ("02-styles-numbering.docx", _styles_numbering_fixture),
        ("03-table.docx", _table_fixture),
        ("04-header-footnote-image.docx", _header_footnote_image_fixture),
        ("05-comments-revisions-equations.docx", _comments_revisions_equations_fixture),
        ("06-advanced-protected.docx", _advanced_protected_fixture),
    ]
    created = []
    for file_name, factory in fixtures:
        path = output_dir / file_name
        _write_docx(path, factory())
        created.append({"fileName": file_name, "path": str(path), "sha256": sha256_file(path), "size": path.stat().st_size})
    return {
        "ok": True,
        "command": "generate-fixtures",
        "output": str(output_dir),
        "count": len(created),
        "fixtures": created,
    }


def _write_docx(path: Path, files: dict[str, str | bytes]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    with zipfile.ZipFile(path, "w", compression=zipfile.ZIP_DEFLATED) as zf:
        for entry_name, payload in files.items():
            zf.writestr(entry_name, payload)


def _minimal_fixture() -> dict[str, str | bytes]:
    return {
        "[Content_Types].xml": _content_types(["/word/document.xml"]),
        "_rels/.rels": _root_rels(),
        "word/document.xml": _document_body("<w:p><w:r><w:t>Minimal pressure fixture.</w:t></w:r></w:p>"),
    }


def _styles_numbering_fixture() -> dict[str, str | bytes]:
    return {
        "[Content_Types].xml": _content_types(["/word/document.xml", "/word/styles.xml", "/word/numbering.xml"]),
        "_rels/.rels": _root_rels(),
        "word/_rels/document.xml.rels": _document_rels(
            [
                ("rId2", "styles", "styles.xml"),
                ("rId3", "numbering", "numbering.xml"),
            ]
        ),
        "word/document.xml": _document_body(
            """
<w:p><w:pPr><w:pStyle w:val="BodyText"/></w:pPr><w:r><w:t>Styled paragraph.</w:t></w:r></w:p>
<w:p><w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr></w:pPr><w:r><w:t>Numbered paragraph.</w:t></w:r></w:p>
"""
        ),
        "word/styles.xml": """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:style w:type="paragraph" w:styleId="BodyText"><w:name w:val="Body Text"/><w:rPr><w:sz w:val="24"/></w:rPr></w:style>
</w:styles>
""",
        "word/numbering.xml": """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:abstractNum w:abstractNumId="0"><w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%1."/></w:lvl></w:abstractNum>
  <w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num>
</w:numbering>
""",
    }


def _table_fixture() -> dict[str, str | bytes]:
    return {
        "[Content_Types].xml": _content_types(["/word/document.xml"]),
        "_rels/.rels": _root_rels(),
        "word/document.xml": _document_body(
            """
<w:tbl>
  <w:tr><w:tc><w:p><w:r><w:t>A1</w:t></w:r></w:p></w:tc><w:tc><w:p><w:r><w:t>B1</w:t></w:r></w:p></w:tc></w:tr>
  <w:tr><w:tc><w:p><w:r><w:t>A2</w:t></w:r></w:p></w:tc><w:tc><w:p><w:r><w:t>B2</w:t></w:r></w:p></w:tc></w:tr>
</w:tbl>
"""
        ),
    }


def _header_footnote_image_fixture() -> dict[str, str | bytes]:
    return {
        "[Content_Types].xml": _content_types(["/word/document.xml", "/word/header1.xml", "/word/footnotes.xml"], {"png": "image/png"}),
        "_rels/.rels": _root_rels(),
        "word/_rels/document.xml.rels": _document_rels(
            [
                ("rId2", "header", "header1.xml"),
                ("rId3", "footnotes", "footnotes.xml"),
                ("rId4", "image", "media/image1.png"),
            ]
        ),
        "word/document.xml": _document_body(
            """
<w:p><w:r><w:t>Main text with note</w:t></w:r><w:r><w:footnoteReference w:id="2"/></w:r></w:p>
<w:p><w:r><w:drawing><wp:inline xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"><wp:extent cx="914400" cy="457200"/><wp:docPr id="1" name="Picture 1" descr="Fixture image"/><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic><pic:blipFill><a:blip r:embed="rId4"/></pic:blipFill></pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r></w:p>
<w:sectPr><w:headerReference w:type="default" r:id="rId2"/></w:sectPr>
"""
        ),
        "word/header1.xml": """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:hdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:p><w:r><w:t>Header fixture.</w:t></w:r></w:p></w:hdr>
""",
        "word/footnotes.xml": """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:footnotes xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:footnote w:id="2"><w:p><w:r><w:t>Footnote fixture.</w:t></w:r></w:p></w:footnote></w:footnotes>
""",
        "word/media/image1.png": b"fixturepng",
    }


def _comments_revisions_equations_fixture() -> dict[str, str | bytes]:
    return {
        "[Content_Types].xml": _content_types(["/word/document.xml", "/word/comments.xml"]),
        "_rels/.rels": _root_rels(),
        "word/_rels/document.xml.rels": _document_rels([("rId2", "comments", "comments.xml")]),
        "word/document.xml": _document_body(
            """
<w:p><w:commentRangeStart w:id="0"/><w:r><w:t>Commented text.</w:t></w:r><w:commentRangeEnd w:id="0"/><w:r><w:commentReference w:id="0"/></w:r></w:p>
<w:p><w:ins w:id="7" w:author="Reviewer"><w:r><w:t>Inserted fixture.</w:t></w:r></w:ins></w:p>
<w:p><m:oMath xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math"><m:r><m:t>x=1</m:t></m:r></m:oMath></w:p>
"""
        ),
        "word/comments.xml": """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:comments xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:comment w:id="0" w:author="Reviewer"><w:p><w:r><w:t>Comment fixture.</w:t></w:r></w:p></w:comment></w:comments>
""",
    }


def _advanced_protected_fixture() -> dict[str, str | bytes]:
    return {
        "[Content_Types].xml": _content_types(
            ["/word/document.xml", "/word/charts/chart1.xml", "/word/diagrams/data1.xml"],
            {"bin": "application/vnd.openxmlformats-officedocument.oleObject"},
        ),
        "_rels/.rels": _root_rels(),
        "word/_rels/document.xml.rels": _document_rels(
            [
                ("rId2", "chart", "charts/chart1.xml"),
                ("rId3", "diagramData", "diagrams/data1.xml"),
                ("rId4", "oleObject", "embeddings/oleObject1.bin"),
            ]
        ),
        "word/document.xml": _document_body(
            """
<mc:AlternateContent xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"><mc:Choice Requires="w14"><w:p><w:r><w:t>Choice branch.</w:t></w:r></w:p></mc:Choice><mc:Fallback><w:p><w:r><w:t>Fallback branch.</w:t></w:r></w:p></mc:Fallback></mc:AlternateContent>
<w:p><w:pict xmlns:v="urn:schemas-microsoft-com:vml"><v:shape><v:textbox><w:txbxContent><w:p><w:r><w:t>Text box.</w:t></w:r></w:p></w:txbxContent></v:textbox></v:shape></w:pict></w:p>
"""
        ),
        "word/charts/chart1.xml": '<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"/>',
        "word/diagrams/data1.xml": '<dgm:dataModel xmlns:dgm="http://schemas.openxmlformats.org/drawingml/2006/diagram"/>',
        "word/embeddings/oleObject1.bin": b"ole",
        "customXml/item1.xml": "<root/>",
    }


def _content_types(overrides: list[str], defaults: dict[str, str] | None = None) -> str:
    defaults = defaults or {}
    content_type_map = {
        "/word/document.xml": "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml",
        "/word/styles.xml": "application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml",
        "/word/numbering.xml": "application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml",
        "/word/header1.xml": "application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml",
        "/word/footnotes.xml": "application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml",
        "/word/comments.xml": "application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml",
        "/word/charts/chart1.xml": "application/vnd.openxmlformats-officedocument.drawingml.chart+xml",
        "/word/diagrams/data1.xml": "application/vnd.openxmlformats-officedocument.drawingml.diagramData+xml",
    }
    default_xml = "\n".join(
        f'  <Default Extension="{extension}" ContentType="{content_type}"/>'
        for extension, content_type in {"rels": "application/vnd.openxmlformats-package.relationships+xml", "xml": "application/xml", **defaults}.items()
    )
    override_xml = "\n".join(
        f'  <Override PartName="{part_name}" ContentType="{content_type_map[part_name]}"/>'
        for part_name in overrides
    )
    return f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
{default_xml}
{override_xml}
</Types>
"""


def _root_rels() -> str:
    return """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>
"""


def _document_rels(rels: list[tuple[str, str, str]]) -> str:
    lines = [
        f'  <Relationship Id="{rel_id}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/{rel_type}" Target="{target}"/>'
        for rel_id, rel_type, target in rels
    ]
    return f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
{chr(10).join(lines)}
</Relationships>
"""


def _document_body(inner_xml: str) -> str:
    return f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document
  xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
  xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <w:body>
{inner_xml.strip()}
  </w:body>
</w:document>
"""
