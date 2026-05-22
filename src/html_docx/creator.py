from __future__ import annotations

import html
import zipfile
from pathlib import Path
from typing import Iterable

from .errors import HDocxError
from .utils import sha256_file


CANONICAL_TEMPLATE = "blank"
ZIP_DATE_TIME = (2024, 1, 1, 0, 0, 0)


def create_canonical_docx(
    output_docx: Path,
    *,
    title: str | None = None,
    paragraphs: Iterable[str] | None = None,
    template: str = CANONICAL_TEMPLATE,
    force: bool = False,
) -> dict[str, object]:
    output_docx = output_docx.resolve()
    if template != CANONICAL_TEMPLATE:
        raise HDocxError("CREATE_TEMPLATE_UNSUPPORTED", "Unsupported DOCX creation template.", {"template": template})
    if output_docx.suffix.lower() != ".docx":
        raise HDocxError("CREATE_OUTPUT_NOT_DOCX", "Created output must be a .docx file.", {"path": str(output_docx)})
    if output_docx.exists():
        if not force:
            raise HDocxError("CREATE_OUTPUT_EXISTS", "Output DOCX already exists.", {"path": str(output_docx)})
        if output_docx.is_dir():
            raise HDocxError("CREATE_OUTPUT_IS_DIRECTORY", "Output DOCX path is a directory.", {"path": str(output_docx)})
        output_docx.unlink()

    paragraph_list = list(paragraphs or [])
    if not paragraph_list:
        paragraph_list = [""]
    files = _blank_docx_files(title=title, paragraphs=paragraph_list)
    _write_deterministic_docx(output_docx, files)
    return {
        "ok": True,
        "command": "create",
        "output": str(output_docx),
        "template": template,
        "title": title,
        "paragraphCount": len(paragraph_list),
        "entryCount": len(files),
        "sha256": sha256_file(output_docx),
        "size": output_docx.stat().st_size,
    }


def _write_deterministic_docx(path: Path, files: dict[str, str | bytes]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    with zipfile.ZipFile(path, "w", compression=zipfile.ZIP_DEFLATED) as zf:
        for entry_name, payload in files.items():
            info = zipfile.ZipInfo(entry_name, ZIP_DATE_TIME)
            info.compress_type = zipfile.ZIP_DEFLATED
            info.create_system = 0
            data = payload.encode("utf-8") if isinstance(payload, str) else payload
            zf.writestr(info, data)


def _blank_docx_files(*, title: str | None, paragraphs: list[str]) -> dict[str, str | bytes]:
    return {
        "[Content_Types].xml": _content_types(),
        "_rels/.rels": _root_rels(),
        "docProps/core.xml": _core_properties(title),
        "docProps/app.xml": _app_properties(),
        "word/document.xml": _document_xml(title=title, paragraphs=paragraphs),
        "word/_rels/document.xml.rels": _document_rels(),
        "word/styles.xml": _styles_xml(),
        "word/settings.xml": _settings_xml(),
    }


def _content_types() -> str:
    return """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>
  <Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
  <Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
  <Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/>
</Types>
"""


def _root_rels() -> str:
    return """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>
  <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>
</Relationships>
"""


def _document_rels() -> str:
    return """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/>
</Relationships>
"""


def _core_properties(title: str | None) -> str:
    title_xml = _text_element("dc:title", title) if title else ""
    return f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<cp:coreProperties
  xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties"
  xmlns:dc="http://purl.org/dc/elements/1.1/"
  xmlns:dcterms="http://purl.org/dc/terms/"
  xmlns:dcmitype="http://purl.org/dc/dcmitype/"
  xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
  {title_xml}
  <dc:creator>H-DOCX</dc:creator>
  <cp:lastModifiedBy>H-DOCX</cp:lastModifiedBy>
  <dcterms:created xsi:type="dcterms:W3CDTF">2024-01-01T00:00:00Z</dcterms:created>
  <dcterms:modified xsi:type="dcterms:W3CDTF">2024-01-01T00:00:00Z</dcterms:modified>
</cp:coreProperties>
"""


def _app_properties() -> str:
    return """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"
  xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">
  <Application>H-DOCX</Application>
  <DocSecurity>0</DocSecurity>
  <ScaleCrop>false</ScaleCrop>
  <LinksUpToDate>false</LinksUpToDate>
  <SharedDoc>false</SharedDoc>
  <HyperlinksChanged>false</HyperlinksChanged>
  <AppVersion>0.1</AppVersion>
</Properties>
"""


def _document_xml(*, title: str | None, paragraphs: list[str]) -> str:
    parts: list[str] = []
    if title:
        parts.append(_paragraph_xml(title, style_id="Title"))
    parts.extend(_paragraph_xml(paragraph, style_id="BodyText") for paragraph in paragraphs)
    body = "\n".join(parts)
    return f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document
  xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
  xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <w:body>
{body}
    <w:sectPr>
      <w:pgSz w:w="12240" w:h="15840"/>
      <w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/>
    </w:sectPr>
  </w:body>
</w:document>
"""


def _paragraph_xml(text: str, *, style_id: str) -> str:
    return f"""    <w:p>
      <w:pPr><w:pStyle w:val="{style_id}"/></w:pPr>
      <w:r>{_text_xml(text)}</w:r>
    </w:p>"""


def _text_xml(text: str) -> str:
    space = ' xml:space="preserve"' if text[:1].isspace() or text[-1:].isspace() else ""
    return f"<w:t{space}>{html.escape(text, quote=False)}</w:t>"


def _text_element(name: str, text: str | None) -> str:
    if text is None:
        return ""
    return f"<{name}>{html.escape(text, quote=False)}</{name}>"


def _styles_xml() -> str:
    return """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:docDefaults>
    <w:rPrDefault>
      <w:rPr>
        <w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman" w:eastAsia="SimSun"/>
        <w:sz w:val="24"/>
        <w:szCs w:val="24"/>
        <w:lang w:val="en-US" w:eastAsia="zh-CN"/>
      </w:rPr>
    </w:rPrDefault>
    <w:pPrDefault>
      <w:pPr>
        <w:spacing w:after="0" w:line="240" w:lineRule="auto"/>
      </w:pPr>
    </w:pPrDefault>
  </w:docDefaults>
  <w:style w:type="paragraph" w:default="1" w:styleId="Normal">
    <w:name w:val="Normal"/>
    <w:qFormat/>
  </w:style>
  <w:style w:type="paragraph" w:styleId="Title">
    <w:name w:val="Title"/>
    <w:basedOn w:val="Normal"/>
    <w:next w:val="BodyText"/>
    <w:qFormat/>
    <w:pPr>
      <w:jc w:val="center"/>
      <w:spacing w:after="240" w:line="240" w:lineRule="auto"/>
    </w:pPr>
    <w:rPr>
      <w:b/>
      <w:sz w:val="32"/>
      <w:szCs w:val="32"/>
    </w:rPr>
  </w:style>
  <w:style w:type="paragraph" w:styleId="BodyText">
    <w:name w:val="Body Text"/>
    <w:basedOn w:val="Normal"/>
    <w:qFormat/>
    <w:pPr>
      <w:spacing w:after="120" w:line="360" w:lineRule="auto"/>
    </w:pPr>
    <w:rPr>
      <w:sz w:val="24"/>
      <w:szCs w:val="24"/>
    </w:rPr>
  </w:style>
</w:styles>
"""


def _settings_xml() -> str:
    return """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:zoom w:percent="100"/>
  <w:defaultTabStop w:val="720"/>
  <w:characterSpacingControl w:val="doNotCompress"/>
</w:settings>
"""
