from __future__ import annotations

import shutil
import sys
import unittest
import zipfile
import xml.etree.ElementTree as ET
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "src"))

from html_docx.hdocx import (
    audit_docx,
    apply_hdocx,
    assert_hdocx,
    create_docx,
    diff_docx,
    export_docx,
    find_hdocx,
    plan_hdocx,
    query_hdocx,
    roundtrip_docx,
    validate_hdocx,
)
from html_docx.utils import read_json, sha256_file


TMP = ROOT / "tests" / "_tmp"
W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
M_NS = "http://schemas.openxmlformats.org/officeDocument/2006/math"
WP_NS = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
NS = {"w": W_NS, "m": M_NS, "wp": WP_NS}


class RoundtripTests(unittest.TestCase):
    def setUp(self) -> None:
        if TMP.exists():
            shutil.rmtree(TMP)
        TMP.mkdir(parents=True)

    def tearDown(self) -> None:
        if TMP.exists():
            shutil.rmtree(TMP)

    def test_unmodified_roundtrip_is_byte_identical(self) -> None:
        input_docx = TMP / "input.docx"
        work = TMP / "work.hdocx"
        output_docx = TMP / "roundtrip.docx"
        _make_minimal_docx(input_docx)

        report = roundtrip_docx(input_docx, work, output_docx, force=True)

        self.assertTrue(report["ok"], report)
        self.assertTrue(report["byteIdentical"], report)
        self.assertEqual(sha256_file(input_docx), sha256_file(output_docx))

    def test_export_creates_manifest_and_html_projection(self) -> None:
        input_docx = TMP / "input.docx"
        work = TMP / "work.hdocx"
        _make_minimal_docx(input_docx)

        report = export_docx(input_docx, work)
        validation = validate_hdocx(work)

        self.assertTrue(report["ok"], report)
        self.assertTrue(validation["ok"], validation)
        self.assertTrue((work / "manifest.json").exists())
        html = (work / "document.html").read_text(encoding="utf-8")
        self.assertIn('data-hdocx-id="p-000001"', html)
        self.assertIn("Hello strict roundtrip.", html)

    def test_create_docx_exports_and_roundtrips(self) -> None:
        created_docx = TMP / "created.docx"
        created_docx_2 = TMP / "created-copy.docx"
        work = TMP / "created.hdocx"
        checked = TMP / "created-checked.docx"
        check_work = TMP / "created-check.hdocx"

        report = create_docx(
            created_docx,
            title="New H-DOCX Document",
            paragraphs=["First paragraph.", "Second paragraph."],
            export_dir=work,
            force=True,
        )
        repeat = create_docx(
            created_docx_2,
            title="New H-DOCX Document",
            paragraphs=["First paragraph.", "Second paragraph."],
            force=True,
        )
        validation = validate_hdocx(work)
        roundtrip = roundtrip_docx(created_docx, check_work, checked, force=True)

        self.assertTrue(report["ok"], report)
        self.assertEqual(report["command"], "create")
        self.assertEqual(report["paragraphCount"], 2)
        self.assertEqual(report["export"]["command"], "export")
        self.assertEqual(sha256_file(created_docx), sha256_file(created_docx_2))
        self.assertTrue(validation["ok"], validation)
        self.assertTrue(roundtrip["byteIdentical"], roundtrip)
        html = (work / "document.html").read_text(encoding="utf-8")
        self.assertIn("New H-DOCX Document", html)
        self.assertIn("First paragraph.", html)
        self.assertIn('data-hdocx-style-id="BodyText"', html)

    def test_complex_academic_fixture_unmodified_roundtrip_is_byte_identical(self) -> None:
        input_docx = TMP / "complex-academic.docx"
        work = TMP / "complex.hdocx"
        output_docx = TMP / "complex-roundtrip.docx"
        _make_complex_academic_docx(input_docx)

        export_report = export_docx(input_docx, work)
        validation = validate_hdocx(work)
        roundtrip = roundtrip_docx(input_docx, TMP / "complex-roundtrip.hdocx", output_docx, force=True)
        html = (work / "document.html").read_text(encoding="utf-8")

        self.assertTrue(export_report["ok"], export_report)
        self.assertTrue(validation["ok"], validation)
        self.assertTrue(roundtrip["byteIdentical"], roundtrip)
        self.assertIn('data-hdocx-style-id="BodyText"', html)
        self.assertIn('data-hdocx-num-format="decimal"', html)
        self.assertIn('data-hdocx-part="/word/header1.xml"', html)
        self.assertIn('data-hdocx-part="/word/footnotes.xml"', html)
        self.assertIn("[comment] Reviewer note.", html)
        self.assertIn("[revision-insert:7]", html)
        self.assertIn("[equation]", html)
        self.assertIn('data-hdocx-width-emu="914400"', html)

    def test_audit_detects_high_risk_and_protected_structures(self) -> None:
        input_docx = TMP / "audit-features.docx"
        _make_audit_feature_docx(input_docx)

        report = audit_docx(input_docx)

        self.assertTrue(report["ok"], report)
        features = report["features"]
        for name in (
            "customXml",
            "chart",
            "smartArt",
            "ole",
            "alternateContent",
            "vml",
            "textBox",
            "field",
            "equation",
            "revision",
            "comment",
            "footnote",
            "endnote",
            "headerFooter",
            "image",
        ):
            self.assertTrue(features[name]["present"], name)
            self.assertIn("policy", features[name], name)
        self.assertIn("chart", report["summary"]["highRiskPresent"])
        self.assertIn("equation", report["summary"]["protectedStructurePresent"])
        self.assertTrue(report["summary"]["hasNonEditableAdvancedObjects"])

    def test_apply_simple_run_text_patch(self) -> None:
        input_docx = TMP / "input.docx"
        work = TMP / "work.hdocx"
        output_docx = TMP / "output.docx"
        _make_minimal_docx(input_docx)
        with zipfile.ZipFile(input_docx, "r") as zf:
            original_document_xml = zf.read("word/document.xml")
        export_docx(input_docx, work)
        html_path = work / "document.html"
        html_path.write_text(
            html_path.read_text(encoding="utf-8").replace("Hello", "Changed"),
            encoding="utf-8",
            newline="\n",
        )

        plan = plan_hdocx(work)
        apply_report = apply_hdocx(work, output_docx)
        diff = diff_docx(input_docx, output_docx)

        self.assertTrue(plan["ok"], plan)
        self.assertEqual(len(plan["patches"]), 1, plan)
        self.assertTrue(apply_report["ok"], apply_report)
        self.assertEqual(apply_report["mode"], "document-patch")
        self.assertEqual(apply_report["patchSummary"]["count"], 1)
        self.assertEqual(apply_report["patchSummary"]["byEntry"], {"word/document.xml": 1})
        self.assertEqual(apply_report["packageDiff"]["changed"], ["word/document.xml"])
        self.assertEqual(diff["changed"], ["word/document.xml"], diff)
        self.assertFalse(diff["byteIdentical"], diff)
        self.assertEqual(diff["entryCounts"]["changed"], 1, diff)
        self.assertEqual(diff["changedEntries"][0]["kind"], "main-document", diff)
        self.assertIn("sha256", diff["changedEntries"][0]["left"], diff)
        self.assertIn("sha256", diff["changedEntries"][0]["right"], diff)
        self.assertIn("sha256", diff["changedEntries"][0]["changedFields"], diff)
        self.assertFalse(diff["semanticDiff"]["identical"], diff)
        self.assertIn("r-000001", diff["semanticDiff"]["changed"], diff)
        self.assertTrue(diff["fragmentDiff"]["available"], diff)
        self.assertEqual(diff["fragmentDiff"]["entryCount"], 1, diff)
        fragment_entry = diff["fragmentDiff"]["entries"][0]
        self.assertEqual(fragment_entry["path"], "word/document.xml")
        self.assertEqual(fragment_entry["status"], "changed")
        self.assertFalse(fragment_entry["byteDiff"]["identical"])
        self.assertGreater(fragment_entry["byteDiff"]["left"]["changedRange"]["length"], 0)
        self.assertGreater(fragment_entry["byteDiff"]["right"]["changedRange"]["length"], 0)
        self.assertIn("r-000001", [item["nodeId"] for item in fragment_entry["linkedNodes"]])
        run_change = next(
            item for item in diff["semanticDiff"]["changedNodes"] if item["nodeId"] == "r-000001"
        )
        self.assertEqual(run_change["changes"]["text"]["left"], "Hello strict roundtrip.")
        self.assertEqual(run_change["changes"]["text"]["right"], "Changed strict roundtrip.")
        with zipfile.ZipFile(output_docx, "r") as zf:
            document_xml_bytes = zf.read("word/document.xml")
        self.assertEqual(
            document_xml_bytes,
            original_document_xml.replace(b"Hello strict roundtrip.", b"Changed strict roundtrip."),
        )
        document_xml = document_xml_bytes.decode("utf-8")
        self.assertIn("Changed strict roundtrip.", document_xml)
        self.assertNotIn("Hello strict roundtrip.", document_xml)

    def test_apply_simple_run_format_patch(self) -> None:
        input_docx = TMP / "input.docx"
        work = TMP / "work.hdocx"
        output_docx = TMP / "output.docx"
        _make_minimal_docx(input_docx)
        with zipfile.ZipFile(input_docx, "r") as zf:
            original_document_xml = zf.read("word/document.xml")
        export_docx(input_docx, work)
        html_path = work / "document.html"
        html_path.write_text(
            html_path.read_text(encoding="utf-8").replace(
                'data-hdocx-id="r-000001" data-hdocx-lock="editable" data-hdocx-part="/word/document.xml">',
                'data-hdocx-id="r-000001" data-hdocx-lock="editable" data-hdocx-part="/word/document.xml" '
                'data-hdocx-bold="true" data-hdocx-font-size="14pt" data-hdocx-color="#ff0000">',
            ),
            encoding="utf-8",
            newline="\n",
        )

        plan = plan_hdocx(work)
        apply_report = apply_hdocx(work, output_docx)
        diff = diff_docx(input_docx, output_docx)

        self.assertTrue(plan["ok"], plan)
        self.assertEqual(plan["patches"][0]["newProperties"]["bold"], "true")
        self.assertEqual(plan["patches"][0]["newProperties"]["font-size"], "14pt")
        self.assertTrue(apply_report["ok"], apply_report)
        self.assertEqual(diff["changed"], ["word/document.xml"], diff)
        with zipfile.ZipFile(output_docx, "r") as zf:
            document_xml_bytes = zf.read("word/document.xml")
        self.assertEqual(
            document_xml_bytes,
            original_document_xml.replace(
                b"<w:r>",
                b'<w:r><w:rPr><w:b/><w:sz w:val="28"/><w:color w:val="FF0000"/></w:rPr>',
                1,
            ),
        )
        document_xml = document_xml_bytes.decode("utf-8")
        self.assertIn("<w:b", document_xml)
        self.assertIn('w:val="28"', document_xml)
        self.assertIn('w:val="FF0000"', document_xml)

    def test_apply_run_split_with_segment_formatting(self) -> None:
        input_docx = TMP / "input.docx"
        work = TMP / "work.hdocx"
        output_docx = TMP / "output.docx"
        _make_minimal_docx(input_docx)
        export_docx(input_docx, work)
        html_path = work / "document.html"
        html_path.write_text(
            html_path.read_text(encoding="utf-8").replace(
                "Hello strict roundtrip.</span>",
                'Hello <span data-hdocx-type="run-segment" data-hdocx-bold="true">strict</span> roundtrip.</span>',
            ),
            encoding="utf-8",
            newline="\n",
        )

        plan = plan_hdocx(work)
        apply_report = apply_hdocx(work, output_docx)

        self.assertTrue(plan["ok"], plan)
        self.assertEqual(plan["patches"][0]["operation"], "split-run")
        self.assertEqual(len(plan["patches"][0]["splitSegments"]), 3)
        self.assertTrue(apply_report["ok"], apply_report)
        with zipfile.ZipFile(output_docx, "r") as zf:
            root = ET.fromstring(zf.read("word/document.xml"))
        runs = root.findall(".//w:p", NS)[0].findall("w:r", NS)
        texts = ["".join(t.text or "" for t in run.findall("w:t", NS)) for run in runs]
        self.assertEqual(texts, ["Hello ", "strict", " roundtrip."])
        self.assertIsNotNone(runs[1].find("w:rPr/w:b", NS))

    def test_apply_paragraph_format_patch(self) -> None:
        input_docx = TMP / "input.docx"
        work = TMP / "work.hdocx"
        output_docx = TMP / "output.docx"
        _make_minimal_docx(input_docx)
        export_docx(input_docx, work)
        html_path = work / "document.html"
        html_path.write_text(
            html_path.read_text(encoding="utf-8").replace(
                'data-hdocx-id="p-000001" data-hdocx-lock="editable" data-hdocx-part="/word/document.xml">',
                'data-hdocx-id="p-000001" data-hdocx-lock="editable" data-hdocx-part="/word/document.xml" '
                'data-hdocx-align="center" data-hdocx-line-spacing="1.5" '
                'data-hdocx-first-line-indent="2char">',
            ),
            encoding="utf-8",
            newline="\n",
        )

        plan = plan_hdocx(work)
        apply_report = apply_hdocx(work, output_docx)

        self.assertTrue(plan["ok"], plan)
        self.assertEqual(plan["patches"][0]["operation"], "patch-paragraph")
        self.assertTrue(apply_report["ok"], apply_report)
        with zipfile.ZipFile(output_docx, "r") as zf:
            root = ET.fromstring(zf.read("word/document.xml"))
        p_pr = root.findall(".//w:p", NS)[0].find("w:pPr", NS)
        self.assertIsNotNone(p_pr)
        self.assertEqual(p_pr.find("w:jc", NS).attrib[f"{{{W_NS}}}val"], "center")
        self.assertEqual(p_pr.find("w:spacing", NS).attrib[f"{{{W_NS}}}line"], "360")
        self.assertEqual(p_pr.find("w:spacing", NS).attrib[f"{{{W_NS}}}lineRule"], "auto")
        self.assertEqual(p_pr.find("w:ind", NS).attrib[f"{{{W_NS}}}firstLineChars"], "200")

    def test_paragraph_format_fragment_patch_preserves_unrelated_ppr_attrs(self) -> None:
        input_docx = TMP / "paragraph-props.docx"
        work = TMP / "work.hdocx"
        output_docx = TMP / "output.docx"
        _make_paragraph_property_docx(input_docx)
        export_docx(input_docx, work)
        (work / "agent.edits.hcss").write_text(
            """
@hdocx-edit mode(paragraph-formatting);

#p-000001 {
  hdocx-align: center;
  hdocx-first-line-indent: 2char;
  hdocx-line-spacing: 1.5;
}
""",
            encoding="utf-8",
            newline="\n",
        )

        plan = plan_hdocx(work)
        apply_report = apply_hdocx(work, output_docx)

        self.assertTrue(plan["ok"], plan)
        self.assertEqual(plan["patches"][0]["operation"], "patch-paragraph")
        with zipfile.ZipFile(output_docx, "r") as zf:
            document_xml = zf.read("word/document.xml").decode("utf-8")
        self.assertIn("<w:keepNext/>", document_xml)
        self.assertIn('w:ind w:left="720" w:firstLineChars="200"', document_xml)
        self.assertNotIn('w:firstLine="240"', document_xml)
        self.assertIn('w:spacing w:before="120" w:line="360" w:lineRule="auto"', document_xml)
        self.assertIn("<w:t>Paragraph with preserved pPr.</w:t>", document_xml)

    def test_hcss_paragraph_formatting(self) -> None:
        input_docx = TMP / "input.docx"
        work = TMP / "work.hdocx"
        output_docx = TMP / "output.docx"
        _make_minimal_docx(input_docx)
        export_docx(input_docx, work)
        (work / "agent.edits.hcss").write_text(
            """
@hdocx-set body {
  select: [data-hdocx-type="paragraph"];
}

@hdocx-edit mode(paragraph-formatting);

body {
  hdocx-align: center;
  hdocx-line-spacing: 1.5;
}
""",
            encoding="utf-8",
            newline="\n",
        )

        plan = plan_hdocx(work)
        apply_report = apply_hdocx(work, output_docx)

        self.assertTrue(plan["ok"], plan)
        self.assertEqual(plan["patches"][0]["operation"], "patch-paragraph")
        with zipfile.ZipFile(output_docx, "r") as zf:
            root = ET.fromstring(zf.read("word/document.xml"))
        p_pr = root.findall(".//w:p", NS)[0].find("w:pPr", NS)
        self.assertEqual(p_pr.find("w:jc", NS).attrib[f"{{{W_NS}}}val"], "center")
        self.assertEqual(p_pr.find("w:spacing", NS).attrib[f"{{{W_NS}}}line"], "360")

    def test_hcss_set_id_function_alias_formats_paragraph(self) -> None:
        input_docx = TMP / "input.docx"
        work = TMP / "work.hdocx"
        output_docx = TMP / "output.docx"
        _make_minimal_docx(input_docx)
        export_docx(input_docx, work)
        (work / "agent.edits.hcss").write_text(
            """
@hdocx-set title {
  select: id(p-000001);
}

@hdocx-edit mode(paragraph-formatting);

title {
  hdocx-line-spacing-exact: 18pt;
}
""",
            encoding="utf-8",
            newline="\n",
        )

        plan = plan_hdocx(work)
        apply_hdocx(work, output_docx)

        self.assertTrue(plan["ok"], plan)
        self.assertEqual(plan["hcss"]["rules"][0]["matchedNodeIds"], ["p-000001"])
        self.assertEqual(plan["patches"][0]["operation"], "patch-paragraph")
        with zipfile.ZipFile(output_docx, "r") as zf:
            root = ET.fromstring(zf.read("word/document.xml"))
        spacing = root.findall(".//w:p", NS)[0].find("w:pPr/w:spacing", NS)
        self.assertEqual(spacing.attrib[f"{{{W_NS}}}line"], "360")
        self.assertEqual(spacing.attrib[f"{{{W_NS}}}lineRule"], "exact")

    def test_hcss_set_selector_list_targets_multiple_paragraphs(self) -> None:
        input_docx = TMP / "created.docx"
        work = TMP / "work.hdocx"
        create_docx(input_docx, title="Title", paragraphs=["One.", "Two."], force=True)
        export_docx(input_docx, work)
        (work / "agent.edits.hcss").write_text(
            """
@hdocx-set body {
  select: id(p-000002), id(p-000003);
}

@hdocx-edit mode(paragraph-formatting);

body {
  hdocx-text-align: justify;
}
""",
            encoding="utf-8",
            newline="\n",
        )

        plan = plan_hdocx(work)

        self.assertTrue(plan["ok"], plan)
        self.assertEqual(plan["hcss"]["rules"][0]["matchedNodeIds"], ["p-000002", "p-000003"])
        self.assertEqual([patch["nodeId"] for patch in plan["patches"]], ["p-000002", "p-000003"])

    def test_hcss_comma_grouping_rule_targets_multiple_paragraphs(self) -> None:
        input_docx = TMP / "created.docx"
        work = TMP / "work.hdocx"
        create_docx(input_docx, title="Title", paragraphs=["One.", "Two."], force=True)
        export_docx(input_docx, work)
        (work / "agent.edits.hcss").write_text(
            """
@hdocx-edit mode(paragraph-formatting);

.hdocx-p[data-hdocx-id="p-000002"],
.hdocx-p[data-hdocx-id="p-000003"] {
  hdocx-text-align: center;
}
""",
            encoding="utf-8",
            newline="\n",
        )

        plan = plan_hdocx(work)

        self.assertTrue(plan["ok"], plan)
        self.assertEqual(plan["hcss"]["rules"][0]["matchedNodeIds"], ["p-000002", "p-000003"])
        self.assertEqual(len(plan["patches"]), 2)

    def test_hcss_all_runs_with_token_format_include(self) -> None:
        input_docx = TMP / "input.docx"
        work = TMP / "work.hdocx"
        output_docx = TMP / "output.docx"
        _make_minimal_docx(input_docx)
        export_docx(input_docx, work)
        (work / "agent.edits.hcss").write_text(
            """
@hdocx-token size.body 14pt;

@hdocx-set body {
  select: [data-hdocx-type="paragraph"];
}

@hdocx-format body-run {
  hdocx-font-size: token(size.body);
  hdocx-bold: true;
}

@hdocx-edit mode(all-runs);

body {
  @hdocx-include body-run;
}
""",
            encoding="utf-8",
            newline="\n",
        )

        plan = plan_hdocx(work)
        apply_report = apply_hdocx(work, output_docx)

        self.assertTrue(plan["ok"], plan)
        self.assertEqual(plan["patches"][0]["operation"], "patch-run")
        with zipfile.ZipFile(output_docx, "r") as zf:
            root = ET.fromstring(zf.read("word/document.xml"))
        r_pr = root.findall(".//w:r", NS)[0].find("w:rPr", NS)
        self.assertIsNotNone(r_pr.find("w:b", NS))
        self.assertEqual(r_pr.find("w:sz", NS).attrib[f"{{{W_NS}}}val"], "28")

    def test_hcss_paper_format_contract_maps_to_ooxml(self) -> None:
        input_docx = TMP / "paper.docx"
        work = TMP / "paper.hdocx"
        output_docx = TMP / "paper-output.docx"
        create_docx(input_docx, title="Paper Title", paragraphs=["Body paragraph."], force=True)
        export_docx(input_docx, work)
        (work / "agent.edits.hcss").write_text(
            """
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
""",
            encoding="utf-8",
            newline="\n",
        )

        plan = plan_hdocx(work)
        apply_hdocx(work, output_docx)

        self.assertTrue(plan["ok"], plan)
        self.assertEqual(plan["hcss"]["summary"]["ruleCount"], 2)
        self.assertEqual(plan["hcss"]["summary"]["unsupportedDeclarationCount"], 0)
        self.assertEqual(plan["hcss"]["rules"][0]["matchedNodeIds"], ["p-000002"])
        self.assertEqual(plan["hcss"]["rules"][1]["patchIds"], ["patch-000002"])
        declarations = plan["hcss"]["rules"][0]["declarations"]
        self.assertEqual(declarations[0]["normalizedProperty"], "align")
        self.assertEqual(declarations[0]["ooxml"], "w:pPr/w:jc @w:val")
        with zipfile.ZipFile(output_docx, "r") as zf:
            root = ET.fromstring(zf.read("word/document.xml"))
        body_paragraph = root.findall(".//w:p", NS)[1]
        p_pr = body_paragraph.find("w:pPr", NS)
        spacing = p_pr.find("w:spacing", NS)
        ind = p_pr.find("w:ind", NS)
        self.assertEqual(p_pr.find("w:jc", NS).attrib[f"{{{W_NS}}}val"], "both")
        self.assertEqual(ind.attrib[f"{{{W_NS}}}firstLineChars"], "200")
        self.assertEqual(spacing.attrib[f"{{{W_NS}}}line"], "360")
        self.assertEqual(spacing.attrib[f"{{{W_NS}}}lineRule"], "exact")
        self.assertEqual(spacing.attrib[f"{{{W_NS}}}before"], "0")
        self.assertEqual(spacing.attrib[f"{{{W_NS}}}after"], "0")
        r_pr = body_paragraph.find("w:r/w:rPr", NS)
        r_fonts = r_pr.find("w:rFonts", NS)
        self.assertEqual(r_fonts.attrib[f"{{{W_NS}}}ascii"], "Times New Roman")
        self.assertEqual(r_fonts.attrib[f"{{{W_NS}}}hAnsi"], "Times New Roman")
        self.assertEqual(r_fonts.attrib[f"{{{W_NS}}}eastAsia"], "SimSun")
        self.assertEqual(r_pr.find("w:sz", NS).attrib[f"{{{W_NS}}}val"], "21")

    def test_hcss_plan_reports_unsupported_declarations_with_lines(self) -> None:
        input_docx = TMP / "input.docx"
        work = TMP / "work.hdocx"
        _make_minimal_docx(input_docx)
        export_docx(input_docx, work)
        (work / "agent.edits.hcss").write_text(
            """
@hdocx-set body {
  select: [data-hdocx-type="paragraph"];
}

@hdocx-edit mode(paragraph-formatting);

body {
  font-family: Arial;
  hdocx-font-size: 12pt;
}
""",
            encoding="utf-8",
            newline="\n",
        )

        plan = plan_hdocx(work)

        self.assertFalse(plan["ok"], plan)
        self.assertEqual(plan["hcss"]["summary"]["unsupportedDeclarationCount"], 2)
        declarations = plan["hcss"]["rules"][0]["declarations"]
        self.assertEqual(declarations[0]["property"], "font-family")
        self.assertEqual(declarations[0]["reason"], "Only hdocx-* declarations are accepted.")
        self.assertIsInstance(declarations[0]["line"], int)
        self.assertEqual(declarations[1]["normalizedProperty"], "font-size")
        self.assertEqual(declarations[1]["reason"], "Property is supported, but not in paragraph-formatting mode.")
        self.assertIsInstance(plan["errors"][0]["line"], int)

    def test_hcss_run_compound_attribute_selector_with_bom(self) -> None:
        input_docx = TMP / "input.docx"
        work = TMP / "work.hdocx"
        output_docx = TMP / "output.docx"
        _make_minimal_docx(input_docx)
        export_docx(input_docx, work)
        hcss = """\ufeff@hdocx-edit mode(all-runs);

.hdocx-r[data-hdocx-id="r-000001"] {
  hdocx-font-family: "Times New Roman";
  hdocx-eastAsia-font: "SimSun";
  hdocx-font-size: 10.5pt;
}
"""
        (work / "agent.edits.hcss").write_bytes(hcss.encode("utf-8"))

        plan = plan_hdocx(work)
        apply_hdocx(work, output_docx)

        self.assertTrue(plan["ok"], plan)
        self.assertEqual(plan["hcss"]["rules"][0]["matchedNodeIds"], ["r-000001"])
        self.assertEqual(plan["hcss"]["rules"][0]["patchIds"], ["patch-000001"])
        self.assertEqual(plan["hcss"]["summary"]["unsupportedDeclarationCount"], 0)
        with zipfile.ZipFile(output_docx, "r") as zf:
            root = ET.fromstring(zf.read("word/document.xml"))
        r_pr = root.findall(".//w:r", NS)[0].find("w:rPr", NS)
        r_fonts = r_pr.find("w:rFonts", NS)
        self.assertEqual(r_fonts.attrib[f"{{{W_NS}}}ascii"], "Times New Roman")
        self.assertEqual(r_fonts.attrib[f"{{{W_NS}}}hAnsi"], "Times New Roman")
        self.assertEqual(r_fonts.attrib[f"{{{W_NS}}}eastAsia"], "SimSun")
        self.assertEqual(r_pr.find("w:sz", NS).attrib[f"{{{W_NS}}}val"], "21")

    def test_hcss_set_alias_can_target_run_compound_selector(self) -> None:
        input_docx = TMP / "input.docx"
        work = TMP / "work.hdocx"
        _make_minimal_docx(input_docx)
        export_docx(input_docx, work)
        (work / "agent.edits.hcss").write_text(
            """
@hdocx-set target-run {
  select: .hdocx-r[data-hdocx-id="r-000001"];
}

@hdocx-edit mode(all-runs);

target-run {
  hdocx-bold: true;
}
""",
            encoding="utf-8",
            newline="\n",
        )

        plan = plan_hdocx(work)

        self.assertTrue(plan["ok"], plan)
        self.assertEqual(plan["hcss"]["rules"][0]["target"], "target-run")
        self.assertEqual(plan["hcss"]["rules"][0]["matchedNodeIds"], ["r-000001"])
        self.assertEqual(plan["patches"][0]["nodeId"], "r-000001")

    def test_hcss_editable_run_selector_excludes_protected_nodes(self) -> None:
        input_docx = TMP / "complex.docx"
        work = TMP / "work.hdocx"
        _make_complex_academic_docx(input_docx)
        export_docx(input_docx, work)
        (work / "agent.edits.hcss").write_text(
            """
@hdocx-set editable-runs {
  select: .hdocx-r[data-hdocx-lock="editable"];
}

@hdocx-edit mode(all-runs);

editable-runs {
  hdocx-font-size: 10.5pt;
}
""",
            encoding="utf-8",
            newline="\n",
        )

        plan = plan_hdocx(work)

        self.assertTrue(plan["ok"], plan)
        self.assertGreater(plan["hcss"]["rules"][0]["matchCount"], 0)
        self.assertTrue(all(node_id.startswith("r-") for node_id in plan["hcss"]["rules"][0]["matchedNodeIds"]))

    def test_hcss_manual_page_break_before_paragraph(self) -> None:
        input_docx = TMP / "input.docx"
        work = TMP / "work.hdocx"
        output_docx = TMP / "output.docx"
        _make_minimal_docx(input_docx)
        export_docx(input_docx, work)
        (work / "agent.edits.hcss").write_text(
            """
@hdocx-edit mode(paragraph-formatting);

.hdocx-p[data-hdocx-id="p-000001"] {
  hdocx-manual-page-break-before: true;
}
""",
            encoding="utf-8",
            newline="\n",
        )

        plan = plan_hdocx(work)
        apply_hdocx(work, output_docx)
        diff = diff_docx(input_docx, output_docx)

        self.assertTrue(plan["ok"], plan)
        self.assertEqual(plan["patches"][0]["operation"], "insert-manual-page-break-before")
        self.assertEqual(plan["hcss"]["rules"][0]["patchIds"], ["patch-000001"])
        self.assertEqual(plan["hcss"]["rules"][0]["declarations"][0]["kind"], "structural-paragraph")
        with zipfile.ZipFile(output_docx, "r") as zf:
            root = ET.fromstring(zf.read("word/document.xml"))
        paragraphs = root.findall(".//w:p", NS)
        self.assertEqual(paragraphs[0].find("w:r/w:br", NS).attrib[f"{{{W_NS}}}type"], "page")
        self.assertEqual(paragraphs[1].find("w:r/w:t", NS).text, "Hello strict roundtrip.")
        self.assertEqual(diff["manualPageBreakDiff"]["summary"]["added"], 1)
        self.assertEqual(diff["manualPageBreakDiff"]["added"][0]["kind"], "manual-page-break")
        self.assertTrue(diff["semanticDiff"]["identical"], diff["semanticDiff"])
        self.assertEqual(diff["semanticDiff"]["manualPageBreakAlignment"]["ignoredRightBreakParagraphs"], 1)

    def test_hcss_manual_page_break_before_is_idempotent_when_existing(self) -> None:
        input_docx = TMP / "input.docx"
        first_work = TMP / "first.hdocx"
        first_output = TMP / "first.docx"
        second_work = TMP / "second.hdocx"
        second_output = TMP / "second.docx"
        _make_minimal_docx(input_docx)
        export_docx(input_docx, first_work)
        (first_work / "agent.edits.hcss").write_text(
            """
@hdocx-edit mode(paragraph-formatting);

#p-000001 {
  hdocx-manual-page-break-before: true;
}
""",
            encoding="utf-8",
            newline="\n",
        )
        apply_hdocx(first_work, first_output)
        export_docx(first_output, second_work)
        (second_work / "agent.edits.hcss").write_text(
            """
@hdocx-edit mode(paragraph-formatting);

#p-000002 {
  hdocx-manual-page-break-before: true;
}
""",
            encoding="utf-8",
            newline="\n",
        )

        apply_hdocx(second_work, second_output)

        with zipfile.ZipFile(second_output, "r") as zf:
            root = ET.fromstring(zf.read("word/document.xml"))
        break_count = sum(
            1
            for br in root.findall(".//w:br", NS)
            if br.attrib.get(f"{{{W_NS}}}type") == "page"
        )
        self.assertEqual(break_count, 1)

    def test_hcss_insert_empty_paragraph_after_with_spacing(self) -> None:
        input_docx = TMP / "input.docx"
        work = TMP / "work.hdocx"
        output_docx = TMP / "output.docx"
        _make_minimal_docx(input_docx)
        export_docx(input_docx, work)
        (work / "agent.edits.hcss").write_text(
            """
@hdocx-edit mode(paragraph-structure);

#p-000001 {
  hdocx-insert-empty-paragraph-after: true;
  hdocx-empty-paragraph-line-spacing-exact: 12pt;
  hdocx-empty-paragraph-space-before: 0;
  hdocx-empty-paragraph-space-after: 0;
}
""",
            encoding="utf-8",
            newline="\n",
        )

        plan = plan_hdocx(work)
        apply_hdocx(work, output_docx)
        diff = diff_docx(input_docx, output_docx)

        self.assertTrue(plan["ok"], plan)
        self.assertEqual(plan["patches"][0]["operation"], "insert-empty-paragraph-after")
        self.assertEqual(plan["patches"][0]["newProperties"]["line-spacing"], "12pt")
        self.assertEqual(plan["hcss"]["rules"][0]["declarations"][0]["kind"], "structural-paragraph")
        with zipfile.ZipFile(output_docx, "r") as zf:
            root = ET.fromstring(zf.read("word/document.xml"))
        paragraphs = root.findall(".//w:p", NS)
        self.assertEqual(len(paragraphs), 2)
        self.assertEqual(paragraphs[0].find("w:r/w:t", NS).text, "Hello strict roundtrip.")
        self.assertEqual(paragraphs[1].findall(".//w:t", NS), [])
        spacing = paragraphs[1].find("w:pPr/w:spacing", NS)
        self.assertEqual(spacing.attrib[f"{{{W_NS}}}line"], "240")
        self.assertEqual(spacing.attrib[f"{{{W_NS}}}lineRule"], "exact")
        self.assertEqual(spacing.attrib[f"{{{W_NS}}}before"], "0")
        self.assertEqual(spacing.attrib[f"{{{W_NS}}}after"], "0")
        self.assertEqual(diff["emptyParagraphDiff"]["summary"]["added"], 1)
        self.assertEqual(diff["emptyParagraphDiff"]["added"][0]["kind"], "empty-paragraph")
        self.assertTrue(diff["semanticDiff"]["identical"], diff["semanticDiff"])
        self.assertEqual(diff["semanticDiff"]["emptyParagraphAlignment"]["ignoredRightEmptyParagraphs"], 1)

    def test_hcss_insert_empty_paragraph_after_is_idempotent_when_existing(self) -> None:
        input_docx = TMP / "input.docx"
        first_work = TMP / "first.hdocx"
        first_output = TMP / "first.docx"
        second_work = TMP / "second.hdocx"
        second_output = TMP / "second.docx"
        _make_minimal_docx(input_docx)
        export_docx(input_docx, first_work)
        (first_work / "agent.edits.hcss").write_text(
            """
@hdocx-edit mode(paragraph-structure);

#p-000001 {
  hdocx-insert-empty-paragraph-after: true;
}
""",
            encoding="utf-8",
            newline="\n",
        )
        apply_hdocx(first_work, first_output)
        export_docx(first_output, second_work)
        (second_work / "agent.edits.hcss").write_text(
            """
@hdocx-edit mode(paragraph-structure);

#p-000001 {
  hdocx-insert-empty-paragraph-after: true;
  hdocx-empty-paragraph-line-spacing-exact: 12pt;
  hdocx-empty-paragraph-space-before: 0;
  hdocx-empty-paragraph-space-after: 0;
}
""",
            encoding="utf-8",
            newline="\n",
        )

        apply_hdocx(second_work, second_output)
        diff = diff_docx(first_output, second_output)

        with zipfile.ZipFile(second_output, "r") as zf:
            root = ET.fromstring(zf.read("word/document.xml"))
        empty_paragraphs = [paragraph for paragraph in root.findall(".//w:p", NS) if not paragraph.findall(".//w:t", NS)]
        empty_count = len(empty_paragraphs)
        self.assertEqual(empty_count, 1)
        spacing = empty_paragraphs[0].find("w:pPr/w:spacing", NS)
        self.assertEqual(spacing.attrib[f"{{{W_NS}}}line"], "240")
        self.assertEqual(spacing.attrib[f"{{{W_NS}}}lineRule"], "exact")
        self.assertEqual(diff["emptyParagraphDiff"]["summary"]["added"], 0)
        self.assertEqual(diff["emptyParagraphDiff"]["summary"]["changed"], 1)

    def test_hcss_insert_empty_paragraph_style_id_applies(self) -> None:
        input_docx = TMP / "input.docx"
        work = TMP / "work.hdocx"
        output_docx = TMP / "output.docx"
        _make_minimal_docx(input_docx)
        export_docx(input_docx, work)
        (work / "agent.edits.hcss").write_text(
            """
@hdocx-edit mode(paragraph-structure);

#p-000001 {
  hdocx-insert-empty-paragraph-after: true;
  hdocx-empty-paragraph-style-id: Normal;
}
""",
            encoding="utf-8",
            newline="\n",
        )

        plan = plan_hdocx(work)
        apply_hdocx(work, output_docx)

        self.assertTrue(plan["ok"], plan)
        self.assertEqual(plan["patches"][0]["newProperties"]["style-id"], "Normal")
        with zipfile.ZipFile(output_docx, "r") as zf:
            root = ET.fromstring(zf.read("word/document.xml"))
        p_style = root.findall(".//w:p", NS)[1].find("w:pPr/w:pStyle", NS)
        self.assertEqual(p_style.attrib[f"{{{W_NS}}}val"], "Normal")

    def test_hcss_insert_empty_paragraph_plan_noop_when_existing_matches(self) -> None:
        input_docx = TMP / "input.docx"
        first_work = TMP / "first.hdocx"
        first_output = TMP / "first.docx"
        second_work = TMP / "second.hdocx"
        _make_minimal_docx(input_docx)
        export_docx(input_docx, first_work)
        (first_work / "agent.edits.hcss").write_text(
            """
@hdocx-edit mode(paragraph-structure);

#p-000001 {
  hdocx-insert-empty-paragraph-after: true;
}
""",
            encoding="utf-8",
            newline="\n",
        )
        apply_hdocx(first_work, first_output)
        export_docx(first_output, second_work)
        (second_work / "agent.edits.hcss").write_text(
            """
@hdocx-edit mode(paragraph-structure);

#p-000001 {
  hdocx-insert-empty-paragraph-after: true;
}
""",
            encoding="utf-8",
            newline="\n",
        )

        plan = plan_hdocx(second_work)

        self.assertTrue(plan["ok"], plan)
        self.assertEqual(plan["patches"], [])
        self.assertEqual(plan["hcss"]["summary"]["patchCount"], 0)
        self.assertEqual(plan["hcss"]["summary"]["noopCount"], 1)
        self.assertEqual(plan["hcss"]["rules"][0]["noops"][0]["reason"], "idempotent-empty-paragraph-already-present")

    def test_empty_paragraph_alignment_ignores_table_cell_child_id_shift(self) -> None:
        input_docx = TMP / "table.docx"
        work = TMP / "work.hdocx"
        output_docx = TMP / "output.docx"
        _make_table_docx(input_docx)
        export_docx(input_docx, work)
        (work / "agent.edits.hcss").write_text(
            """
@hdocx-edit mode(paragraph-structure);

#p-000001 {
  hdocx-insert-empty-paragraph-after: true;
}
""",
            encoding="utf-8",
            newline="\n",
        )

        apply_hdocx(work, output_docx)
        diff = diff_docx(input_docx, output_docx)

        self.assertEqual(diff["emptyParagraphDiff"]["summary"]["added"], 1)
        self.assertTrue(diff["semanticDiff"]["identical"], diff["semanticDiff"])
        self.assertEqual(diff["semanticDiff"]["changedNodes"], [])

    def test_query_finds_text_formatting_headings_and_images(self) -> None:
        input_docx = TMP / "query.docx"
        work = TMP / "query.hdocx"
        _make_query_docx(input_docx)
        export_docx(input_docx, work)

        keyword = query_hdocx(work, text="Keywords")
        self.assertTrue(keyword["ok"], keyword)
        self.assertEqual(keyword["matches"][0]["nodeId"], "p-000003")

        heading = query_hdocx(
            work,
            align="center",
            font_size="14pt",
            font_family="SimHei",
            suspected_heading_level1=True,
        )
        self.assertTrue(heading["ok"], heading)
        self.assertEqual([match["nodeId"] for match in heading["matches"]], ["p-000002"])

        images = find_hdocx(work, kind="image")
        self.assertTrue(images["ok"], images)
        self.assertEqual(images["matches"][0]["nodeId"], "r-000003")
        self.assertEqual(images["matches"][0]["hostParagraph"]["nodeId"], "p-000004")

    def test_assertion_checks_text_empty_paragraphs_and_image_spacing(self) -> None:
        input_docx = TMP / "assert.docx"
        work = TMP / "assert.hdocx"
        _make_query_docx(input_docx)
        export_docx(input_docx, work)

        report = assert_hdocx(
            work,
            [
                "text-payload-unchanged",
                {"type": "paragraphs-have-empty-before", "paragraphIds": ["p-000002"]},
                "images-host-paragraph-not-exact-line-spacing",
                {"type": "level1-headings-have-empty-paragraph-before", "minScore": 3},
            ],
        )

        self.assertTrue(report["ok"], report)
        self.assertEqual(report["passedCount"], 4)

    def test_assertion_fails_for_exact_image_host_spacing(self) -> None:
        input_docx = TMP / "assert-image-exact.docx"
        work = TMP / "assert-image-exact.hdocx"
        _make_image_docx(input_docx, exact_line_spacing=True)
        export_docx(input_docx, work)

        report = assert_hdocx(work, ["images-host-paragraph-not-exact-line-spacing"])

        self.assertFalse(report["ok"], report)
        failure = report["assertions"][0]["failures"][0]
        self.assertEqual(failure["imageRunId"], "r-000001")
        self.assertEqual(failure["lineSpacingRule"]["lineRule"], "exact")

    def test_level1_heading_assertion_excludes_preface_titles_by_default(self) -> None:
        input_docx = TMP / "preface-heading.docx"
        work = TMP / "preface-heading.hdocx"
        _make_preface_heading_docx(input_docx)
        export_docx(input_docx, work)

        report = assert_hdocx(work, [{"type": "level1-headings-have-empty-paragraph-before", "minScore": 3}])

        self.assertTrue(report["ok"], report)
        assertion = report["assertions"][0]
        self.assertEqual([heading["nodeId"] for heading in assertion["headings"]], ["p-000004"])
        skipped_by_default = [
            item for item in assertion["skippedHeadings"] if item["reason"] == "default-exclude-regex-match"
        ]
        self.assertEqual([item["nodeId"] for item in skipped_by_default], ["p-000002"])

    def test_level1_heading_assertion_supports_include_and_exclude_regex(self) -> None:
        input_docx = TMP / "preface-heading-regex.docx"
        work = TMP / "preface-heading-regex.hdocx"
        _make_preface_heading_docx(input_docx)
        export_docx(input_docx, work)

        included = assert_hdocx(
            work,
            [{"type": "level1-headings-have-empty-paragraph-before", "includeRegex": "Introduction"}],
        )
        excluded = assert_hdocx(
            work,
            [{"type": "level1-headings-have-empty-paragraph-before", "excludeRegex": "Introduction"}],
        )

        self.assertTrue(included["ok"], included)
        self.assertEqual([heading["nodeId"] for heading in included["assertions"][0]["headings"]], ["p-000004"])
        self.assertTrue(excluded["ok"], excluded)
        self.assertEqual(excluded["assertions"][0]["checkedHeadingCount"], 0)
        self.assertIn(
            "exclude-regex-match",
            [item["reason"] for item in excluded["assertions"][0]["skippedHeadings"]],
        )

    def test_assertion_after_apply_checks_planned_output_state(self) -> None:
        input_docx = TMP / "after-apply.docx"
        work = TMP / "after-apply.hdocx"
        _make_minimal_docx(input_docx)
        export_docx(input_docx, work)
        (work / "agent.edits.hcss").write_text(
            """
@hdocx-edit mode(paragraph-structure);

#p-000001 {
  hdocx-insert-empty-paragraph-before: true;
}
""",
            encoding="utf-8",
            newline="\n",
        )

        current = assert_hdocx(work, [{"type": "paragraphs-have-empty-before", "paragraphIds": ["p-000001"]}])
        planned = assert_hdocx(
            work,
            [{"type": "paragraphs-have-empty-before", "paragraphIds": ["p-000001"], "afterApply": True}],
        )

        self.assertFalse(current["ok"], current)
        self.assertTrue(planned["ok"], planned)
        assertion = planned["assertions"][0]
        self.assertEqual(assertion["scope"], "planned-output")
        self.assertEqual(assertion["nodeIdTranslations"], {"p-000001": "p-000002"})

    def test_chinese_and_special_filename_path_roundtrips(self) -> None:
        special_dir = TMP / "中文路径（测试）"
        special_dir.mkdir()
        input_docx = special_dir / "课程论文_人工智能高速发展对人类社会的影响3(1).docx"
        work = special_dir / "工作副本.hdocx"
        output_docx = special_dir / "输出（应用）.docx"
        _make_minimal_docx(input_docx)

        audit = audit_docx(input_docx)
        export_report = export_docx(input_docx, work)
        apply_report = apply_hdocx(work, output_docx)
        diff = diff_docx(input_docx, output_docx)

        self.assertTrue(audit["ok"], audit)
        self.assertTrue(export_report["ok"], export_report)
        self.assertEqual(apply_report["mode"], "copy-original-if-unmodified")
        self.assertTrue(diff["byteIdentical"], diff)

    def test_hcss_class_selector_formats_paragraphs(self) -> None:
        input_docx = TMP / "input.docx"
        work = TMP / "work.hdocx"
        output_docx = TMP / "output.docx"
        _make_minimal_docx(input_docx)
        export_docx(input_docx, work)
        (work / "agent.edits.hcss").write_text(
            """
@hdocx-edit mode(paragraph-formatting);

.hdocx-p {
  hdocx-align: right;
}
""",
            encoding="utf-8",
            newline="\n",
        )

        plan = plan_hdocx(work)
        apply_hdocx(work, output_docx)

        self.assertTrue(plan["ok"], plan)
        self.assertEqual(plan["patches"][0]["operation"], "patch-paragraph")
        with zipfile.ZipFile(output_docx, "r") as zf:
            root = ET.fromstring(zf.read("word/document.xml"))
        p_pr = root.findall(".//w:p", NS)[0].find("w:pPr", NS)
        self.assertEqual(p_pr.find("w:jc", NS).attrib[f"{{{W_NS}}}val"], "right")

    def test_hcss_id_selector_formats_single_run(self) -> None:
        input_docx = TMP / "input.docx"
        work = TMP / "work.hdocx"
        output_docx = TMP / "output.docx"
        _make_minimal_docx(input_docx)
        export_docx(input_docx, work)
        (work / "agent.edits.hcss").write_text(
            """
@hdocx-edit mode(all-runs);

#r-000001 {
  hdocx-italic: true;
}
""",
            encoding="utf-8",
            newline="\n",
        )

        plan = plan_hdocx(work)
        apply_report = apply_hdocx(work, output_docx)

        self.assertTrue(plan["ok"], plan)
        self.assertEqual(plan["patches"][0]["nodeId"], "r-000001")
        self.assertEqual(apply_report["patchSummary"]["byRiskClass"], {"fragment-preserving-eligible": 1})
        with zipfile.ZipFile(output_docx, "r") as zf:
            root = ET.fromstring(zf.read("word/document.xml"))
        r_pr = root.findall(".//w:r", NS)[0].find("w:rPr", NS)
        self.assertIsNotNone(r_pr.find("w:i", NS))

    def test_hcss_allow_empty_set_does_not_fail(self) -> None:
        input_docx = TMP / "input.docx"
        work = TMP / "work.hdocx"
        _make_minimal_docx(input_docx)
        export_docx(input_docx, work)
        (work / "agent.edits.hcss").write_text(
            """
@hdocx-set optional {
  select: [data-hdocx-style-id="Missing"];
  allow-empty: true;
}

@hdocx-edit mode(paragraph-formatting);

optional {
  hdocx-align: center;
}
""",
            encoding="utf-8",
            newline="\n",
        )

        plan = plan_hdocx(work)

        self.assertTrue(plan["ok"], plan)
        self.assertEqual(plan["patches"], [])

    def test_hcss_selector_functions_match_type_style_list_and_part(self) -> None:
        input_docx = TMP / "complex-academic.docx"
        work = TMP / "work.hdocx"
        _make_complex_academic_docx(input_docx)
        export_docx(input_docx, work)
        (work / "agent.edits.hcss").write_text(
            """
@hdocx-set body-style {
  select: style(BodyText);
}

@hdocx-set numbered {
  select: list(1, 0);
}

@hdocx-set header-part {
  select: part(/word/header1.xml, paragraph);
}

@hdocx-edit mode(paragraph-formatting);

body-style {
  hdocx-align: center;
}

numbered {
  hdocx-line-spacing: 1.5;
}

header-part {
  hdocx-first-line-indent: 1char;
}
""",
            encoding="utf-8",
            newline="\n",
        )

        plan = plan_hdocx(work)

        self.assertTrue(plan["ok"], plan)
        self.assertEqual(len(plan["patches"]), 3, plan)
        self.assertEqual(
            sorted((patch["nodeId"], patch["partPath"]) for patch in plan["patches"]),
            [("p-000001", "/word/document.xml"), ("p-000002", "/word/document.xml"), ("p-000010", "/word/header1.xml")],
        )

    def test_hcss_style_definition_patch(self) -> None:
        input_docx = TMP / "styled.docx"
        work = TMP / "work.hdocx"
        output_docx = TMP / "output.docx"
        _make_styled_docx(input_docx)
        export_docx(input_docx, work)
        manifest = read_json(work / "manifest.json")
        self.assertEqual(manifest["styles"]["Normal"]["name"], "Normal")
        self.assertEqual(manifest["styles"]["Normal"]["type"], "paragraph")
        (work / "agent.edits.hcss").write_text(
            """
@hdocx-set body {
  select: [data-hdocx-style-id="Normal"];
}

@hdocx-edit mode(style-definition);

body {
  hdocx-font-size: 14pt;
  hdocx-line-spacing: 1.5;
}
""",
            encoding="utf-8",
            newline="\n",
        )

        plan = plan_hdocx(work)
        apply_hdocx(work, output_docx)
        diff = diff_docx(input_docx, output_docx)

        self.assertTrue(plan["ok"], plan)
        self.assertEqual(plan["patches"][0]["operation"], "patch-style")
        self.assertEqual(diff["changed"], ["word/styles.xml"], diff)
        with zipfile.ZipFile(output_docx, "r") as zf:
            root = ET.fromstring(zf.read("word/styles.xml"))
        style = root.find("w:style", NS)
        self.assertEqual(style.attrib[f"{{{W_NS}}}styleId"], "Normal")
        self.assertEqual(style.find("w:rPr/w:sz", NS).attrib[f"{{{W_NS}}}val"], "28")
        self.assertEqual(style.find("w:pPr/w:spacing", NS).attrib[f"{{{W_NS}}}line"], "360")

    def test_hcss_paragraph_style_patch(self) -> None:
        input_docx = TMP / "styled.docx"
        work = TMP / "work.hdocx"
        output_docx = TMP / "output.docx"
        _make_styled_docx(input_docx)
        export_docx(input_docx, work)
        (work / "agent.edits.hcss").write_text(
            """
@hdocx-edit mode(paragraph-style);

#p-000001 {
  hdocx-style-id: Heading1;
}
""",
            encoding="utf-8",
            newline="\n",
        )

        plan = plan_hdocx(work)
        apply_hdocx(work, output_docx)
        diff = diff_docx(input_docx, output_docx)

        self.assertTrue(plan["ok"], plan)
        self.assertEqual(plan["patches"][0]["operation"], "patch-paragraph-style")
        self.assertEqual(diff["changed"], ["word/document.xml"], diff)
        with zipfile.ZipFile(output_docx, "r") as zf:
            root = ET.fromstring(zf.read("word/document.xml"))
        p_style = root.find(".//w:pPr/w:pStyle", NS)
        self.assertEqual(p_style.attrib[f"{{{W_NS}}}val"], "Heading1")

    def test_hcss_create_and_apply_paragraph_style(self) -> None:
        input_docx = TMP / "styled.docx"
        work = TMP / "work.hdocx"
        output_docx = TMP / "output.docx"
        _make_styled_docx(input_docx)
        export_docx(input_docx, work)
        (work / "agent.edits.hcss").write_text(
            """
@hdocx-style AgentBody {
  type: paragraph;
  name: "Agent Body";
  based-on: Normal;
  hdocx-font-size: 13pt;
  hdocx-line-spacing: 1.5;
}

@hdocx-edit mode(paragraph-style);

#p-000001 {
  hdocx-style-id: AgentBody;
}
""",
            encoding="utf-8",
            newline="\n",
        )

        plan = plan_hdocx(work)
        apply_hdocx(work, output_docx)
        diff = diff_docx(input_docx, output_docx)

        self.assertTrue(plan["ok"], plan)
        self.assertEqual([patch["operation"] for patch in plan["patches"]], ["create-style", "patch-paragraph-style"])
        self.assertEqual(diff["changed"], ["word/document.xml", "word/styles.xml"], diff)
        with zipfile.ZipFile(output_docx, "r") as zf:
            document_root = ET.fromstring(zf.read("word/document.xml"))
            styles_root = ET.fromstring(zf.read("word/styles.xml"))
        self.assertEqual(document_root.find(".//w:pPr/w:pStyle", NS).attrib[f"{{{W_NS}}}val"], "AgentBody")
        style = next(
            item for item in styles_root.findall("w:style", NS)
            if item.attrib.get(f"{{{W_NS}}}styleId") == "AgentBody"
        )
        self.assertEqual(style.find("w:name", NS).attrib[f"{{{W_NS}}}val"], "Agent Body")
        self.assertEqual(style.find("w:basedOn", NS).attrib[f"{{{W_NS}}}val"], "Normal")
        self.assertEqual(style.find("w:rPr/w:sz", NS).attrib[f"{{{W_NS}}}val"], "26")
        self.assertEqual(style.find("w:pPr/w:spacing", NS).attrib[f"{{{W_NS}}}line"], "360")

    def test_hcss_create_style_and_list_when_parts_are_missing(self) -> None:
        input_docx = TMP / "minimal.docx"
        work = TMP / "work.hdocx"
        output_docx = TMP / "output.docx"
        _make_minimal_docx(input_docx)
        export_docx(input_docx, work)
        (work / "agent.edits.hcss").write_text(
            """
@hdocx-style AgentBody {
  type: paragraph;
  name: "Agent Body";
  hdocx-font-size: 12pt;
}

@hdocx-list AgentList {
  hdocx-num-format: decimal;
  hdocx-level-text: "%1.";
}

@hdocx-edit mode(paragraph-style);

#p-000001 {
  hdocx-style-id: AgentBody;
}

@hdocx-edit mode(paragraph-numbering);

#p-000001 {
  hdocx-list-id: AgentList;
  hdocx-ilvl: 0;
}
""",
            encoding="utf-8",
            newline="\n",
        )

        plan = plan_hdocx(work)
        apply_hdocx(work, output_docx)
        diff = diff_docx(input_docx, output_docx)

        self.assertTrue(plan["ok"], plan)
        self.assertEqual(
            [patch["operation"] for patch in plan["patches"]],
            [
                "patch-content-types",
                "patch-relationships",
                "create-style",
                "patch-content-types",
                "patch-relationships",
                "create-numbering-list",
                "patch-paragraph-style",
                "patch-paragraph-numbering",
            ],
        )
        self.assertEqual(
            sorted(diff["changed"]),
            ["[Content_Types].xml", "word/document.xml"],
        )
        self.assertEqual(sorted(diff["rightOnly"]), ["word/_rels/document.xml.rels", "word/numbering.xml", "word/styles.xml"])
        with zipfile.ZipFile(output_docx, "r") as zf:
            content_types = ET.fromstring(zf.read("[Content_Types].xml"))
            rels = ET.fromstring(zf.read("word/_rels/document.xml.rels"))
            document_root = ET.fromstring(zf.read("word/document.xml"))
            styles_root = ET.fromstring(zf.read("word/styles.xml"))
            numbering_root = ET.fromstring(zf.read("word/numbering.xml"))
        overrides = {
            item.attrib["PartName"]: item.attrib["ContentType"]
            for item in content_types.findall("{http://schemas.openxmlformats.org/package/2006/content-types}Override")
        }
        self.assertIn("/word/styles.xml", overrides)
        self.assertIn("/word/numbering.xml", overrides)
        rel_ids = [rel.attrib["Id"] for rel in rels]
        rel_types = {rel.attrib["Type"] for rel in rels}
        self.assertEqual(len(rel_ids), len(set(rel_ids)))
        self.assertIn("http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles", rel_types)
        self.assertIn("http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering", rel_types)
        self.assertEqual(document_root.find(".//w:pPr/w:pStyle", NS).attrib[f"{{{W_NS}}}val"], "AgentBody")
        self.assertEqual(document_root.find(".//w:pPr/w:numPr/w:numId", NS).attrib[f"{{{W_NS}}}val"], "0")
        self.assertEqual(styles_root.find("w:style", NS).attrib[f"{{{W_NS}}}styleId"], "AgentBody")
        self.assertEqual(numbering_root.find("w:num", NS).attrib[f"{{{W_NS}}}numId"], "0")

    def test_hcss_delete_unused_style(self) -> None:
        input_docx = TMP / "styled.docx"
        work = TMP / "work.hdocx"
        output_docx = TMP / "output.docx"
        _make_styled_docx(input_docx)
        export_docx(input_docx, work)
        (work / "agent.edits.hcss").write_text(
            "@hdocx-delete-style(Heading1);\n",
            encoding="utf-8",
            newline="\n",
        )

        plan = plan_hdocx(work)
        apply_hdocx(work, output_docx)
        diff = diff_docx(input_docx, output_docx)

        self.assertTrue(plan["ok"], plan)
        self.assertEqual(plan["patches"][0]["operation"], "delete-style")
        self.assertEqual(diff["changed"], ["word/styles.xml"], diff)
        with zipfile.ZipFile(output_docx, "r") as zf:
            styles_root = ET.fromstring(zf.read("word/styles.xml"))
        self.assertIsNone(
            next((style for style in styles_root.findall("w:style", NS) if style.attrib.get(f"{{{W_NS}}}styleId") == "Heading1"), None)
        )

    def test_hcss_delete_used_style_is_rejected(self) -> None:
        input_docx = TMP / "styled.docx"
        work = TMP / "work.hdocx"
        _make_styled_docx(input_docx)
        export_docx(input_docx, work)
        (work / "agent.edits.hcss").write_text(
            "@hdocx-delete-style(Normal);\n",
            encoding="utf-8",
            newline="\n",
        )

        plan = plan_hdocx(work)

        self.assertFalse(plan["ok"], plan)
        self.assertEqual(plan["errors"][0]["code"], "HCSS_DELETE_STYLE_IN_USE")

    def test_hcss_zero_match_fails(self) -> None:
        input_docx = TMP / "input.docx"
        work = TMP / "work.hdocx"
        _make_minimal_docx(input_docx)
        export_docx(input_docx, work)
        (work / "agent.edits.hcss").write_text(
            """
@hdocx-edit mode(paragraph-formatting);

[data-hdocx-style-id="Missing"] {
  hdocx-align: center;
}
""",
            encoding="utf-8",
            newline="\n",
        )

        plan = plan_hdocx(work)

        self.assertFalse(plan["ok"], plan)
        self.assertEqual(plan["errors"][0]["code"], "HCSS_SELECTOR_ZERO_MATCH")

    def test_numbered_paragraph_projection_keeps_num_metadata(self) -> None:
        input_docx = TMP / "numbered.docx"
        work = TMP / "work.hdocx"
        output_docx = TMP / "output.docx"
        _make_numbered_docx(input_docx)

        export_docx(input_docx, work)
        html = (work / "document.html").read_text(encoding="utf-8")
        manifest = read_json(work / "manifest.json")
        report = roundtrip_docx(input_docx, TMP / "roundtrip.hdocx", output_docx, force=True)

        self.assertIn('class="hdocx-p hdocx-list"', html)
        self.assertIn('data-hdocx-num-id="1"', html)
        self.assertIn('data-hdocx-ilvl="0"', html)
        self.assertIn('data-hdocx-abstract-num-id="0"', html)
        self.assertIn('data-hdocx-num-format="decimal"', html)
        self.assertIn('data-hdocx-level-text="%1."', html)
        self.assertIn('data-hdocx-start="1"', html)
        self.assertIn('data-hdocx-number-suffix="tab"', html)
        self.assertIn('data-hdocx-num-indent-left="720"', html)
        self.assertEqual(manifest["numbering"]["nums"]["1"]["abstractNumId"], "0")
        self.assertEqual(manifest["numbering"]["abstractNums"]["0"]["levels"]["0"]["numFmt"], "decimal")
        self.assertTrue(report["byteIdentical"], report)

    def test_numbering_metadata_tamper_is_rejected_even_with_text_patch(self) -> None:
        input_docx = TMP / "numbered.docx"
        work = TMP / "work.hdocx"
        _make_numbered_docx(input_docx)
        export_docx(input_docx, work)
        html_path = work / "document.html"
        html_path.write_text(
            html_path.read_text(encoding="utf-8")
            .replace('data-hdocx-num-format="decimal"', 'data-hdocx-num-format="bullet"')
            .replace("Numbered item.", "Changed item."),
            encoding="utf-8",
            newline="\n",
        )

        plan = plan_hdocx(work)

        self.assertFalse(plan["ok"], plan)
        self.assertTrue(
            any(error["code"] == "HTML_READONLY_METADATA_MODIFIED" for error in plan["errors"]),
            plan,
        )

    def test_hcss_numbering_definition_patch(self) -> None:
        input_docx = TMP / "numbered.docx"
        work = TMP / "work.hdocx"
        output_docx = TMP / "output.docx"
        _make_numbered_docx(input_docx)
        export_docx(input_docx, work)
        (work / "agent.edits.hcss").write_text(
            """
@hdocx-set list-items {
  select: [data-hdocx-num-id="1"][data-hdocx-ilvl="0"];
}

@hdocx-edit mode(numbering-definition);

list-items {
  hdocx-num-format: upperLetter;
  hdocx-level-text: "Appendix %1)";
  hdocx-start: 3;
  hdocx-number-suffix: space;
  hdocx-num-indent-left: 960;
  hdocx-num-indent-hanging: 240;
}
""",
            encoding="utf-8",
            newline="\n",
        )

        plan = plan_hdocx(work)
        apply_hdocx(work, output_docx)
        diff = diff_docx(input_docx, output_docx)

        self.assertTrue(plan["ok"], plan)
        self.assertEqual(plan["patches"][0]["operation"], "patch-numbering-level")
        self.assertEqual(diff["changed"], ["word/numbering.xml"], diff)
        with zipfile.ZipFile(output_docx, "r") as zf:
            root = ET.fromstring(zf.read("word/numbering.xml"))
        level = root.find(".//w:lvl", NS)
        self.assertEqual(level.find("w:numFmt", NS).attrib[f"{{{W_NS}}}val"], "upperLetter")
        self.assertEqual(level.find("w:lvlText", NS).attrib[f"{{{W_NS}}}val"], "Appendix %1)")
        self.assertEqual(level.find("w:start", NS).attrib[f"{{{W_NS}}}val"], "3")
        self.assertEqual(level.find("w:suff", NS).attrib[f"{{{W_NS}}}val"], "space")
        ind = level.find("w:pPr/w:ind", NS)
        self.assertEqual(ind.attrib[f"{{{W_NS}}}left"], "960")
        self.assertEqual(ind.attrib[f"{{{W_NS}}}hanging"], "240")

    def test_hcss_create_list_and_assign_paragraph_numbering(self) -> None:
        input_docx = TMP / "numbered.docx"
        work = TMP / "work.hdocx"
        output_docx = TMP / "output.docx"
        _make_numbered_docx(input_docx)
        export_docx(input_docx, work)
        (work / "agent.edits.hcss").write_text(
            """
@hdocx-list AgentList {
  hdocx-num-format: lowerLetter;
  hdocx-level-text: "%1)";
  hdocx-start: 2;
  hdocx-number-suffix: space;
  hdocx-num-indent-left: 1080;
  hdocx-num-indent-hanging: 360;
}

@hdocx-edit mode(paragraph-numbering);

#p-000001 {
  hdocx-list-id: AgentList;
  hdocx-ilvl: 0;
}
""",
            encoding="utf-8",
            newline="\n",
        )

        plan = plan_hdocx(work)
        apply_hdocx(work, output_docx)
        diff = diff_docx(input_docx, output_docx)

        self.assertTrue(plan["ok"], plan)
        self.assertEqual([patch["operation"] for patch in plan["patches"]], ["create-numbering-list", "patch-paragraph-numbering"])
        self.assertEqual(diff["changed"], ["word/document.xml", "word/numbering.xml"], diff)
        with zipfile.ZipFile(output_docx, "r") as zf:
            document_root = ET.fromstring(zf.read("word/document.xml"))
            numbering_root = ET.fromstring(zf.read("word/numbering.xml"))
        self.assertEqual(document_root.find(".//w:numPr/w:numId", NS).attrib[f"{{{W_NS}}}val"], "2")
        new_level = next(
            level
            for abstract in numbering_root.findall("w:abstractNum", NS)
            if abstract.attrib.get(f"{{{W_NS}}}abstractNumId") == "1"
            for level in abstract.findall("w:lvl", NS)
        )
        self.assertEqual(new_level.find("w:numFmt", NS).attrib[f"{{{W_NS}}}val"], "lowerLetter")
        self.assertEqual(new_level.find("w:lvlText", NS).attrib[f"{{{W_NS}}}val"], "%1)")
        self.assertEqual(new_level.find("w:start", NS).attrib[f"{{{W_NS}}}val"], "2")
        self.assertEqual(numbering_root.find('w:num[@w:numId="2"]', NS).find("w:abstractNumId", NS).attrib[f"{{{W_NS}}}val"], "1")

    def test_hcss_create_multilevel_list_and_assign_second_level(self) -> None:
        input_docx = TMP / "numbered.docx"
        work = TMP / "work.hdocx"
        output_docx = TMP / "output.docx"
        _make_numbered_docx(input_docx)
        export_docx(input_docx, work)
        (work / "agent.edits.hcss").write_text(
            """
@hdocx-list AgentMulti {
  hdocx-num-format: decimal;
  hdocx-level-text: "%1.";
  hdocx-level-1-num-format: lowerLetter;
  hdocx-level-1-level-text: "%2)";
  hdocx-level-1-start: 4;
  hdocx-level-1-number-suffix: space;
  hdocx-level-1-num-indent-left: 1440;
  hdocx-level-1-num-indent-hanging: 360;
}

@hdocx-edit mode(paragraph-numbering);

#p-000001 {
  hdocx-list-id: AgentMulti;
  hdocx-ilvl: 1;
}
""",
            encoding="utf-8",
            newline="\n",
        )

        plan = plan_hdocx(work)
        apply_hdocx(work, output_docx)

        self.assertTrue(plan["ok"], plan)
        self.assertEqual(plan["patches"][0]["levels"]["1"]["num-format"], "lowerLetter")
        with zipfile.ZipFile(output_docx, "r") as zf:
            document_root = ET.fromstring(zf.read("word/document.xml"))
            numbering_root = ET.fromstring(zf.read("word/numbering.xml"))
        self.assertEqual(document_root.find(".//w:numPr/w:ilvl", NS).attrib[f"{{{W_NS}}}val"], "1")
        abstract = next(
            item
            for item in numbering_root.findall("w:abstractNum", NS)
            if item.attrib.get(f"{{{W_NS}}}abstractNumId") == "1"
        )
        levels = {level.attrib[f"{{{W_NS}}}ilvl"]: level for level in abstract.findall("w:lvl", NS)}
        self.assertEqual(set(levels), {"0", "1"})
        self.assertEqual(levels["1"].find("w:numFmt", NS).attrib[f"{{{W_NS}}}val"], "lowerLetter")
        self.assertEqual(levels["1"].find("w:lvlText", NS).attrib[f"{{{W_NS}}}val"], "%2)")
        self.assertEqual(levels["1"].find("w:start", NS).attrib[f"{{{W_NS}}}val"], "4")
        self.assertEqual(levels["1"].find("w:suff", NS).attrib[f"{{{W_NS}}}val"], "space")
        ind = levels["1"].find("w:pPr/w:ind", NS)
        self.assertEqual(ind.attrib[f"{{{W_NS}}}left"], "1440")
        self.assertEqual(ind.attrib[f"{{{W_NS}}}hanging"], "360")

    def test_image_alt_text_patch_changes_only_document_xml(self) -> None:
        input_docx = TMP / "image.docx"
        work = TMP / "work.hdocx"
        output_docx = TMP / "output.docx"
        _make_image_docx(input_docx)
        with zipfile.ZipFile(input_docx, "r") as zf:
            original_document_xml = zf.read("word/document.xml")
        export_docx(input_docx, work)
        html_path = work / "document.html"
        html = html_path.read_text(encoding="utf-8")
        self.assertIn("[drawing]", html)
        self.assertIn('<span class="hdocx-drawing-text" hidden>[drawing]</span>', html)
        self.assertIn('<img class="hdocx-image-preview"', html)
        self.assertIn('src="parts/word/media/image1.png"', html)
        self.assertIn('style="width: 96px; height: 48px; object-fit: contain"', html)
        self.assertIn('data-hdocx-alt="Old alt"', html)
        html_path.write_text(
            html.replace('data-hdocx-alt="Old alt"', 'data-hdocx-alt="New alt"'),
            encoding="utf-8",
            newline="\n",
        )

        plan = plan_hdocx(work)
        apply_hdocx(work, output_docx)
        diff = diff_docx(input_docx, output_docx)

        self.assertTrue(plan["ok"], plan)
        self.assertEqual(plan["patches"][0]["operation"], "patch-drawing-alt")
        self.assertEqual(diff["changed"], ["word/document.xml"], diff)
        with zipfile.ZipFile(output_docx, "r") as zf:
            document_xml_bytes = zf.read("word/document.xml")
            image_bytes = zf.read("word/media/image1.png")
        self.assertEqual(
            document_xml_bytes,
            original_document_xml.replace(b'descr="Old alt"', b'descr="New alt"'),
        )
        document_xml = document_xml_bytes.decode("utf-8")
        self.assertIn('descr="New alt"', document_xml)
        self.assertEqual(image_bytes, b"fakepng")

    def test_image_size_patch_changes_extent_only(self) -> None:
        input_docx = TMP / "image.docx"
        work = TMP / "work.hdocx"
        output_docx = TMP / "output.docx"
        _make_image_docx(input_docx)
        with zipfile.ZipFile(input_docx, "r") as zf:
            original_document_xml = zf.read("word/document.xml")
        export_docx(input_docx, work)
        html_path = work / "document.html"
        html = html_path.read_text(encoding="utf-8")
        self.assertIn('data-hdocx-width-emu="914400"', html)
        self.assertIn('data-hdocx-height-emu="457200"', html)
        html_path.write_text(
            html.replace('data-hdocx-width-emu="914400"', 'data-hdocx-width-emu="1828800"')
            .replace('data-hdocx-height-emu="457200"', 'data-hdocx-height-emu="914400"'),
            encoding="utf-8",
            newline="\n",
        )

        plan = plan_hdocx(work)
        apply_hdocx(work, output_docx)
        diff = diff_docx(input_docx, output_docx)

        self.assertTrue(plan["ok"], plan)
        self.assertEqual(plan["patches"][0]["operation"], "patch-drawing-properties")
        self.assertEqual(plan["patches"][0]["newProperties"]["width-emu"], "1828800")
        self.assertEqual(plan["patches"][0]["newProperties"]["height-emu"], "914400")
        self.assertEqual(diff["changed"], ["word/document.xml"], diff)
        with zipfile.ZipFile(output_docx, "r") as zf:
            document_xml_bytes = zf.read("word/document.xml")
            root = ET.fromstring(document_xml_bytes)
            image_bytes = zf.read("word/media/image1.png")
        self.assertEqual(
            document_xml_bytes,
            original_document_xml.replace(b'cx="914400"', b'cx="1828800"').replace(
                b'cy="457200"', b'cy="914400"'
            ),
        )
        extent = root.find(".//wp:extent", NS)
        self.assertEqual(extent.attrib["cx"], "1828800")
        self.assertEqual(extent.attrib["cy"], "914400")
        self.assertEqual(image_bytes, b"fakepng")

    def test_hcss_image_formatting_sets_size_and_host_paragraph_spacing(self) -> None:
        input_docx = TMP / "image.docx"
        work = TMP / "work.hdocx"
        output_docx = TMP / "output.docx"
        _make_image_docx(input_docx)
        export_docx(input_docx, work)
        (work / "agent.edits.hcss").write_text(
            """
@hdocx-edit mode(image-formatting);

#r-000001 {
  hdocx-width-emu: 1828800;
  hdocx-height-emu: 914400;
  hdocx-alt: "Scaled image";
  hdocx-paragraph-line-spacing: 1;
  hdocx-paragraph-space-before: 0;
  hdocx-paragraph-space-after: 0;
}
""",
            encoding="utf-8",
            newline="\n",
        )

        plan = plan_hdocx(work)
        apply_hdocx(work, output_docx)

        self.assertTrue(plan["ok"], plan)
        self.assertEqual(
            [patch["operation"] for patch in plan["patches"]],
            ["patch-drawing-properties", "patch-paragraph"],
        )
        with zipfile.ZipFile(output_docx, "r") as zf:
            root = ET.fromstring(zf.read("word/document.xml"))
        extent = root.find(".//wp:extent", NS)
        self.assertEqual(extent.attrib["cx"], "1828800")
        self.assertEqual(extent.attrib["cy"], "914400")
        doc_pr = root.find(".//wp:docPr", NS)
        self.assertEqual(doc_pr.attrib["descr"], "Scaled image")
        spacing = root.find(".//w:p/w:pPr/w:spacing", NS)
        self.assertEqual(spacing.attrib[f"{{{W_NS}}}line"], "240")
        self.assertEqual(spacing.attrib[f"{{{W_NS}}}lineRule"], "auto")
        self.assertEqual(spacing.attrib[f"{{{W_NS}}}before"], "0")
        self.assertEqual(spacing.attrib[f"{{{W_NS}}}after"], "0")

    def test_image_size_removal_is_rejected(self) -> None:
        input_docx = TMP / "image.docx"
        work = TMP / "work.hdocx"
        _make_image_docx(input_docx)
        export_docx(input_docx, work)
        html_path = work / "document.html"
        html_path.write_text(
            html_path.read_text(encoding="utf-8").replace(' data-hdocx-width-emu="914400"', ""),
            encoding="utf-8",
            newline="\n",
        )

        plan = plan_hdocx(work)

        self.assertFalse(plan["ok"], plan)
        self.assertEqual(plan["errors"][0]["code"], "HTML_INVALID_DRAWING_PROPERTY")

    def test_media_part_replacement_changes_only_media_entry(self) -> None:
        input_docx = TMP / "image.docx"
        work = TMP / "work.hdocx"
        output_docx = TMP / "output.docx"
        _make_image_docx(input_docx)
        export_docx(input_docx, work)
        (work / "parts" / "word" / "media" / "image1.png").write_bytes(b"newpng")

        plan = plan_hdocx(work)
        apply_report = apply_hdocx(work, output_docx)
        diff = diff_docx(input_docx, output_docx)

        self.assertTrue(plan["ok"], plan)
        self.assertEqual(plan["patches"][0]["operation"], "replace-media")
        self.assertEqual(apply_report["mode"], "media-patch")
        self.assertEqual(diff["changed"], ["word/media/image1.png"], diff)
        self.assertTrue(diff["semanticDiff"]["identical"], diff)
        with zipfile.ZipFile(output_docx, "r") as zf:
            self.assertEqual(zf.read("word/media/image1.png"), b"newpng")
            document_xml = zf.read("word/document.xml").decode("utf-8")
        self.assertIn('descr="Old alt"', document_xml)

    def test_non_media_part_replacement_is_rejected(self) -> None:
        input_docx = TMP / "input.docx"
        work = TMP / "work.hdocx"
        output_docx = TMP / "output.docx"
        _make_minimal_docx(input_docx)
        export_docx(input_docx, work)
        document_part = work / "parts" / "word" / "document.xml"
        document_part.write_text(
            document_part.read_text(encoding="utf-8").replace("Hello", "Tampered"),
            encoding="utf-8",
            newline="\n",
        )

        plan = plan_hdocx(work)

        self.assertFalse(plan["ok"], plan)
        self.assertEqual(plan["errors"][0]["code"], "UNSUPPORTED_PART_REPLACEMENT")
        with self.assertRaises(Exception):
            apply_hdocx(work, output_docx)

    def test_hcss_insert_image_adds_media_relationship_and_content_type(self) -> None:
        input_docx = TMP / "input.docx"
        work = TMP / "work.hdocx"
        output_docx = TMP / "output.docx"
        _make_minimal_docx(input_docx)
        export_docx(input_docx, work)
        assets_dir = work / "assets"
        assets_dir.mkdir()
        (assets_dir / "new.png").write_bytes(b"insertedpng")
        (work / "agent.edits.hcss").write_text(
            """
@hdocx-insert-image after(#p-000001) {
  source: assets/new.png;
  alt: "Inserted figure";
  width-emu: 914400;
  height-emu: 457200;
}
""",
            encoding="utf-8",
            newline="\n",
        )

        plan = plan_hdocx(work)
        apply_report = apply_hdocx(work, output_docx)
        diff = diff_docx(input_docx, output_docx)

        self.assertTrue(plan["ok"], plan)
        self.assertEqual(
            [patch["operation"] for patch in plan["patches"]],
            ["insert-image-after-paragraph", "patch-relationships", "patch-content-types", "add-media"],
        )
        self.assertEqual(apply_report["mode"], "document-and-media-patch")
        self.assertIn("word/media/hdocx-image-000001.png", diff["rightOnly"], diff)
        self.assertIn("[Content_Types].xml", diff["changed"], diff)
        self.assertIn("word/document.xml", diff["changed"], diff)
        self.assertIn("word/_rels/document.xml.rels", diff["rightOnly"], diff)
        with zipfile.ZipFile(output_docx, "r") as zf:
            self.assertEqual(zf.read("word/media/hdocx-image-000001.png"), b"insertedpng")
            document_xml = zf.read("word/document.xml").decode("utf-8")
            rels_xml = zf.read("word/_rels/document.xml.rels").decode("utf-8")
            content_types = zf.read("[Content_Types].xml").decode("utf-8")
        self.assertIn('descr="Inserted figure"', document_xml)
        self.assertIn('cx="914400"', document_xml)
        self.assertIn('cy="457200"', document_xml)
        self.assertIn('embed="rId1"', document_xml)
        self.assertIn('Target="media/hdocx-image-000001.png"', rels_xml)
        self.assertIn('Extension="png"', content_types)
        self.assertIn('ContentType="image/png"', content_types)

    def test_hcss_insert_image_before_paragraph(self) -> None:
        input_docx = TMP / "input.docx"
        work = TMP / "work.hdocx"
        output_docx = TMP / "output.docx"
        _make_minimal_docx(input_docx)
        export_docx(input_docx, work)
        assets_dir = work / "assets"
        assets_dir.mkdir()
        (assets_dir / "before.png").write_bytes(b"beforepng")
        (work / "agent.edits.hcss").write_text(
            """
@hdocx-insert-image before(#p-000001) {
  source: assets/before.png;
  width-emu: 914400;
  height-emu: 457200;
}
""",
            encoding="utf-8",
            newline="\n",
        )

        plan = plan_hdocx(work)
        apply_hdocx(work, output_docx)

        self.assertTrue(plan["ok"], plan)
        self.assertEqual(plan["patches"][0]["operation"], "insert-image-before-paragraph")
        with zipfile.ZipFile(output_docx, "r") as zf:
            root = ET.fromstring(zf.read("word/document.xml"))
        body_children = list(root.find("w:body", NS))
        self.assertEqual(body_children[0].find(".//wp:inline", NS).tag, f"{{{WP_NS}}}inline")
        self.assertIn("Hello strict roundtrip.", "".join(t.text or "" for t in body_children[1].findall(".//w:t", NS)))

    def test_hcss_insert_image_into_header_part(self) -> None:
        input_docx = TMP / "header.docx"
        work = TMP / "work.hdocx"
        output_docx = TMP / "output.docx"
        _make_header_docx(input_docx)
        export_docx(input_docx, work)
        assets_dir = work / "assets"
        assets_dir.mkdir()
        (assets_dir / "header.png").write_bytes(b"headerpng")
        (work / "agent.edits.hcss").write_text(
            """
@hdocx-insert-image after(#p-000002) {
  source: assets/header.png;
  alt: "Header image";
  width-emu: 914400;
  height-emu: 457200;
}
""",
            encoding="utf-8",
            newline="\n",
        )

        plan = plan_hdocx(work)
        apply_hdocx(work, output_docx)
        diff = diff_docx(input_docx, output_docx)

        self.assertTrue(plan["ok"], plan)
        self.assertEqual(plan["patches"][0]["entryName"], "word/header1.xml")
        self.assertEqual(plan["patches"][1]["entryName"], "word/_rels/header1.xml.rels")
        self.assertIn("word/header1.xml", diff["changed"], diff)
        self.assertIn("word/_rels/header1.xml.rels", diff["rightOnly"], diff)
        with zipfile.ZipFile(output_docx, "r") as zf:
            header_xml = zf.read("word/header1.xml").decode("utf-8")
            header_rels = zf.read("word/_rels/header1.xml.rels").decode("utf-8")
            media_bytes = zf.read("word/media/hdocx-image-000001.png")
        self.assertIn("Header image", header_xml)
        self.assertIn('Target="media/hdocx-image-000001.png"', header_rels)
        self.assertEqual(media_bytes, b"headerpng")

    def test_hcss_insert_image_rejects_outside_source(self) -> None:
        input_docx = TMP / "input.docx"
        work = TMP / "work.hdocx"
        _make_minimal_docx(input_docx)
        export_docx(input_docx, work)
        (work / "agent.edits.hcss").write_text(
            """
@hdocx-insert-image after(#p-000001) {
  source: ../outside.png;
  width-emu: 914400;
  height-emu: 457200;
}
""",
            encoding="utf-8",
            newline="\n",
        )

        plan = plan_hdocx(work)

        self.assertFalse(plan["ok"], plan)
        self.assertEqual(plan["errors"][0]["code"], "HCSS_INSERT_IMAGE_UNSAFE_SOURCE")

    def test_image_drawing_placeholder_text_is_protected(self) -> None:
        input_docx = TMP / "image.docx"
        work = TMP / "work.hdocx"
        output_docx = TMP / "output.docx"
        _make_image_docx(input_docx)
        export_docx(input_docx, work)
        html_path = work / "document.html"
        html = html_path.read_text(encoding="utf-8")
        html_path.write_text(
            html.replace("[drawing]", "tampered drawing"),
            encoding="utf-8",
            newline="\n",
        )

        plan = plan_hdocx(work)

        self.assertFalse(plan["ok"], plan)
        self.assertEqual(plan["errors"][0]["code"], "PROTECTED_RUN_TEXT_MODIFIED")
        with self.assertRaises(Exception):
            apply_hdocx(work, output_docx)

    def test_table_cell_text_patch(self) -> None:
        input_docx = TMP / "table.docx"
        work = TMP / "work.hdocx"
        output_docx = TMP / "output.docx"
        _make_table_docx(input_docx)
        export_docx(input_docx, work)
        html_path = work / "document.html"
        html = html_path.read_text(encoding="utf-8")
        self.assertIn("<table", html)
        self.assertIn("<td", html)
        html_path.write_text(
            html.replace("Cell text.", "Updated cell."),
            encoding="utf-8",
            newline="\n",
        )

        plan = plan_hdocx(work)
        apply_hdocx(work, output_docx)

        self.assertTrue(plan["ok"], plan)
        with zipfile.ZipFile(output_docx, "r") as zf:
            document_xml = zf.read("word/document.xml").decode("utf-8")
        self.assertIn("Updated cell.", document_xml)
        self.assertNotIn("Cell text.", document_xml)

    def test_hcss_insert_table_row_after_simple_row(self) -> None:
        input_docx = TMP / "table-two-cell.docx"
        work = TMP / "work.hdocx"
        output_docx = TMP / "output.docx"
        _make_two_cell_table_docx(input_docx)
        export_docx(input_docx, work)
        (work / "agent.edits.hcss").write_text(
            """
@hdocx-insert-table-row after(#tr-000001) {
  cells: "Inserted A|Inserted B";
}
""",
            encoding="utf-8",
            newline="\n",
        )

        plan = plan_hdocx(work)
        apply_hdocx(work, output_docx)
        diff = diff_docx(input_docx, output_docx)

        self.assertTrue(plan["ok"], plan)
        self.assertEqual(plan["patches"][0]["operation"], "insert-table-row-after")
        self.assertEqual(diff["changed"], ["word/document.xml"], diff)
        with zipfile.ZipFile(output_docx, "r") as zf:
            root = ET.fromstring(zf.read("word/document.xml"))
        rows = root.findall(".//w:tr", NS)
        self.assertEqual(len(rows), 2)
        inserted_texts = ["".join(t.text or "" for t in cell.findall(".//w:t", NS)) for cell in rows[1].findall("w:tc", NS)]
        self.assertEqual(inserted_texts, ["Inserted A", "Inserted B"])

    def test_hcss_insert_table_row_rejects_merged_table(self) -> None:
        input_docx = TMP / "merged-table.docx"
        work = TMP / "work.hdocx"
        _make_merged_table_docx(input_docx)
        export_docx(input_docx, work)
        html = (work / "document.html").read_text(encoding="utf-8")
        self.assertIn('data-hdocx-complex-merge="true"', html)
        (work / "agent.edits.hcss").write_text(
            """
@hdocx-insert-table-row after(#tr-000001) {
  cells: "Unsafe";
}
""",
            encoding="utf-8",
            newline="\n",
        )

        plan = plan_hdocx(work)

        self.assertFalse(plan["ok"], plan)
        self.assertEqual(plan["errors"][0]["code"], "HCSS_INSERT_TABLE_ROW_COMPLEX_TABLE")

    def test_hcss_delete_table_row(self) -> None:
        input_docx = TMP / "two-row-table.docx"
        work = TMP / "work.hdocx"
        output_docx = TMP / "output.docx"
        _make_two_row_table_docx(input_docx)
        export_docx(input_docx, work)
        (work / "agent.edits.hcss").write_text(
            """
@hdocx-delete-table-row(#tr-000001);
""",
            encoding="utf-8",
            newline="\n",
        )

        plan = plan_hdocx(work)
        apply_hdocx(work, output_docx)
        diff = diff_docx(input_docx, output_docx)

        self.assertTrue(plan["ok"], plan)
        self.assertEqual(plan["patches"][0]["operation"], "delete-table-row")
        self.assertEqual(diff["changed"], ["word/document.xml"], diff)
        with zipfile.ZipFile(output_docx, "r") as zf:
            root = ET.fromstring(zf.read("word/document.xml"))
        rows = root.findall(".//w:tr", NS)
        self.assertEqual(len(rows), 1)
        remaining_texts = ["".join(t.text or "" for t in cell.findall(".//w:t", NS)) for cell in rows[0].findall("w:tc", NS)]
        self.assertEqual(remaining_texts, ["A2", "B2"])

    def test_hcss_delete_last_table_row_is_rejected(self) -> None:
        input_docx = TMP / "table.docx"
        work = TMP / "work.hdocx"
        _make_table_docx(input_docx)
        export_docx(input_docx, work)
        (work / "agent.edits.hcss").write_text(
            """
@hdocx-delete-table-row(#tr-000001);
""",
            encoding="utf-8",
            newline="\n",
        )

        plan = plan_hdocx(work)

        self.assertFalse(plan["ok"], plan)
        self.assertEqual(plan["errors"][0]["code"], "HCSS_DELETE_TABLE_ROW_LAST_ROW")

    def test_hcss_insert_table_column_after_simple_cell(self) -> None:
        input_docx = TMP / "two-row-table.docx"
        work = TMP / "work.hdocx"
        output_docx = TMP / "output.docx"
        _make_two_row_table_docx(input_docx)
        export_docx(input_docx, work)
        (work / "agent.edits.hcss").write_text(
            """
@hdocx-insert-table-column after(#tc-000001) {
  cells: "Inserted 1|Inserted 2";
}
""",
            encoding="utf-8",
            newline="\n",
        )

        plan = plan_hdocx(work)
        apply_hdocx(work, output_docx)

        self.assertTrue(plan["ok"], plan)
        self.assertEqual(plan["patches"][0]["operation"], "insert-table-column-after")
        with zipfile.ZipFile(output_docx, "r") as zf:
            root = ET.fromstring(zf.read("word/document.xml"))
        rows = root.findall(".//w:tr", NS)
        row_texts = [
            ["".join(t.text or "" for t in cell.findall(".//w:t", NS)) for cell in row.findall("w:tc", NS)]
            for row in rows
        ]
        self.assertEqual(row_texts, [["A1", "Inserted 1", "B1"], ["A2", "Inserted 2", "B2"]])

    def test_hcss_delete_table_column(self) -> None:
        input_docx = TMP / "two-row-table.docx"
        work = TMP / "work.hdocx"
        output_docx = TMP / "output.docx"
        _make_two_row_table_docx(input_docx)
        export_docx(input_docx, work)
        (work / "agent.edits.hcss").write_text(
            """
@hdocx-delete-table-column(#tc-000002);
""",
            encoding="utf-8",
            newline="\n",
        )

        plan = plan_hdocx(work)
        apply_hdocx(work, output_docx)

        self.assertTrue(plan["ok"], plan)
        self.assertEqual(plan["patches"][0]["operation"], "delete-table-column")
        with zipfile.ZipFile(output_docx, "r") as zf:
            root = ET.fromstring(zf.read("word/document.xml"))
        rows = root.findall(".//w:tr", NS)
        row_texts = [
            ["".join(t.text or "" for t in cell.findall(".//w:t", NS)) for cell in row.findall("w:tc", NS)]
            for row in rows
        ]
        self.assertEqual(row_texts, [["A1"], ["A2"]])

    def test_header_part_text_patch_changes_only_header_entry(self) -> None:
        input_docx = TMP / "header.docx"
        work = TMP / "work.hdocx"
        output_docx = TMP / "output.docx"
        _make_header_docx(input_docx)
        export_docx(input_docx, work)
        html_path = work / "document.html"
        html = html_path.read_text(encoding="utf-8")
        self.assertIn('data-hdocx-part="/word/header1.xml"', html)
        html_path.write_text(
            html.replace("Header text.", "Updated header."),
            encoding="utf-8",
            newline="\n",
        )

        plan = plan_hdocx(work)
        apply_hdocx(work, output_docx)
        diff = diff_docx(input_docx, output_docx)

        self.assertTrue(plan["ok"], plan)
        self.assertEqual(plan["patches"][0]["entryName"], "word/header1.xml")
        self.assertEqual(diff["changed"], ["word/header1.xml"], diff)
        with zipfile.ZipFile(output_docx, "r") as zf:
            header_xml = zf.read("word/header1.xml").decode("utf-8")
        self.assertIn("Updated header.", header_xml)
        self.assertNotIn("Header text.", header_xml)

    def test_footnote_part_text_patch_changes_only_footnotes_entry(self) -> None:
        input_docx = TMP / "footnote.docx"
        work = TMP / "work.hdocx"
        output_docx = TMP / "output.docx"
        _make_footnote_docx(input_docx)
        export_docx(input_docx, work)
        html_path = work / "document.html"
        html = html_path.read_text(encoding="utf-8")
        self.assertIn('data-hdocx-part="/word/footnotes.xml"', html)
        html_path.write_text(
            html.replace("Footnote text.", "Updated footnote."),
            encoding="utf-8",
            newline="\n",
        )

        plan = plan_hdocx(work)
        apply_hdocx(work, output_docx)
        diff = diff_docx(input_docx, output_docx)

        self.assertTrue(plan["ok"], plan)
        self.assertEqual(plan["patches"][0]["entryName"], "word/footnotes.xml")
        self.assertEqual(diff["changed"], ["word/footnotes.xml"], diff)
        with zipfile.ZipFile(output_docx, "r") as zf:
            footnotes_xml = zf.read("word/footnotes.xml").decode("utf-8")
        self.assertIn("Updated footnote.", footnotes_xml)
        self.assertNotIn("Footnote text.", footnotes_xml)

    def test_footnote_reference_run_is_protected(self) -> None:
        input_docx = TMP / "footnote.docx"
        work = TMP / "work.hdocx"
        output_docx = TMP / "output.docx"
        _make_footnote_docx(input_docx)
        export_docx(input_docx, work)
        html_path = work / "document.html"
        html = html_path.read_text(encoding="utf-8")
        self.assertIn("[footnote-ref]", html)
        html_path.write_text(
            html.replace("[footnote-ref]", "tampered-ref"),
            encoding="utf-8",
            newline="\n",
        )

        plan = plan_hdocx(work)

        self.assertFalse(plan["ok"], plan)
        self.assertEqual(plan["errors"][0]["code"], "PROTECTED_RUN_TEXT_MODIFIED")
        with self.assertRaises(Exception):
            apply_hdocx(work, output_docx)

    def test_comment_ranges_and_comment_body_are_protected(self) -> None:
        input_docx = TMP / "comment.docx"
        work = TMP / "work.hdocx"
        output_docx = TMP / "output.docx"
        _make_comment_docx(input_docx)
        export_docx(input_docx, work)
        html_path = work / "document.html"
        html = html_path.read_text(encoding="utf-8")
        self.assertIn("[comment-range-start:0]", html)
        self.assertIn("[comment-range-end:0]", html)
        self.assertIn("[comment-ref]", html)
        self.assertIn("[comment] Reviewer note.", html)
        html_path.write_text(
            html.replace("[comment] Reviewer note.", "[comment] Changed note."),
            encoding="utf-8",
            newline="\n",
        )

        plan = plan_hdocx(work)

        self.assertFalse(plan["ok"], plan)
        self.assertEqual(plan["errors"][0]["code"], "PROTECTED_NODE_TEXT_MODIFIED")
        with self.assertRaises(Exception):
            apply_hdocx(work, output_docx)

    def test_hcss_comment_text_patch_changes_only_comments_part(self) -> None:
        input_docx = TMP / "comment.docx"
        work = TMP / "work.hdocx"
        output_docx = TMP / "output.docx"
        _make_comment_docx(input_docx)
        export_docx(input_docx, work)
        (work / "agent.edits.hcss").write_text(
            """
@hdocx-edit mode(comment-text);

[data-hdocx-protected-kind="comment"] {
  hdocx-text: "Changed note.";
}
""",
            encoding="utf-8",
            newline="\n",
        )

        plan = plan_hdocx(work)
        apply_hdocx(work, output_docx)
        diff = diff_docx(input_docx, output_docx)

        self.assertTrue(plan["ok"], plan)
        self.assertEqual(plan["patches"][0]["operation"], "patch-comment-text")
        self.assertEqual(diff["changed"], ["word/comments.xml"], diff)
        with zipfile.ZipFile(output_docx, "r") as zf:
            document_xml = zf.read("word/document.xml").decode("utf-8")
            comments_xml = zf.read("word/comments.xml").decode("utf-8")
        self.assertIn("Commented text.", document_xml)
        self.assertIn("Changed note.", comments_xml)
        self.assertNotIn("Reviewer note.", comments_xml)

    def test_revision_wrapper_is_protected(self) -> None:
        input_docx = TMP / "revision.docx"
        work = TMP / "work.hdocx"
        output_docx = TMP / "output.docx"
        _make_revision_docx(input_docx)
        export_docx(input_docx, work)
        html_path = work / "document.html"
        html = html_path.read_text(encoding="utf-8")
        self.assertIn("[revision-insert:7]", html)
        html_path.write_text(
            html.replace("[revision-insert:7]", "tampered revision"),
            encoding="utf-8",
            newline="\n",
        )

        plan = plan_hdocx(work)

        self.assertFalse(plan["ok"], plan)
        self.assertEqual(plan["errors"][0]["code"], "PROTECTED_NODE_TEXT_MODIFIED")
        with self.assertRaises(Exception):
            apply_hdocx(work, output_docx)

    def test_hcss_revision_insert_accept_unwraps_revision(self) -> None:
        input_docx = TMP / "revision.docx"
        work = TMP / "work.hdocx"
        output_docx = TMP / "output.docx"
        _make_revision_docx(input_docx)
        export_docx(input_docx, work)
        (work / "agent.edits.hcss").write_text(
            """
@hdocx-edit mode(revision-action);

[data-hdocx-protected-kind="revision-insert"] {
  hdocx-action: accept;
}
""",
            encoding="utf-8",
            newline="\n",
        )

        plan = plan_hdocx(work)
        apply_hdocx(work, output_docx)
        diff = diff_docx(input_docx, output_docx)

        self.assertTrue(plan["ok"], plan)
        self.assertEqual(plan["patches"][0]["operation"], "patch-revision-action")
        self.assertEqual(diff["changed"], ["word/document.xml"], diff)
        with zipfile.ZipFile(output_docx, "r") as zf:
            document_xml = zf.read("word/document.xml").decode("utf-8")
        self.assertIn("Inserted text.", document_xml)
        self.assertNotIn("<w:ins", document_xml)

    def test_hcss_equation_omml_replaces_protected_equation(self) -> None:
        input_docx = TMP / "equation.docx"
        work = TMP / "work.hdocx"
        output_docx = TMP / "output.docx"
        _make_equation_docx(input_docx)
        export_docx(input_docx, work)
        equation_dir = work / "equations"
        equation_dir.mkdir()
        (equation_dir / "replacement.omml").write_text(
            '<m:oMath xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math"><m:r><m:t>y=2</m:t></m:r></m:oMath>',
            encoding="utf-8",
            newline="\n",
        )
        (work / "agent.edits.hcss").write_text(
            """
@hdocx-edit mode(equation-omml);

[data-hdocx-protected-kind="equation"] {
  hdocx-omml-source: equations/replacement.omml;
}
""",
            encoding="utf-8",
            newline="\n",
        )

        plan = plan_hdocx(work)
        apply_hdocx(work, output_docx)
        diff = diff_docx(input_docx, output_docx)

        self.assertTrue(plan["ok"], plan)
        self.assertEqual(plan["patches"][0]["operation"], "patch-equation-omml")
        self.assertEqual(diff["changed"], ["word/document.xml"], diff)
        with zipfile.ZipFile(output_docx, "r") as zf:
            root = ET.fromstring(zf.read("word/document.xml"))
        math_text = "".join(text.text or "" for text in root.findall(".//m:t", NS))
        self.assertIn("y=2", math_text)
        self.assertNotIn("x=1", math_text)

    def test_hcss_equation_omml_rejects_non_omml_source(self) -> None:
        input_docx = TMP / "equation.docx"
        work = TMP / "work.hdocx"
        _make_equation_docx(input_docx)
        export_docx(input_docx, work)
        equation_dir = work / "equations"
        equation_dir.mkdir()
        (equation_dir / "bad.xml").write_text(
            '<w:r xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:t>not math</w:t></w:r>',
            encoding="utf-8",
            newline="\n",
        )
        (work / "agent.edits.hcss").write_text(
            """
@hdocx-edit mode(equation-omml);

[data-hdocx-protected-kind="equation"] {
  hdocx-omml-source: equations/bad.xml;
}
""",
            encoding="utf-8",
            newline="\n",
        )

        plan = plan_hdocx(work)

        self.assertFalse(plan["ok"], plan)
        self.assertEqual(plan["errors"][0]["code"], "HCSS_EQUATION_OMML_INVALID_ROOT")

    def test_apply_rejects_protected_run_text_patch(self) -> None:
        input_docx = TMP / "protected.docx"
        work = TMP / "work.hdocx"
        output_docx = TMP / "output.docx"
        _make_protected_run_docx(input_docx)
        export_docx(input_docx, work)
        html_path = work / "document.html"
        html_path.write_text(
            html_path.read_text(encoding="utf-8").replace("[field]", "tampered"),
            encoding="utf-8",
            newline="\n",
        )

        plan = plan_hdocx(work)

        self.assertFalse(plan["ok"], plan)
        self.assertEqual(plan["errors"][0]["code"], "PROTECTED_RUN_TEXT_MODIFIED")
        with self.assertRaises(Exception):
            apply_hdocx(work, output_docx)

    def test_apply_rejects_protected_run_format_patch(self) -> None:
        input_docx = TMP / "protected.docx"
        work = TMP / "work.hdocx"
        output_docx = TMP / "output.docx"
        _make_protected_run_docx(input_docx)
        export_docx(input_docx, work)
        html_path = work / "document.html"
        html_path.write_text(
            html_path.read_text(encoding="utf-8").replace(
                'data-hdocx-id="r-000001" data-hdocx-lock="protected" data-hdocx-part="/word/document.xml">',
                'data-hdocx-id="r-000001" data-hdocx-lock="protected" data-hdocx-part="/word/document.xml" data-hdocx-bold="true">',
            ),
            encoding="utf-8",
            newline="\n",
        )

        plan = plan_hdocx(work)

        self.assertFalse(plan["ok"], plan)
        self.assertEqual(plan["errors"][0]["code"], "PROTECTED_RUN_FORMAT_MODIFIED")
        with self.assertRaises(Exception):
            apply_hdocx(work, output_docx)


    def test_apply_rejects_hdocx_id_tamper(self) -> None:
        input_docx = TMP / "input.docx"
        work = TMP / "work.hdocx"
        output_docx = TMP / "output.docx"
        _make_minimal_docx(input_docx)
        export_docx(input_docx, work)
        html_path = work / "document.html"
        html_path.write_text(
            html_path.read_text(encoding="utf-8").replace('data-hdocx-id="r-000001"', 'data-hdocx-id="r-tampered"'),
            encoding="utf-8",
            newline="\n",
        )

        validation = validate_hdocx(work)

        self.assertFalse(validation["ok"], validation)
        self.assertTrue(any(error["code"] == "HTML_UNKNOWN_HDOCX_ID" for error in validation["errors"]))
        with self.assertRaises(Exception):
            apply_hdocx(work, output_docx)


def _make_minimal_docx(path: Path) -> None:
    content_types = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
</Types>
"""
    rels = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>
"""
    document = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p>
      <w:r>
        <w:t>Hello strict roundtrip.</w:t>
      </w:r>
    </w:p>
    <w:sectPr>
      <w:pgSz w:w="12240" w:h="15840"/>
      <w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/>
    </w:sectPr>
  </w:body>
</w:document>
"""
    with zipfile.ZipFile(path, "w", compression=zipfile.ZIP_DEFLATED) as zf:
        zf.writestr("[Content_Types].xml", content_types)
        zf.writestr("_rels/.rels", rels)
        zf.writestr("word/document.xml", document)


def _make_paragraph_property_docx(path: Path) -> None:
    content_types = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
</Types>
"""
    rels = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>
"""
    document = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p>
      <w:pPr><w:keepNext/><w:ind w:left="720" w:firstLine="240"/><w:spacing w:before="120" w:line="240" w:lineRule="auto"/></w:pPr>
      <w:r><w:t>Paragraph with preserved pPr.</w:t></w:r>
    </w:p>
  </w:body>
</w:document>
"""
    with zipfile.ZipFile(path, "w", compression=zipfile.ZIP_DEFLATED) as zf:
        zf.writestr("[Content_Types].xml", content_types)
        zf.writestr("_rels/.rels", rels)
        zf.writestr("word/document.xml", document)


def _make_audit_feature_docx(path: Path) -> None:
    content_types = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Default Extension="png" ContentType="image/png"/>
  <Default Extension="bin" ContentType="application/vnd.openxmlformats-officedocument.oleObject"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
  <Override PartName="/word/header1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/>
  <Override PartName="/word/footnotes.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml"/>
  <Override PartName="/word/endnotes.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.endnotes+xml"/>
  <Override PartName="/word/comments.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml"/>
  <Override PartName="/word/charts/chart1.xml" ContentType="application/vnd.openxmlformats-officedocument.drawingml.chart+xml"/>
  <Override PartName="/word/diagrams/data1.xml" ContentType="application/vnd.openxmlformats-officedocument.drawingml.diagramData+xml"/>
</Types>
"""
    rels = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>
"""
    document_rels = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header" Target="header1.xml"/>
  <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes" Target="footnotes.xml"/>
  <Relationship Id="rId4" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/endnotes" Target="endnotes.xml"/>
  <Relationship Id="rId5" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments" Target="comments.xml"/>
  <Relationship Id="rId6" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/image1.png"/>
  <Relationship Id="rId7" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart" Target="charts/chart1.xml"/>
  <Relationship Id="rId8" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/diagramData" Target="diagrams/data1.xml"/>
  <Relationship Id="rId9" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/oleObject" Target="embeddings/oleObject1.bin"/>
</Relationships>
"""
    document = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document
  xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
  xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
  xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math"
  xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
  xmlns:v="urn:schemas-microsoft-com:vml">
  <w:body>
    <mc:AlternateContent><mc:Choice Requires="w14"><w:p><w:r><w:t>Choice</w:t></w:r></w:p></mc:Choice><mc:Fallback><w:p><w:r><w:t>Fallback</w:t></w:r></w:p></mc:Fallback></mc:AlternateContent>
    <w:p><w:ins w:id="1"><w:r><w:t>Inserted</w:t></w:r></w:ins></w:p>
    <w:p><w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText>PAGE</w:instrText></w:r></w:p>
    <w:p><m:oMath><m:r><m:t>x=1</m:t></m:r></m:oMath></w:p>
    <w:p><w:pict><v:shape><v:textbox><w:txbxContent><w:p><w:r><w:t>Box</w:t></w:r></w:p></w:txbxContent></v:textbox></v:shape></w:pict></w:p>
    <w:p><w:commentRangeStart w:id="0"/><w:r><w:t>Commented</w:t></w:r><w:commentRangeEnd w:id="0"/><w:r><w:commentReference w:id="0"/></w:r></w:p>
    <w:sectPr><w:headerReference w:type="default" r:id="rId2"/></w:sectPr>
  </w:body>
</w:document>
"""
    header = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:hdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:p><w:r><w:t>Header</w:t></w:r></w:p></w:hdr>
"""
    footnotes = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:footnotes xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:footnote w:id="2"><w:p><w:r><w:t>Footnote</w:t></w:r></w:p></w:footnote></w:footnotes>
"""
    endnotes = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:endnotes xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:endnote w:id="3"><w:p><w:r><w:t>Endnote</w:t></w:r></w:p></w:endnote></w:endnotes>
"""
    comments = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:comments xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:comment w:id="0"><w:p><w:r><w:t>Note</w:t></w:r></w:p></w:comment></w:comments>
"""
    empty_rels = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"/>
"""
    with zipfile.ZipFile(path, "w", compression=zipfile.ZIP_DEFLATED) as zf:
        zf.writestr("[Content_Types].xml", content_types)
        zf.writestr("_rels/.rels", rels)
        zf.writestr("word/_rels/document.xml.rels", document_rels)
        zf.writestr("word/document.xml", document)
        zf.writestr("word/header1.xml", header)
        zf.writestr("word/footnotes.xml", footnotes)
        zf.writestr("word/endnotes.xml", endnotes)
        zf.writestr("word/comments.xml", comments)
        zf.writestr("word/media/image1.png", b"png")
        zf.writestr("word/charts/chart1.xml", '<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"/>')
        zf.writestr("word/diagrams/data1.xml", '<dgm:dataModel xmlns:dgm="http://schemas.openxmlformats.org/drawingml/2006/diagram"/>')
        zf.writestr("word/embeddings/oleObject1.bin", b"ole")
        zf.writestr("customXml/item1.xml", "<root/>")
        zf.writestr("customXml/_rels/item1.xml.rels", empty_rels)


def _make_complex_academic_docx(path: Path) -> None:
    content_types = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Default Extension="png" ContentType="image/png"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
  <Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
  <Override PartName="/word/numbering.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"/>
  <Override PartName="/word/header1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/>
  <Override PartName="/word/footnotes.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml"/>
  <Override PartName="/word/comments.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml"/>
</Types>
"""
    rels = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>
"""
    document_rels = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
  <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering" Target="numbering.xml"/>
  <Relationship Id="rId4" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header" Target="header1.xml"/>
  <Relationship Id="rId5" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes" Target="footnotes.xml"/>
  <Relationship Id="rId6" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments" Target="comments.xml"/>
  <Relationship Id="rId7" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/image1.png"/>
</Relationships>
"""
    document = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document
  xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
  xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
  xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
  xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
  xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math">
  <w:body>
    <w:p>
      <w:pPr><w:pStyle w:val="BodyText"/></w:pPr>
      <w:r><w:t>Academic body text.</w:t></w:r>
      <w:r><w:footnoteReference w:id="2"/></w:r>
    </w:p>
    <w:p>
      <w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr></w:pPr>
      <w:r><w:t>Numbered claim.</w:t></w:r>
    </w:p>
    <w:tbl>
      <w:tr>
        <w:tc><w:p><w:r><w:t>Table A</w:t></w:r></w:p></w:tc>
        <w:tc><w:p><w:r><w:t>Table B</w:t></w:r></w:p></w:tc>
      </w:tr>
    </w:tbl>
    <w:p>
      <w:r>
        <w:drawing>
          <wp:inline>
            <wp:extent cx="914400" cy="457200"/>
            <wp:docPr id="1" name="Picture 1" descr="Figure alt"/>
            <a:graphic><a:graphicData><a:pic><a:blipFill><a:blip r:embed="rId7"/></a:blipFill></a:pic></a:graphicData></a:graphic>
          </wp:inline>
        </w:drawing>
      </w:r>
    </w:p>
    <w:p>
      <w:commentRangeStart w:id="0"/>
      <w:r><w:t>Commented text.</w:t></w:r>
      <w:commentRangeEnd w:id="0"/>
      <w:r><w:commentReference w:id="0"/></w:r>
    </w:p>
    <w:p><w:ins w:id="7" w:author="Reviewer"><w:r><w:t>Inserted text.</w:t></w:r></w:ins></w:p>
    <w:p><m:oMath><m:r><m:t>x=1</m:t></m:r></m:oMath></w:p>
    <w:sectPr><w:headerReference w:type="default" r:id="rId4"/></w:sectPr>
  </w:body>
</w:document>
"""
    styles = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:style w:type="paragraph" w:styleId="BodyText">
    <w:name w:val="Body Text"/>
    <w:pPr><w:spacing w:line="360" w:lineRule="auto"/></w:pPr>
    <w:rPr><w:sz w:val="24"/></w:rPr>
  </w:style>
</w:styles>
"""
    numbering = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:abstractNum w:abstractNumId="0"><w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%1."/></w:lvl></w:abstractNum>
  <w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num>
</w:numbering>
"""
    header = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:hdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:p><w:r><w:t>Header text.</w:t></w:r></w:p></w:hdr>
"""
    footnotes = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:footnotes xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:footnote w:id="2"><w:p><w:r><w:t>Footnote text.</w:t></w:r></w:p></w:footnote></w:footnotes>
"""
    comments = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:comments xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:comment w:id="0" w:author="Reviewer"><w:p><w:r><w:t>Reviewer note.</w:t></w:r></w:p></w:comment></w:comments>
"""
    with zipfile.ZipFile(path, "w", compression=zipfile.ZIP_DEFLATED) as zf:
        zf.writestr("[Content_Types].xml", content_types)
        zf.writestr("_rels/.rels", rels)
        zf.writestr("word/_rels/document.xml.rels", document_rels)
        zf.writestr("word/document.xml", document)
        zf.writestr("word/styles.xml", styles)
        zf.writestr("word/numbering.xml", numbering)
        zf.writestr("word/header1.xml", header)
        zf.writestr("word/footnotes.xml", footnotes)
        zf.writestr("word/comments.xml", comments)
        zf.writestr("word/media/image1.png", b"complexpng")


def _make_protected_run_docx(path: Path) -> None:
    content_types = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
</Types>
"""
    rels = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>
"""
    document = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p>
      <w:r>
        <w:fldChar w:fldCharType="begin"/>
      </w:r>
      <w:r>
        <w:t>Editable text.</w:t>
      </w:r>
    </w:p>
  </w:body>
</w:document>
"""
    with zipfile.ZipFile(path, "w", compression=zipfile.ZIP_DEFLATED) as zf:
        zf.writestr("[Content_Types].xml", content_types)
        zf.writestr("_rels/.rels", rels)
        zf.writestr("word/document.xml", document)


def _make_table_docx(path: Path) -> None:
    content_types = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
</Types>
"""
    rels = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>
"""
    document = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:tbl>
      <w:tr>
        <w:tc>
          <w:p>
            <w:r>
              <w:t>Cell text.</w:t>
            </w:r>
          </w:p>
        </w:tc>
      </w:tr>
    </w:tbl>
  </w:body>
</w:document>
"""
    with zipfile.ZipFile(path, "w", compression=zipfile.ZIP_DEFLATED) as zf:
        zf.writestr("[Content_Types].xml", content_types)
        zf.writestr("_rels/.rels", rels)
        zf.writestr("word/document.xml", document)


def _make_two_cell_table_docx(path: Path) -> None:
    content_types = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
</Types>
"""
    rels = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>
"""
    document = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:tbl>
      <w:tr>
        <w:tc><w:tcPr><w:tcW w:w="2400" w:type="dxa"/></w:tcPr><w:p><w:r><w:t>A1</w:t></w:r></w:p></w:tc>
        <w:tc><w:tcPr><w:tcW w:w="2400" w:type="dxa"/></w:tcPr><w:p><w:r><w:t>B1</w:t></w:r></w:p></w:tc>
      </w:tr>
    </w:tbl>
  </w:body>
</w:document>
"""
    with zipfile.ZipFile(path, "w", compression=zipfile.ZIP_DEFLATED) as zf:
        zf.writestr("[Content_Types].xml", content_types)
        zf.writestr("_rels/.rels", rels)
        zf.writestr("word/document.xml", document)


def _make_merged_table_docx(path: Path) -> None:
    content_types = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
</Types>
"""
    rels = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>
"""
    document = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:tbl>
      <w:tr>
        <w:tc>
          <w:tcPr><w:gridSpan w:val="2"/></w:tcPr>
          <w:p><w:r><w:t>Merged cell</w:t></w:r></w:p>
        </w:tc>
      </w:tr>
    </w:tbl>
  </w:body>
</w:document>
"""
    with zipfile.ZipFile(path, "w", compression=zipfile.ZIP_DEFLATED) as zf:
        zf.writestr("[Content_Types].xml", content_types)
        zf.writestr("_rels/.rels", rels)
        zf.writestr("word/document.xml", document)


def _make_two_row_table_docx(path: Path) -> None:
    content_types = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
</Types>
"""
    rels = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>
"""
    document = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:tbl>
      <w:tr>
        <w:tc><w:p><w:r><w:t>A1</w:t></w:r></w:p></w:tc>
        <w:tc><w:p><w:r><w:t>B1</w:t></w:r></w:p></w:tc>
      </w:tr>
      <w:tr>
        <w:tc><w:p><w:r><w:t>A2</w:t></w:r></w:p></w:tc>
        <w:tc><w:p><w:r><w:t>B2</w:t></w:r></w:p></w:tc>
      </w:tr>
    </w:tbl>
  </w:body>
</w:document>
"""
    with zipfile.ZipFile(path, "w", compression=zipfile.ZIP_DEFLATED) as zf:
        zf.writestr("[Content_Types].xml", content_types)
        zf.writestr("_rels/.rels", rels)
        zf.writestr("word/document.xml", document)


def _make_header_docx(path: Path) -> None:
    content_types = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
  <Override PartName="/word/header1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/>
</Types>
"""
    rels = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>
"""
    document_rels = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header" Target="header1.xml"/>
</Relationships>
"""
    document = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
  xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <w:body>
    <w:p><w:r><w:t>Main text.</w:t></w:r></w:p>
    <w:sectPr>
      <w:headerReference w:type="default" r:id="rId2"/>
    </w:sectPr>
  </w:body>
</w:document>
"""
    header = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:hdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:p>
    <w:r>
      <w:t>Header text.</w:t>
    </w:r>
  </w:p>
</w:hdr>
"""
    with zipfile.ZipFile(path, "w", compression=zipfile.ZIP_DEFLATED) as zf:
        zf.writestr("[Content_Types].xml", content_types)
        zf.writestr("_rels/.rels", rels)
        zf.writestr("word/_rels/document.xml.rels", document_rels)
        zf.writestr("word/document.xml", document)
        zf.writestr("word/header1.xml", header)


def _make_footnote_docx(path: Path) -> None:
    content_types = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
  <Override PartName="/word/footnotes.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml"/>
</Types>
"""
    rels = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>
"""
    document_rels = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes" Target="footnotes.xml"/>
</Relationships>
"""
    document = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
  xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <w:body>
    <w:p>
      <w:r><w:t>Main text</w:t></w:r>
      <w:r><w:footnoteReference w:id="2"/></w:r>
    </w:p>
  </w:body>
</w:document>
"""
    footnotes = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:footnotes xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:footnote w:id="2">
    <w:p>
      <w:r>
        <w:t>Footnote text.</w:t>
      </w:r>
    </w:p>
  </w:footnote>
</w:footnotes>
"""
    with zipfile.ZipFile(path, "w", compression=zipfile.ZIP_DEFLATED) as zf:
        zf.writestr("[Content_Types].xml", content_types)
        zf.writestr("_rels/.rels", rels)
        zf.writestr("word/_rels/document.xml.rels", document_rels)
        zf.writestr("word/document.xml", document)
        zf.writestr("word/footnotes.xml", footnotes)


def _make_styled_docx(path: Path) -> None:
    content_types = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
  <Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
</Types>
"""
    rels = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>
"""
    document_rels = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
</Relationships>
"""
    document = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p>
      <w:pPr><w:pStyle w:val="Normal"/></w:pPr>
      <w:r><w:t>Styled text.</w:t></w:r>
    </w:p>
  </w:body>
</w:document>
"""
    styles = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:style w:type="paragraph" w:styleId="Normal">
    <w:name w:val="Normal"/>
  </w:style>
  <w:style w:type="paragraph" w:styleId="Heading1">
    <w:name w:val="Heading 1"/>
    <w:basedOn w:val="Normal"/>
    <w:qFormat/>
    <w:rPr><w:b/><w:sz w:val="32"/></w:rPr>
  </w:style>
</w:styles>
"""
    with zipfile.ZipFile(path, "w", compression=zipfile.ZIP_DEFLATED) as zf:
        zf.writestr("[Content_Types].xml", content_types)
        zf.writestr("_rels/.rels", rels)
        zf.writestr("word/_rels/document.xml.rels", document_rels)
        zf.writestr("word/document.xml", document)
        zf.writestr("word/styles.xml", styles)


def _make_comment_docx(path: Path) -> None:
    content_types = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
  <Override PartName="/word/comments.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml"/>
</Types>
"""
    rels = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>
"""
    document_rels = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments" Target="comments.xml"/>
</Relationships>
"""
    document = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p>
      <w:commentRangeStart w:id="0"/>
      <w:r><w:t>Commented text.</w:t></w:r>
      <w:commentRangeEnd w:id="0"/>
      <w:r><w:commentReference w:id="0"/></w:r>
    </w:p>
  </w:body>
</w:document>
"""
    comments = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:comments xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:comment w:id="0" w:author="Reviewer">
    <w:p>
      <w:r><w:t>Reviewer note.</w:t></w:r>
    </w:p>
  </w:comment>
</w:comments>
"""
    with zipfile.ZipFile(path, "w", compression=zipfile.ZIP_DEFLATED) as zf:
        zf.writestr("[Content_Types].xml", content_types)
        zf.writestr("_rels/.rels", rels)
        zf.writestr("word/_rels/document.xml.rels", document_rels)
        zf.writestr("word/document.xml", document)
        zf.writestr("word/comments.xml", comments)


def _make_numbered_docx(path: Path) -> None:
    content_types = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
  <Override PartName="/word/numbering.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"/>
</Types>
"""
    rels = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>
"""
    document_rels = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering" Target="numbering.xml"/>
</Relationships>
"""
    document = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p>
      <w:pPr>
        <w:numPr>
          <w:ilvl w:val="0"/>
          <w:numId w:val="1"/>
        </w:numPr>
      </w:pPr>
      <w:r><w:t>Numbered item.</w:t></w:r>
    </w:p>
  </w:body>
</w:document>
"""
    numbering = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:abstractNum w:abstractNumId="0">
    <w:lvl w:ilvl="0">
      <w:start w:val="1"/>
      <w:numFmt w:val="decimal"/>
      <w:lvlText w:val="%1."/>
      <w:suff w:val="tab"/>
      <w:pPr><w:ind w:left="720" w:hanging="360"/></w:pPr>
    </w:lvl>
  </w:abstractNum>
  <w:num w:numId="1">
    <w:abstractNumId w:val="0"/>
  </w:num>
</w:numbering>
"""
    with zipfile.ZipFile(path, "w", compression=zipfile.ZIP_DEFLATED) as zf:
        zf.writestr("[Content_Types].xml", content_types)
        zf.writestr("_rels/.rels", rels)
        zf.writestr("word/_rels/document.xml.rels", document_rels)
        zf.writestr("word/document.xml", document)
        zf.writestr("word/numbering.xml", numbering)


def _make_image_docx(path: Path, *, exact_line_spacing: bool = False) -> None:
    content_types = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Default Extension="png" ContentType="image/png"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
</Types>
"""
    rels = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>
"""
    document_rels = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/image1.png"/>
</Relationships>
"""
    paragraph_properties = '<w:pPr><w:spacing w:line="240" w:lineRule="exact"/></w:pPr>' if exact_line_spacing else ""
    document = f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document
  xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
  xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
  xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
  xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
  <w:body>
    <w:p>
      {paragraph_properties}
      <w:r>
        <w:drawing>
          <wp:inline>
            <wp:extent cx="914400" cy="457200"/>
            <wp:docPr id="1" name="Picture 1" descr="Old alt"/>
            <a:graphic>
              <a:graphicData>
                <a:pic>
                  <a:blipFill>
                    <a:blip r:embed="rId2"/>
                  </a:blipFill>
                </a:pic>
              </a:graphicData>
            </a:graphic>
          </wp:inline>
        </w:drawing>
      </w:r>
    </w:p>
  </w:body>
</w:document>
"""
    with zipfile.ZipFile(path, "w", compression=zipfile.ZIP_DEFLATED) as zf:
        zf.writestr("[Content_Types].xml", content_types)
        zf.writestr("_rels/.rels", rels)
        zf.writestr("word/_rels/document.xml.rels", document_rels)
        zf.writestr("word/document.xml", document)
        zf.writestr("word/media/image1.png", b"fakepng")


def _make_query_docx(path: Path) -> None:
    content_types = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Default Extension="png" ContentType="image/png"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
</Types>
"""
    rels = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>
"""
    document_rels = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/image1.png"/>
</Relationships>
"""
    document = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document
  xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
  xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
  xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
  xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
  <w:body>
    <w:p>
      <w:pPr><w:spacing w:line="240" w:lineRule="exact"/></w:pPr>
    </w:p>
    <w:p>
      <w:pPr><w:jc w:val="center"/></w:pPr>
      <w:r>
        <w:rPr><w:b/><w:rFonts w:ascii="SimHei" w:hAnsi="SimHei" w:eastAsia="SimHei"/><w:sz w:val="28"/></w:rPr>
        <w:t>1 First Heading</w:t>
      </w:r>
    </w:p>
    <w:p>
      <w:r><w:t>Keywords: artificial intelligence</w:t></w:r>
    </w:p>
    <w:p>
      <w:pPr><w:jc w:val="center"/><w:spacing w:line="240" w:lineRule="auto"/></w:pPr>
      <w:r>
        <w:drawing>
          <wp:inline>
            <wp:extent cx="914400" cy="457200"/>
            <wp:docPr id="1" name="Picture 1" descr="Query image"/>
            <a:graphic>
              <a:graphicData>
                <a:pic>
                  <a:blipFill>
                    <a:blip r:embed="rId2"/>
                  </a:blipFill>
                </a:pic>
              </a:graphicData>
            </a:graphic>
          </wp:inline>
        </w:drawing>
      </w:r>
    </w:p>
  </w:body>
</w:document>
"""
    with zipfile.ZipFile(path, "w", compression=zipfile.ZIP_DEFLATED) as zf:
        zf.writestr("[Content_Types].xml", content_types)
        zf.writestr("_rels/.rels", rels)
        zf.writestr("word/_rels/document.xml.rels", document_rels)
        zf.writestr("word/document.xml", document)
        zf.writestr("word/media/image1.png", b"fakepng")


def _make_preface_heading_docx(path: Path) -> None:
    content_types = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
</Types>
"""
    rels = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>
"""
    abstract_title = "\u6458 \u8981"
    document = f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p>
      <w:pPr><w:jc w:val="center"/></w:pPr>
      <w:r>
        <w:rPr><w:b/><w:rFonts w:ascii="SimHei" w:hAnsi="SimHei" w:eastAsia="SimHei"/><w:sz w:val="32"/></w:rPr>
        <w:t>Paper Title</w:t>
      </w:r>
    </w:p>
    <w:p>
      <w:pPr><w:jc w:val="center"/></w:pPr>
      <w:r>
        <w:rPr><w:b/><w:rFonts w:ascii="SimHei" w:hAnsi="SimHei" w:eastAsia="SimHei"/><w:sz w:val="28"/></w:rPr>
        <w:t>{abstract_title}</w:t>
      </w:r>
    </w:p>
    <w:p/>
    <w:p>
      <w:pPr><w:jc w:val="center"/></w:pPr>
      <w:r>
        <w:rPr><w:b/><w:rFonts w:ascii="SimHei" w:hAnsi="SimHei" w:eastAsia="SimHei"/><w:sz w:val="28"/></w:rPr>
        <w:t>1 Introduction</w:t>
      </w:r>
    </w:p>
  </w:body>
</w:document>
"""
    with zipfile.ZipFile(path, "w", compression=zipfile.ZIP_DEFLATED) as zf:
        zf.writestr("[Content_Types].xml", content_types)
        zf.writestr("_rels/.rels", rels)
        zf.writestr("word/document.xml", document)


def _make_equation_docx(path: Path) -> None:
    content_types = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
</Types>
"""
    rels = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>
"""
    document = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document
  xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
  xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math">
  <w:body>
    <w:p>
      <m:oMath><m:r><m:t>x=1</m:t></m:r></m:oMath>
    </w:p>
  </w:body>
</w:document>
"""
    with zipfile.ZipFile(path, "w", compression=zipfile.ZIP_DEFLATED) as zf:
        zf.writestr("[Content_Types].xml", content_types)
        zf.writestr("_rels/.rels", rels)
        zf.writestr("word/document.xml", document)


def _make_revision_docx(path: Path) -> None:
    content_types = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
</Types>
"""
    rels = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>
"""
    document = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p>
      <w:ins w:id="7" w:author="Reviewer">
        <w:r><w:t>Inserted text.</w:t></w:r>
      </w:ins>
    </w:p>
  </w:body>
</w:document>
"""
    with zipfile.ZipFile(path, "w", compression=zipfile.ZIP_DEFLATED) as zf:
        zf.writestr("[Content_Types].xml", content_types)
        zf.writestr("_rels/.rels", rels)
        zf.writestr("word/document.xml", document)


if __name__ == "__main__":
    unittest.main()
