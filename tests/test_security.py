from __future__ import annotations

import tempfile
import unittest
import zipfile
from pathlib import Path

from corporate_report_compiler.diagnostics import Diagnostics
from corporate_report_compiler.security import inspect_docx_zip

from tests.support import build_valid_docx, rewrite_zip


class DocxSecurityTests(unittest.TestCase):
    def setUp(self) -> None:
        self.temp_directory = tempfile.TemporaryDirectory()
        self.root = Path(self.temp_directory.name)
        self.valid = build_valid_docx(self.root / "valid.docx")

    def tearDown(self) -> None:
        self.temp_directory.cleanup()

    def inspect(self, path: Path) -> Diagnostics:
        diagnostics = Diagnostics()
        inspect_docx_zip(path, diagnostics)
        return diagnostics

    def test_valid_docx_passes_preinspection(self) -> None:
        diagnostics = self.inspect(self.valid)
        self.assertFalse(diagnostics.has_errors, diagnostics.to_dict())

    def test_wrong_extension_and_non_zip_are_rejected(self) -> None:
        renamed = self.root / "template.docm"
        renamed.write_bytes(self.valid.read_bytes())
        self.assertTrue(self.inspect(renamed).has_errors)

        bogus = self.root / "bogus.docx"
        bogus.write_text("not a zip", encoding="utf-8")
        self.assertTrue(self.inspect(bogus).has_errors)

    def test_macro_active_x_embedding_and_media_are_rejected(self) -> None:
        for index, part in enumerate(
            (
                "word/vbaProject.bin",
                "word/activeX/activeX1.bin",
                "word/embeddings/oleObject1.bin",
                "word/media/image1.png",
                "word/charts/chart1.xml",
            )
        ):
            with self.subTest(part=part):
                hostile = rewrite_zip(
                    self.valid,
                    self.root / f"hostile-{index}.docx",
                    additions={part: b"hostile"},
                )
                diagnostics = self.inspect(hostile)
                self.assertTrue(diagnostics.has_errors)
                self.assertIn("OOXML-001", {item.code for item in diagnostics.items})

    def test_custom_xml_is_rejected_when_explicitly_present(self) -> None:
        hostile = rewrite_zip(
            self.valid,
            self.root / "custom-xml.docx",
            additions={"customXml/item1.xml": b"<root><secret>value</secret></root>"},
        )
        diagnostics = self.inspect(hostile)
        self.assertTrue(diagnostics.has_errors)
        self.assertIn("OOXML-001", {item.code for item in diagnostics.items})

    def test_external_relationship_is_rejected(self) -> None:
        relationships = b'''<?xml version="1.0" encoding="UTF-8"?>
        <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
          <Relationship Id="rId99" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink"
             Target="https://attacker.invalid/leak" TargetMode="External"/>
        </Relationships>'''
        hostile = rewrite_zip(
            self.valid,
            self.root / "external.docx",
            replacements={"word/_rels/document.xml.rels": relationships},
        )
        diagnostics = self.inspect(hostile)
        self.assertTrue(diagnostics.has_errors)
        self.assertIn("OOXML-004", {item.code for item in diagnostics.items})

    def test_doctype_and_entity_are_rejected_without_resolution(self) -> None:
        with zipfile.ZipFile(self.valid) as archive:
            document_xml = archive.read("word/document.xml")
        declaration_end = document_xml.find(b"?>") + 2
        hostile_xml = (
            document_xml[:declaration_end]
            + b'<!DOCTYPE w:document [<!ENTITY xxe SYSTEM "file:///etc/passwd">]>'
            + document_xml[declaration_end:]
        )
        hostile = rewrite_zip(
            self.valid,
            self.root / "xxe.docx",
            replacements={"word/document.xml": hostile_xml},
        )
        diagnostics = self.inspect(hostile)
        self.assertTrue(diagnostics.has_errors)
        self.assertIn("OOXML-002", {item.code for item in diagnostics.items})

    def test_doctype_is_detected_anywhere_in_complete_xml_part(self) -> None:
        """The security decision must not depend on a 4 KiB prefix probe."""
        with zipfile.ZipFile(self.valid) as archive:
            document_xml = archive.read("word/document.xml")
        declaration_end = document_xml.find(b"?>") + 2
        hostile_xml = (
            document_xml[:declaration_end]
            + (b" " * 5000)
            + b'<!DOCTYPE w:document [<!ENTITY safe "expanded">]>'
            + document_xml[declaration_end:]
        )
        hostile = rewrite_zip(
            self.valid,
            self.root / "late-doctype.docx",
            replacements={"word/document.xml": hostile_xml},
        )
        diagnostics = self.inspect(hostile)
        self.assertTrue(diagnostics.has_errors)
        self.assertIn("OOXML-002", {item.code for item in diagnostics.items})

    def test_path_traversal_and_backslash_are_rejected(self) -> None:
        for index, name in enumerate(("../evil.xml", "/absolute.xml", "word\\evil.xml")):
            with self.subTest(name=name):
                hostile = rewrite_zip(
                    self.valid,
                    self.root / f"traversal-{index}.docx",
                    additions={name: b"<x/>",},
                )
                diagnostics = self.inspect(hostile)
                self.assertTrue(diagnostics.has_errors)
                self.assertIn("ZIP-003", {item.code for item in diagnostics.items})

    def test_duplicate_casefolded_entry_is_rejected(self) -> None:
        hostile = self.root / "duplicate.docx"
        with zipfile.ZipFile(self.valid) as source:
            with zipfile.ZipFile(hostile, "w", zipfile.ZIP_DEFLATED) as target:
                for info in source.infolist():
                    target.writestr(info, source.read(info.filename))
                target.writestr("WORD/DOCUMENT.XML", b"<x/>")
        diagnostics = self.inspect(hostile)
        self.assertTrue(diagnostics.has_errors)
        self.assertIn("ZIP-004", {item.code for item in diagnostics.items})

    def test_missing_required_ooxml_parts_are_rejected(self) -> None:
        for index, missing in enumerate(("[Content_Types].xml", "word/document.xml")):
            with self.subTest(missing=missing):
                hostile = rewrite_zip(
                    self.valid,
                    self.root / f"missing-{index}.docx",
                    omit={missing},
                )
                diagnostics = self.inspect(hostile)
                self.assertTrue(
                    diagnostics.has_errors,
                    f"Un DOCX sin {missing} no debe aprobar la preinspección.",
                )

    def test_macro_enabled_content_type_is_rejected(self) -> None:
        with zipfile.ZipFile(self.valid) as archive:
            content_types = archive.read("[Content_Types].xml")
        macro_types = content_types.replace(
            b"application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml",
            b"application/vnd.ms-word.document.macroEnabled.main+xml",
        )
        hostile = rewrite_zip(
            self.valid,
            self.root / "macro-content-type.docx",
            replacements={"[Content_Types].xml": macro_types},
        )
        diagnostics = self.inspect(hostile)
        self.assertTrue(diagnostics.has_errors)
        self.assertIn("OOXML-005", {item.code for item in diagnostics.items})


if __name__ == "__main__":
    unittest.main()
