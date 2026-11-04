import os
import tempfile
import unittest
from contextlib import redirect_stdout
from io import StringIO

from docx import Document
from lxml import etree

import docx_reader
from app import main

ROOT = os.path.join(os.path.dirname(__file__), "..")
SCHEMA = os.path.join(ROOT, "schema.xsd")


class MainTest(unittest.TestCase):
    def setUp(self) -> None:
        docx_reader.warnings.clear()
        tmp = tempfile.TemporaryDirectory()
        self.addCleanup(tmp.cleanup)
        self.input = os.path.join(tmp.name, "in.docx")
        self.output = os.path.join(tmp.name, "out.xml")
        doc = Document()
        doc.add_heading("Title", 1)
        doc.add_paragraph("Text", style="Quote")
        doc.save(self.input)

    def run_main(self, *args: str) -> str:
        buffer = StringIO()
        with redirect_stdout(buffer):
            main(["app.py", *args])
        return buffer.getvalue()

    def assert_output_is_valid(self) -> None:
        schema = etree.XMLSchema(etree.parse(SCHEMA))
        doc = etree.parse(self.output)
        self.assertTrue(schema.validate(doc), schema.error_log)

    def test_too_few_arguments_exits(self) -> None:
        with self.assertRaises(SystemExit):
            main(["app.py", self.input])

    def test_writes_xml_without_schema(self) -> None:
        out = self.run_main(self.input, self.output)
        root = etree.parse(self.output).getroot()
        self.assertEqual(root.tag, "document")
        self.assertEqual(root.findtext("section/title"), "Title")
        self.assertIn("schema-valide: None", out)

    def test_writes_utf8_declaration(self) -> None:
        self.run_main(self.input, self.output)
        with open(self.output, "rb") as f:
            self.assertIn(b"encoding='UTF-8'", f.readline())

    def test_validates_with_schema(self) -> None:
        out = self.run_main(self.input, self.output, SCHEMA)
        self.assertIn("schema-valide: True", out)
        self.assert_output_is_valid()

    def test_prints_warnings(self) -> None:
        out = self.run_main(self.input, self.output)
        self.assertIn("WARNUNG:", out)
        self.assertIn("Quote", out)

    def test_sample_documents_are_schema_valid(self) -> None:
        for name in ("test.docx", "test2.docx"):
            with self.subTest(name):
                sample = os.path.join(ROOT, "input", name)
                out = self.run_main(sample, self.output, SCHEMA)
                self.assertIn("schema-valide: True", out)
                self.assert_output_is_valid()


if __name__ == "__main__":
    unittest.main()
