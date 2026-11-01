import os
import tempfile
import unittest

from lxml import etree

import docx_reader
from xml_validator import validate

XSD = """<xs:schema xmlns:xs="http://www.w3.org/2001/XMLSchema">
  <xs:element name="document">
    <xs:complexType><xs:sequence>
      <xs:element name="p" type="xs:string" maxOccurs="unbounded"/>
    </xs:sequence></xs:complexType>
  </xs:element>
</xs:schema>"""


class ValidateTest(unittest.TestCase):
    def setUp(self) -> None:
        docx_reader.warnings.clear()
        fd, self.xsd = tempfile.mkstemp(suffix=".xsd")
        with os.fdopen(fd, "w", encoding="utf-8") as f:
            f.write(XSD)
        self.addCleanup(os.remove, self.xsd)

    def test_valid(self) -> None:
        root = etree.fromstring("<document><p>x</p></document>")
        self.assertTrue(validate(root, self.xsd))
        self.assertEqual(docx_reader.warnings, [])

    def test_invalid_reports_schema_error(self) -> None:
        root = etree.fromstring("<document><q/></document>")
        self.assertFalse(validate(root, self.xsd))
        self.assertTrue(docx_reader.warnings)
        self.assertTrue(docx_reader.warnings[0].startswith("Schemafehler Zeile"))

    def test_real_schema(self) -> None:
        schema = os.path.join(os.path.dirname(__file__), "..", "schema.xsd")
        root = etree.fromstring("<document><section><title>T</title></section></document>")
        self.assertTrue(validate(root, schema))

    def test_real_schema_new_formats(self) -> None:
        schema = os.path.join(os.path.dirname(__file__), "..", "schema.xsd")
        root = etree.fromstring(
            "<document><p>a <b><u>b</u></b> H<sub>2</sub>O</p></document>")
        self.assertTrue(validate(root, schema))

    def real_schema(self) -> str:
        return os.path.join(os.path.dirname(__file__), "..", "schema.xsd")

    def test_real_schema_lists_and_tables(self) -> None:
        root = etree.fromstring(
            "<document><ul><li>a</li></ul><ol><li>b</li></ol>"
            "<table><row header=\"true\"><entry>x</entry></row></table></document>")
        self.assertTrue(validate(root, self.real_schema()))

    def test_real_schema_rejects_unknown_tag(self) -> None:
        root = etree.fromstring("<document><foo/></document>")
        self.assertFalse(validate(root, self.real_schema()))

    def test_real_schema_rejects_paragraph_in_table(self) -> None:
        root = etree.fromstring(
            "<document><table><row><p>x</p></row></table></document>")
        self.assertFalse(validate(root, self.real_schema()))

    def test_real_schema_rejects_section_without_title(self) -> None:
        root = etree.fromstring("<document><section><p>x</p></section></document>")
        self.assertFalse(validate(root, self.real_schema()))

    def test_error_contains_line_number(self) -> None:
        root = etree.fromstring("<document>\n<q/>\n</document>")
        validate(root, self.xsd)
        self.assertRegex(docx_reader.warnings[0], r"Schemafehler Zeile \d+: ")

    def test_multiple_errors_give_multiple_warnings(self) -> None:
        root = etree.fromstring(
            "<document><table><row><p/></row></table>"
            "<table><row><p/></row></table></document>")
        validate(root, self.real_schema())
        self.assertGreaterEqual(len(docx_reader.warnings), 2)


if __name__ == "__main__":
    unittest.main()
