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

    def test_gueltig(self) -> None:
        root = etree.fromstring("<document><p>x</p></document>")
        self.assertTrue(validate(root, self.xsd))
        self.assertEqual(docx_reader.warnings, [])

    def test_ungueltig_meldet_schemafehler(self) -> None:
        root = etree.fromstring("<document><q/></document>")
        self.assertFalse(validate(root, self.xsd))
        self.assertTrue(docx_reader.warnings)
        self.assertTrue(docx_reader.warnings[0].startswith("Schemafehler Zeile"))

    def test_echtes_schema(self) -> None:
        schema = os.path.join(os.path.dirname(__file__), "..", "schema.xsd")
        root = etree.fromstring("<document><section><title>T</title></section></document>")
        self.assertTrue(validate(root, schema))

    def test_echtes_schema_neue_formate(self) -> None:
        schema = os.path.join(os.path.dirname(__file__), "..", "schema.xsd")
        root = etree.fromstring(
            "<document><p>a <b><u>b</u></b> H<sub>2</sub>O</p></document>")
        self.assertTrue(validate(root, schema))

if __name__ == "__main__":
    unittest.main()
