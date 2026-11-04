import unittest

from docx import Document
from lxml import etree

import docx_reader
from docx_converter import DocxConverter


def xml_to_string(el: etree._Element) -> str:
    return etree.tostring(el, encoding="unicode")


class DocxConverterTest(unittest.TestCase):
    def setUp(self) -> None:
        docx_reader.warnings.clear()
        self.doc = Document()

    def convert(self) -> str:
        return xml_to_string(DocxConverter().convert(self.doc))

    def test_empty_document(self) -> None:
        self.assertEqual(self.convert(), "<document/>")

    def test_plain_paragraph(self) -> None:
        self.doc.add_paragraph("Text")
        self.assertEqual(self.convert(), "<document><p>Text</p></document>")
        self.assertEqual(docx_reader.warnings, [])

    def test_empty_and_whitespace_paragraphs_are_ignored(self) -> None:
        self.doc.add_paragraph("")
        self.doc.add_paragraph("   ")
        self.assertEqual(self.convert(), "<document/>")

    def test_inline_formatting_in_paragraph(self) -> None:
        p = self.doc.add_paragraph()
        p.add_run("a ")
        p.add_run("b").bold = True
        self.assertEqual(self.convert(), "<document><p>a <b>b</b></p></document>")

    def test_nested_headings(self) -> None:
        self.doc.add_heading("H1", 1)
        self.doc.add_heading("H2", 2)
        self.doc.add_heading("H3", 3)
        self.assertEqual(
            self.convert(),
            "<document><section><title>H1</title><section><title>H2</title>"
            "<section><title>H3</title></section></section></section></document>")

    def test_heading_closes_deeper_sections(self) -> None:
        self.doc.add_heading("A", 1)
        self.doc.add_heading("B", 2)
        self.doc.add_heading("C", 1)
        self.doc.add_paragraph("x")
        self.assertEqual(
            self.convert(),
            "<document><section><title>A</title><section><title>B</title></section>"
            "</section><section><title>C</title><p>x</p></section></document>")

    def test_heading_closes_sections_of_same_level(self) -> None:
        self.doc.add_heading("A", 2)
        self.doc.add_heading("B", 2)
        self.assertEqual(
            self.convert(),
            "<document><section><title>A</title></section>"
            "<section><title>B</title></section></document>")

    def test_skipped_heading_level(self) -> None:
        self.doc.add_heading("A", 1)
        self.doc.add_heading("C", 3)
        self.doc.add_paragraph("x")
        self.assertEqual(
            self.convert(),
            "<document><section><title>A</title><section><title>C</title>"
            "<p>x</p></section></section></document>")

    def test_heading_title_is_plain(self) -> None:
        h = self.doc.add_heading("", 1)
        h.add_run("bold").bold = True
        self.assertEqual(
            self.convert(), "<document><section><title>bold</title></section></document>")

    def test_paragraph_goes_into_innermost_section(self) -> None:
        self.doc.add_paragraph("before")
        self.doc.add_heading("H", 1)
        self.doc.add_paragraph("after")
        self.assertEqual(
            self.convert(),
            "<document><p>before</p><section><title>H</title><p>after</p></section></document>")

    def test_container_property(self) -> None:
        converter = DocxConverter()
        self.assertIs(converter.container, converter.root)
        self.doc.add_heading("H", 1)
        converter.convert(self.doc)
        self.assertEqual(converter.container.tag, "section")

    def test_list_items_are_grouped(self) -> None:
        self.doc.add_paragraph("a", style="List Bullet")
        self.doc.add_paragraph("b", style="List Bullet")
        self.assertEqual(
            self.convert(), "<document><ul><li>a</li><li>b</li></ul></document>")

    def test_list_type_change_starts_new_list(self) -> None:
        self.doc.add_paragraph("a", style="List Bullet")
        self.doc.add_paragraph("b", style="List Number")
        self.assertEqual(
            self.convert(),
            "<document><ul><li>a</li></ul><ol><li>b</li></ol></document>")

    def test_paragraph_ends_list(self) -> None:
        self.doc.add_paragraph("a", style="List Bullet")
        self.doc.add_paragraph("p")
        self.doc.add_paragraph("b", style="List Bullet")
        self.assertEqual(
            self.convert(),
            "<document><ul><li>a</li></ul><p>p</p><ul><li>b</li></ul></document>")

    def test_heading_ends_list(self) -> None:
        self.doc.add_paragraph("a", style="List Bullet")
        self.doc.add_heading("H", 1)
        self.doc.add_paragraph("b", style="List Bullet")
        self.assertEqual(
            self.convert(),
            "<document><ul><li>a</li></ul><section><title>H</title>"
            "<ul><li>b</li></ul></section></document>")

    def test_table_ends_list(self) -> None:
        self.doc.add_paragraph("a", style="List Bullet")
        self.doc.add_table(rows=1, cols=1)
        self.doc.add_paragraph("b", style="List Bullet")
        root = DocxConverter().convert(self.doc)
        self.assertEqual([c.tag for c in root], ["ul", "table", "ul"])

    def test_empty_paragraph_does_not_end_list(self) -> None:
        self.doc.add_paragraph("a", style="List Bullet")
        self.doc.add_paragraph("")
        self.doc.add_paragraph("b", style="List Bullet")
        self.assertEqual(
            self.convert(), "<document><ul><li>a</li><li>b</li></ul></document>")

    def test_table_goes_into_current_section(self) -> None:
        self.doc.add_heading("H", 1)
        table = self.doc.add_table(rows=1, cols=1)
        table.cell(0, 0).text = "c"
        self.assertEqual(
            self.convert(),
            '<document><section><title>H</title><table><row header="true">'
            "<entry>c</entry></row></table></section></document>")

    def test_unknown_style_warns_but_keeps_paragraph(self) -> None:
        self.doc.add_paragraph("q", style="Quote")
        self.assertEqual(self.convert(), "<document><p>q</p></document>")
        self.assertEqual(len(docx_reader.warnings), 1)
        self.assertIn("Quote", docx_reader.warnings[0])

    def test_body_text_style_does_not_warn(self) -> None:
        self.doc.add_paragraph("q", style="Body Text")
        self.convert()
        self.assertEqual(docx_reader.warnings, [])

    def test_bold_heuristic_creates_section(self) -> None:
        p = self.doc.add_paragraph()
        p.add_run("Looks like heading").bold = True
        self.assertEqual(
            self.convert(),
            "<document><section><title>Looks like heading</title></section></document>")


if __name__ == "__main__":
    unittest.main()
