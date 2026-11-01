import unittest

from docx import Document
from docx.table import Table
from docx.text.paragraph import Paragraph

import docx_reader
from docx_reader import heading_level, iter_blocks, list_tag_for


class IterBlocksTest(unittest.TestCase):
    def test_order_and_types(self) -> None:
        doc = Document()
        doc.add_paragraph("One")
        doc.add_table(rows=1, cols=1)
        doc.add_paragraph("Two")
        blocks = list(iter_blocks(doc))
        self.assertEqual([type(b) for b in blocks], [Paragraph, Table, Paragraph])
        self.assertEqual(blocks[2].text, "Two")

    def test_empty_document(self) -> None:
        self.assertEqual(list(iter_blocks(Document())), [])


class HeadingLevelTest(unittest.TestCase):
    def setUp(self) -> None:
        docx_reader.warnings.clear()
        self.doc = Document()

    def test_heading_style(self) -> None:
        self.assertEqual(heading_level(self.doc.add_heading("A", 3)), 3)

    def test_title_is_level_1(self) -> None:
        self.assertEqual(heading_level(self.doc.add_heading("A", 0)), 1)

    def test_heuristic_short_bold_paragraph(self) -> None:
        p = self.doc.add_paragraph()
        p.add_run("Short and bold").bold = True
        self.assertEqual(heading_level(p), 2)
        self.assertEqual(len(docx_reader.warnings), 1)

    def test_heuristic_not_applied_with_sentence_end(self) -> None:
        p = self.doc.add_paragraph()
        p.add_run("A sentence.").bold = True
        self.assertIsNone(heading_level(p))

    def test_heuristic_not_applied_to_long_text(self) -> None:
        p = self.doc.add_paragraph()
        p.add_run("x" * 80).bold = True
        self.assertIsNone(heading_level(p))

    def test_heuristic_not_applied_when_partly_bold(self) -> None:
        p = self.doc.add_paragraph()
        p.add_run("bold").bold = True
        p.add_run(" normal text")
        self.assertIsNone(heading_level(p))

    def test_normal_paragraph(self) -> None:
        self.assertIsNone(heading_level(self.doc.add_paragraph("Text")))
        self.assertEqual(docx_reader.warnings, [])


class ListTagForTest(unittest.TestCase):
    def test_mapping(self) -> None:
        self.assertEqual(list_tag_for("List Bullet"), "ul")
        self.assertEqual(list_tag_for("List Bullet 2"), "ul")
        self.assertEqual(list_tag_for("List Number"), "ol")
        self.assertIsNone(list_tag_for("Normal"))


if __name__ == "__main__":
    unittest.main()
