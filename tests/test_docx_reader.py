import unittest

from docx import Document
from docx.table import Table
from docx.text.paragraph import Paragraph

import docx_reader
from docx_reader import heading_level, iter_blocks, list_tag_for


class IterBlocksTest(unittest.TestCase):
    def test_order_and_types(self):
        doc = Document()
        doc.add_paragraph("Eins")
        doc.add_table(rows=1, cols=1)
        doc.add_paragraph("Zwei")
        blocks = list(iter_blocks(doc))
        self.assertEqual([type(b) for b in blocks], [Paragraph, Table, Paragraph])
        self.assertEqual(blocks[2].text, "Zwei")

    def test_empty_document(self):
        self.assertEqual(list(iter_blocks(Document())), [])


class HeadingLevelTest(unittest.TestCase):
    def setUp(self):
        docx_reader.warnings.clear()
        self.doc = Document()

    def test_heading_style(self):
        self.assertEqual(heading_level(self.doc.add_heading("A", 3)), 3)

    def test_title_is_level_1(self):
        self.assertEqual(heading_level(self.doc.add_heading("A", 0)), 1)

    def test_heuristic_short_bold_paragraph(self):
        p = self.doc.add_paragraph()
        p.add_run("Kurz und fett").bold = True
        self.assertEqual(heading_level(p), 2)
        self.assertEqual(len(docx_reader.warnings), 1)

    def test_heuristic_not_applied_with_sentence_end(self):
        p = self.doc.add_paragraph()
        p.add_run("Ein Satz.").bold = True
        self.assertIsNone(heading_level(p))

    def test_heuristic_not_applied_to_long_text(self):
        p = self.doc.add_paragraph()
        p.add_run("x" * 80).bold = True
        self.assertIsNone(heading_level(p))

    def test_heuristic_not_applied_when_partly_bold(self):
        p = self.doc.add_paragraph()
        p.add_run("fett").bold = True
        p.add_run(" normal")
        self.assertIsNone(heading_level(p))

    def test_normal_paragraph(self):
        self.assertIsNone(heading_level(self.doc.add_paragraph("Text")))
        self.assertEqual(docx_reader.warnings, [])


class ListTagForTest(unittest.TestCase):
    def test_mapping(self):
        self.assertEqual(list_tag_for("List Bullet"), "ul")
        self.assertEqual(list_tag_for("List Bullet 2"), "ul")
        self.assertEqual(list_tag_for("List Number"), "ol")
        self.assertIsNone(list_tag_for("Normal"))


if __name__ == "__main__":
    unittest.main()
