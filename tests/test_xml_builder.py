import unittest

from docx import Document
from docx.enum.text import WD_UNDERLINE
from lxml import etree

from xml_builder import add_text, append_table, fill_inline, format_tags


def xml_to_string(el: etree._Element) -> str:
    return etree.tostring(el, encoding="unicode")


class AddTextTest(unittest.TestCase):
    def test_appends_to_text_without_children(self) -> None:
        el = etree.Element("p")
        add_text(el, "a")
        add_text(el, "b")
        self.assertEqual(el.text, "ab")

    def test_appends_to_tail_of_last_child(self) -> None:
        el = etree.Element("p")
        etree.SubElement(el, "b").text = "x"
        add_text(el, "y")
        add_text(el, "z")
        self.assertEqual(xml_to_string(el), "<p><b>x</b>yz</p>")

    def test_keeps_text_before_first_child(self) -> None:
        el = etree.Element("p")
        add_text(el, "a")
        etree.SubElement(el, "b").text = "x"
        add_text(el, "y")
        self.assertEqual(xml_to_string(el), "<p>a<b>x</b>y</p>")


class FormatTagsTest(unittest.TestCase):
    def setUp(self) -> None:
        self.par = Document().add_paragraph()

    def test_no_formatting(self) -> None:
        self.assertEqual(format_tags(self.par.add_run("x")), [])

    def test_superscript(self) -> None:
        run = self.par.add_run("x")
        run.font.superscript = True
        self.assertEqual(format_tags(run), ["sup"])

    def test_all_formats_in_nesting_order(self) -> None:
        run = self.par.add_run("x")
        run.bold = True
        run.italic = True
        run.underline = True
        run.font.strike = True
        run.font.superscript = True  # sup und sub schließen sich in python-docx gegenseitig aus
        self.assertEqual(format_tags(run), ["b", "i", "u", "s", "sup"])

    def test_explicit_false_is_inactive(self) -> None:
        run = self.par.add_run("x")
        run.bold = False
        self.assertEqual(format_tags(run), [])

    def test_each_single_format(self) -> None:
        setters = {
            "b": lambda r: setattr(r, "bold", True),
            "i": lambda r: setattr(r, "italic", True),
            "u": lambda r: setattr(r, "underline", True),
            "s": lambda r: setattr(r.font, "strike", True),
            "sup": lambda r: setattr(r.font, "superscript", True),
            "sub": lambda r: setattr(r.font, "subscript", True),
        }
        for tag, apply in setters.items():
            with self.subTest(tag):
                run = self.par.add_run("x")
                apply(run)
                self.assertEqual(format_tags(run), [tag])

    def test_underline_style_counts_as_underline(self) -> None:
        run = self.par.add_run("x")
        run.underline = WD_UNDERLINE.DOUBLE
        self.assertEqual(format_tags(run), ["u"])

    def test_underline_none_is_inactive(self) -> None:
        run = self.par.add_run("x")
        run.underline = WD_UNDERLINE.NONE
        self.assertEqual(format_tags(run), [])

    def test_subscript_replaces_superscript(self) -> None:
        run = self.par.add_run("x")
        run.font.superscript = True
        run.font.subscript = True
        self.assertEqual(format_tags(run), ["sub"])


class FillInlineTest(unittest.TestCase):
    def setUp(self) -> None:
        self.par = Document().add_paragraph()

    def test_bold_italic_normal(self) -> None:
        self.par.add_run("a ")
        self.par.add_run("b").bold = True
        self.par.add_run(" c ")
        self.par.add_run("d").italic = True
        el = etree.Element("p")
        fill_inline(el, self.par)
        self.assertEqual(xml_to_string(el), "<p>a <b>b</b> c <i>d</i></p>")

    def test_bold_and_italic_nested(self) -> None:
        run = self.par.add_run("x")
        run.bold = True
        run.italic = True
        el = etree.Element("p")
        fill_inline(el, self.par)
        self.assertEqual(xml_to_string(el), "<p><b><i>x</i></b></p>")

    def test_new_formats(self) -> None:
        run = self.par.add_run("x")
        run.underline = True
        run.font.strike = True
        run.font.subscript = True
        el = etree.Element("p")
        fill_inline(el, self.par)
        self.assertEqual(xml_to_string(el), "<p><u><s><sub>x</sub></s></u></p>")

    def test_text_around_formatted_runs(self) -> None:
        self.par.add_run("a")
        self.par.add_run("b").bold = True
        self.par.add_run("c")
        self.par.add_run("d")
        el = etree.Element("p")
        fill_inline(el, self.par)
        self.assertEqual(xml_to_string(el), "<p>a<b>b</b>cd</p>")

    def test_superscript(self) -> None:
        self.par.add_run("x").font.superscript = True
        el = etree.Element("p")
        fill_inline(el, self.par)
        self.assertEqual(xml_to_string(el), "<p><sup>x</sup></p>")

    def test_plain_ignores_formatting(self) -> None:
        self.par.add_run("a").bold = True
        self.par.add_run("b").italic = True
        el = etree.Element("title")
        fill_inline(el, self.par, plain=True)
        self.assertEqual(xml_to_string(el), "<title>ab</title>")

    def test_empty_runs_are_skipped(self) -> None:
        self.par.add_run("").bold = True
        el = etree.Element("p")
        fill_inline(el, self.par)
        self.assertEqual(xml_to_string(el), "<p/>")

    def test_paragraph_without_runs(self) -> None:
        el = etree.Element("p")
        fill_inline(el, self.par)
        self.assertEqual(xml_to_string(el), "<p/>")

    def test_adjacent_formatted_runs_are_not_merged(self) -> None:
        self.par.add_run("a").bold = True
        self.par.add_run("b").bold = True
        el = etree.Element("p")
        fill_inline(el, self.par)
        self.assertEqual(xml_to_string(el), "<p><b>a</b><b>b</b></p>")

    def test_special_characters_are_escaped(self) -> None:
        self.par.add_run("a < b & c > d")
        el = etree.Element("p")
        fill_inline(el, self.par)
        self.assertEqual(el.text, "a < b & c > d")
        self.assertEqual(xml_to_string(el), "<p>a &lt; b &amp; c &gt; d</p>")

    def test_line_break_and_tab_are_kept(self) -> None:
        self.par.add_run("a\nb\tc")
        el = etree.Element("p")
        fill_inline(el, self.par)
        self.assertEqual(el.text, "a\nb\tc")

    def test_plain_skips_empty_runs(self) -> None:
        self.par.add_run("")
        self.par.add_run("x")
        el = etree.Element("title")
        fill_inline(el, self.par, plain=True)
        self.assertEqual(xml_to_string(el), "<title>x</title>")

    def test_appends_to_existing_content(self) -> None:
        self.par.add_run("b").bold = True
        el = etree.Element("p")
        el.text = "a"
        fill_inline(el, self.par)
        self.assertEqual(xml_to_string(el), "<p>a<b>b</b></p>")


class AppendTableTest(unittest.TestCase):
    def test_structure_and_header(self) -> None:
        table = Document().add_table(rows=2, cols=2)
        for i, row in enumerate(table.rows):
            for j, cell in enumerate(row.cells):
                cell.text = f"{i}{j}"
        container = etree.Element("section")
        append_table(container, table)
        rows = container.findall("table/row")
        self.assertEqual(len(rows), 2)
        self.assertEqual(rows[0].get("header"), "true")
        self.assertIsNone(rows[1].get("header"))
        self.assertEqual([e.text for e in rows[1]], ["10", "11"])

    def test_formatting_in_cell(self) -> None:
        table = Document().add_table(rows=1, cols=1)
        table.cell(0, 0).paragraphs[0].add_run("x").bold = True
        container = etree.Element("section")
        append_table(container, table)
        self.assertEqual(xml_to_string(container.find("table/row/entry")), "<entry><b>x</b></entry>")

    def test_empty_cell(self) -> None:
        table = Document().add_table(rows=1, cols=1)
        container = etree.Element("section")
        append_table(container, table)
        entry = container.find("table/row/entry")
        self.assertIsNone(entry.text)
        self.assertEqual(len(entry), 0)

    def test_single_row_is_header(self) -> None:
        table = Document().add_table(rows=1, cols=2)
        container = etree.Element("section")
        append_table(container, table)
        rows = container.findall("table/row")
        self.assertEqual(len(rows), 1)
        self.assertEqual(rows[0].get("header"), "true")

    def test_only_first_paragraph_of_cell(self) -> None:
        table = Document().add_table(rows=1, cols=1)
        cell = table.cell(0, 0)
        cell.text = "first"
        cell.add_paragraph("second")
        container = etree.Element("section")
        append_table(container, table)
        self.assertEqual(container.find("table/row/entry").text, "first")

    def test_every_cell_becomes_an_entry(self) -> None:
        table = Document().add_table(rows=3, cols=3)
        container = etree.Element("section")
        append_table(container, table)
        rows = container.findall("table/row")
        self.assertEqual([len(r) for r in rows], [3, 3, 3])

    def test_only_first_row_is_header(self) -> None:
        table = Document().add_table(rows=3, cols=1)
        container = etree.Element("section")
        append_table(container, table)
        headers = [r.get("header") for r in container.findall("table/row")]
        self.assertEqual(headers, ["true", None, None])

    def test_table_is_appended_after_existing_children(self) -> None:
        container = etree.Element("section")
        etree.SubElement(container, "title").text = "T"
        append_table(container, Document().add_table(rows=1, cols=1))
        self.assertEqual([c.tag for c in container], ["title", "table"])

    def test_two_tables_in_same_container(self) -> None:
        doc = Document()
        container = etree.Element("section")
        append_table(container, doc.add_table(rows=1, cols=1))
        append_table(container, doc.add_table(rows=1, cols=1))
        self.assertEqual([c.tag for c in container], ["table", "table"])

    def test_cell_keeps_mixed_formatting(self) -> None:
        table = Document().add_table(rows=1, cols=1)
        par = table.cell(0, 0).paragraphs[0]
        par.add_run("a")
        par.add_run("b").italic = True
        container = etree.Element("section")
        append_table(container, table)
        self.assertEqual(
            xml_to_string(container.find("table/row/entry")), "<entry>a<i>b</i></entry>")


if __name__ == "__main__":
    unittest.main()
