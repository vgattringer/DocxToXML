import unittest

from docx import Document
from lxml import etree

from xml_builder import add_text, append_table, fill_inline


def xml(el: etree._Element) -> str:
    return etree.tostring(el, encoding="unicode")


class AddTextTest(unittest.TestCase):
    def test_an_text_ohne_kinder(self) -> None:
        el = etree.Element("p")
        add_text(el, "a")
        add_text(el, "b")
        self.assertEqual(el.text, "ab")

    def test_an_tail_des_letzten_kindes(self) -> None:
        el = etree.Element("p")
        etree.SubElement(el, "b").text = "x"
        add_text(el, "y")
        add_text(el, "z")
        self.assertEqual(xml(el), "<p><b>x</b>yz</p>")


class FillInlineTest(unittest.TestCase):
    def setUp(self) -> None:
        self.par = Document().add_paragraph()

    def test_fett_kursiv_normal(self) -> None:
        self.par.add_run("a ")
        self.par.add_run("b").bold = True
        self.par.add_run(" c ")
        self.par.add_run("d").italic = True
        el = etree.Element("p")
        fill_inline(el, self.par)
        self.assertEqual(xml(el), "<p>a <b>b</b> c <i>d</i></p>")

    def test_fett_und_kursiv_verschachtelt(self) -> None:
        run = self.par.add_run("x")
        run.bold = True
        run.italic = True
        el = etree.Element("p")
        fill_inline(el, self.par)
        self.assertEqual(xml(el), "<p><b><i>x</i></b></p>")

    def test_neue_formate(self) -> None:
        run = self.par.add_run("x")
        run.underline = True
        run.font.strike = True
        run.font.subscript = True
        el = etree.Element("p")
        fill_inline(el, self.par)
        self.assertEqual(xml(el), "<p><u><s><sub>x</sub></s></u></p>")

    def test_plain_ignoriert_formatierung(self) -> None:
        self.par.add_run("a").bold = True
        self.par.add_run("b").italic = True
        el = etree.Element("title")
        fill_inline(el, self.par, plain=True)
        self.assertEqual(xml(el), "<title>ab</title>")

    def test_leere_runs_werden_uebersprungen(self) -> None:
        self.par.add_run("").bold = True
        el = etree.Element("p")
        fill_inline(el, self.par)
        self.assertEqual(xml(el), "<p/>")


class AppendTableTest(unittest.TestCase):
    def test_struktur_und_header(self) -> None:
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

    def test_nur_erster_absatz_der_zelle(self) -> None:
        table = Document().add_table(rows=1, cols=1)
        cell = table.cell(0, 0)
        cell.text = "erster"
        cell.add_paragraph("zweiter")
        container = etree.Element("section")
        append_table(container, table)
        self.assertEqual(container.find("table/row/entry").text, "erster")


if __name__ == "__main__":
    unittest.main()
