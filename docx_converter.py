"""Die Klasse DocxConverter wandelt ein Word-Dokument in einen XML-Baum um."""
from docx.document import Document as DocxDocument
from docx.table import Table
from docx.text.paragraph import Paragraph
from lxml import etree

from docx_reader import (heuristic_heading_level, iter_blocks, list_tag_for,
                         style_heading_level, warnings)
from xml_builder import append_table, fill_inline


class DocxConverter:
    """Hält den Zustand der Umwandlung (offene Abschnitte, offene Liste)."""

    def __init__(self) -> None:
        self.root: etree._Element = etree.Element("document")
        # (Ebene, Element) der offenen Abschnitte
        self.stack: list[tuple[int, etree._Element]] = [(0, self.root)]
        # offene Liste, damit <li> zusammengefasst werden
        self.current_list: etree._Element | None = None

    @property
    def container(self) -> etree._Element:
        """Das Element, in das neue Inhalte eingefügt werden (innerster Abschnitt)."""
        return self.stack[-1][1]

    def convert(self, doc: DocxDocument) -> etree._Element:
        for block in iter_blocks(doc):
            self.handle_block(block)
        return self.root

    def handle_block(self, block: Paragraph | Table) -> None:
        """Verteilt einen Baustein auf die passende Behandlung."""
        if isinstance(block, Table):
            self.current_list = None
            append_table(self.container, block)
            return
        # TODO Hyperlinks: block.text enthält den Linktext, fill_inline (xml_builder) bisher nicht.
        #  Ein Absatz nur aus einem Link ergibt deshalb ein leeres <p/>. Löst sich, sobald
        #  fill_inline Hyperlinks übernimmt.
        if not block.text.strip():
            return  # leere Absätze ignorieren
        # Priorität: ausdrückliche Vorlage vor Heuristik
        style = block.style.name or ""
        if level := style_heading_level(style):
            self.handle_heading(block, level)
        elif list_tag := list_tag_for(style):
            self.handle_list_item(block, list_tag)
        elif level := heuristic_heading_level(block):
            self.handle_heading(block, level)
        else:
            self.handle_paragraph(block, style)

    def handle_heading(self, block: Paragraph, level: int) -> None:
        self.current_list = None
        while self.stack[-1][0] >= level:
            self.stack.pop()
        section = etree.SubElement(self.container, "section")
        title = etree.SubElement(section, "title")
        fill_inline(title, block, plain=True)
        self.stack.append((level, section))

    def handle_list_item(self, block: Paragraph, list_tag: str) -> None:
        if self.current_list is None or self.current_list.tag != list_tag:
            self.current_list = etree.SubElement(self.container, list_tag)
        fill_inline(etree.SubElement(self.current_list, "li"), block)

    def handle_paragraph(self, block: Paragraph, style: str) -> None:
        self.current_list = None
        if style not in ("Normal", "Body Text"):
            warnings.append(f"Unbekannte Formatvorlage '{style}' -> als <p> übernommen")
        fill_inline(etree.SubElement(self.container, "p"), block)
