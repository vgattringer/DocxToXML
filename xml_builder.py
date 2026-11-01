"""Aufbau des XML-Baums aus Absätzen und Tabellen."""
from docx.table import Table
from docx.text.run import Run
from docx.text.paragraph import Paragraph
from lxml import etree


def add_text(parent: etree._Element, text: str) -> None:
    """Hängt Text korrekt an (an .text oder an .tail des letzten Kindes)."""
    if len(parent):
        parent[-1].tail = (parent[-1].tail or "") + text
    else:
        parent.text = (parent.text or "") + text


def format_tags(run: Run) -> list[str]:
    """Liefert die aktiven Formatierungen eines Runs als Tag-Namen (Reihenfolge = Verschachtelung)."""
    paare: list[tuple[str, bool | None]] = [
        ("b", run.bold),
        ("i", run.italic),
        ("u", run.underline),
        ("s", run.font.strike),
        ("sup", run.font.superscript),
        ("sub", run.font.subscript),
    ]
    return [t for t, aktiv in paare if aktiv]


def fill_inline(parent: etree._Element, paragraph: Paragraph, plain: bool = False) -> None:
    """Überträgt Runs als Text bzw. verschachtelte Formatierungs-Tags. plain=True: nur Text (z. B. Titel)."""
    for run in paragraph.runs:
        if not run.text:
            continue
        tags = [] if plain else format_tags(run)
        if not tags:
            add_text(parent, run.text)
            continue
        element = parent
        for tag in tags:
            element = etree.SubElement(element, tag)
        element.text = run.text

def append_table(container: etree._Element, table: Table) -> None:
    """Überträgt eine Word-Tabelle als <table> in den Container."""
    table_el = etree.SubElement(container, "table")
    for i, row in enumerate(table.rows):
        row_el = etree.SubElement(table_el, "row")
        if i == 0:
            row_el.set("header", "true")
        for cell in row.cells:
            entry = etree.SubElement(row_el, "entry")
            # Pro Zelle nur der Text des ersten Absatzes (bewusste Vereinfachung)
            fill_inline(entry, cell.paragraphs[0])
