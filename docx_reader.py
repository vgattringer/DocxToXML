"""Lesen des Word-Dokuments: Bausteine, Überschriften und Listen erkennen."""
from collections.abc import Iterator

from docx.document import Document as DocxDocument
from docx.table import Table
from docx.text.paragraph import Paragraph

# Warnungen sammeln, damit die Redaktion unsichere Stellen prüfen kann
warnings: list[str] = []


def iter_blocks(doc: DocxDocument) -> Iterator[Paragraph | Table]:
    """Liefert Absätze und Tabellen in der Reihenfolge des Dokuments."""
    for child in doc.element.body.iterchildren():
        if child.tag.endswith("}p"):
            yield Paragraph(child, doc)
        elif child.tag.endswith("}tbl"):
            yield Table(child, doc)


def style_heading_level(style: str) -> int | None:
    """Ermittelt die Überschriftenebene allein aus der Formatvorlage (None = keine)."""
    if style.startswith("Heading "):
        return int(style.split()[-1])
    if style == "Title":
        return 1
    return None


def heuristic_heading_level(paragraph: Paragraph) -> int | None:
    """Letzter Ausweg: erkennt Pseudo-Überschriften an Fettdruck, Länge und Satzende."""
    text = paragraph.text.strip()
    # TODO Hyperlinks: paragraph.runs übergeht Runs in Hyperlinks, ein fetter Absatz mit Link
    #  wird dadurch evtl. falsch (oder gar nicht) als Überschrift erkannt.
    runs = [r for r in paragraph.runs if r.text.strip()]
    # Heuristik: kurzer, komplett fetter Absatz ohne Satzende = Überschrift
    if runs and all(r.bold for r in runs) and len(text) < 80 and not text.endswith("."):
        warnings.append(f"Heuristik: '{text}' als Überschrift (Ebene 2) erkannt")
        return 2
    return None


def list_tag_for(style: str) -> str | None:
    """Ordnet eine Word-Formatvorlage einem Listen-Tag zu (None = keine Liste)."""
    if style.startswith("List Bullet"):
        return "ul"
    if style.startswith("List Number"):
        return "ol"
    return None
