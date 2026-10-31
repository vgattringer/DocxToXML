"""Prüfung des erzeugten XML gegen das Schema."""
from lxml import etree

from docx_reader import warnings


def validate(root: etree._Element, xsd_path: str) -> bool:
    """Prüft das Ergebnis gegen das Schema und meldet Fehler lesbar."""
    schema = etree.XMLSchema(etree.parse(xsd_path))
    if schema.validate(root):
        return True
    for err in schema.error_log:
        warnings.append(f"Schemafehler Zeile {err.line}: {err.message}")
    return False
