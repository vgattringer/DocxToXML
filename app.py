"""Einstiegspunkt: Word (.docx) -> medienneutrales XML.

Aufruf: python app.py input.docx output.xml [schema.xsd]
"""
import sys

from docx import Document
from lxml import etree

from docx_converter import DocxConverter
from docx_reader import warnings
from xml_validator import validate


def main(args: list[str]) -> None:
    if len(args) < 3:
        sys.exit("Aufruf: python app.py input.docx output.xml [schema.xsd]")
    tree = DocxConverter().convert(Document(args[1]))
    valid = validate(tree, args[3]) if len(args) > 3 else None
    etree.ElementTree(tree).write(
        args[2], pretty_print=True, encoding="utf-8", xml_declaration=True)
    print(f"XML geschrieben: {args[2]} | schema-valide: {valid}")
    for w in warnings:
        print("WARNUNG:", w)


if __name__ == "__main__":
    main(sys.argv)
