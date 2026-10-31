# Docx → XML

Wandelt Word-Dokumente (`.docx`) in medienneutrales XML um.

## So erstellst du die XML-Dateien

1. Lege die Word-Dateien in den Ordner **`eingaben`** (im Projektordner).
2. Öffne PowerShell im Projektordner.
3. Führe diesen Befehl aus. Er wandelt **alle** `.docx`-Dateien aus `eingaben` um:

```powershell
Get-ChildItem eingaben\*.docx | ForEach-Object { .venv\Scripts\python.exe app.py $_.FullName "ausgaben\$($_.BaseName).xml" schema.xsd }
```

4. Die fertigen XML-Dateien liegen im Ordner **`ausgaben`** (`test.docx` → `test.xml`).

**Möchtest du nur eine einzige Datei umwandeln?** → Für eine einzelne Datei:

```powershell
.venv\Scripts\python.exe app.py eingaben\test.docx ausgaben\test.xml schema.xsd
```

Beim ersten Mal in einer neuen Umgebung: `.venv\Scripts\python.exe -m pip install python-docx lxml`

Im Terminal erscheint am Ende `schema-valide: True/False` sowie **WARNUNG**-Zeilen für Stellen, die die Redaktion prüfen sollte.

## Wie die Konvertierung funktioniert

| Datei | Aufgabe |
|---|---|
| `app.py` | Einstiegspunkt: liest die Datei, startet die Umwandlung, prüft, schreibt das XML |
| `docx_reader.py` | Liest Absätze und Tabellen in Dokumentreihenfolge; erkennt Überschriften und Listen |
| `docx_converter.py` | Baut aus den Bausteinen den XML-Baum |
| `xml_builder.py` | Überträgt Text, Formatierung und Tabellen |
| `xml_validator.py` | Prüft das Ergebnis gegen `schema.xsd` |

Regeln:

- **Überschriften** (`Heading 1–n`, `Title`) werden zu verschachtelten `<section>` mit `<title>`. Die Ebene bestimmt die Verschachtelung.
- **Heuristik:** Ein kurzer (< 80 Zeichen), komplett fetter Absatz ohne Punkt am Ende gilt als Überschrift Ebene 2 (mit Warnung).
- **Listen** (`List Bullet…` / `List Number…`) werden zu `<ul>` / `<ol>` mit `<li>`; aufeinanderfolgende Einträge landen in derselben Liste.
- **Alles andere** wird `<p>`. Unbekannte Formatvorlagen (z. B. `Quote`) werden als `<p>` übernommen, mit Warnung.
- **Formatierung** (fett, kursiv, unterstrichen, durchgestrichen, hoch-/tiefgestellt) wird zu `<b>`, `<i>`, `<u>`, `<s>`, `<sup>`, `<sub>`.
- **Tabellen** werden zu `<table>/<row>/<entry>`; die erste Zeile ist `header="true"`. Pro Zelle wird nur der erste Absatz übernommen.
- **Leere Absätze** werden ignoriert.
