# Docx → XML

Wandelt Word-Dokumente (`.docx`) in medienneutrales XML um.

## Voraussetzungen

Wähle **einen** der beiden Wege. Microsoft Word und eine Internetverbindung (nach der Einrichtung) sind nicht nötig. Es werden (vorerst) nur `.docx`-Dateien unterstützt (kein `.doc`, kein `.odt`).

**Weg A: Docker** (sonst nichts erforderlich)

- Docker Desktop (unter Windows mit WSL2 oder Hyper-V)
- Einmalig im Projektordner: `docker build -t docx2xml .`

**Weg B: Lokal**

- Windows mit PowerShell
- Python 3.10 oder neuer
- Einmalig im Projektordner:

```powershell
python -m venv .venv
.venv\Scripts\python.exe -m pip install -r requirements.txt
New-Item -ItemType Directory -Force input, output
```

## DANACH: So erstellst du die XML-Dateien

1. Lege die Word-Dateien in den Ordner **`input`** (im Projektordner). Schließe sie vorher in Word, sonst entstehen Temp-Dateien (`~$*.docx`), die die Umwandlung stören.
2. Öffne PowerShell im Projektordner.
3. Führe den Befehl aus. Er wandelt **alle** `.docx`-Dateien aus `input` um.

**Weg A: Docker**

Alle Dateien im Ordner `input` konvertieren:
```powershell
Get-ChildItem input\*.docx | ForEach-Object { docker run --rm -v "${PWD}/input:/in" -v "${PWD}/output:/out" docx2xml "/in/$($_.Name)" "/out/$($_.BaseName).xml" schema.xsd }
```

Oder eine einzelne Datei:

```powershell
docker run --rm -v "${PWD}/input:/in" -v "${PWD}/output:/out" docx2xml /in/test.docx /out/test.xml schema.xsd
```

**Weg B: Lokal**

Alle Dateien (lokal):

```powershell
Get-ChildItem input\*.docx | ForEach-Object { .venv\Scripts\python.exe app.py $_.FullName "output\$($_.BaseName).xml" schema.xsd }
```



Einzelne Datei (lokal):

```powershell
.venv\Scripts\python.exe app.py input\test.docx output\test.xml schema.xsd
```
 

4. Die fertigen XML-Dateien liegen im Ordner **`output`** (`test.docx` → `test.xml`).

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
