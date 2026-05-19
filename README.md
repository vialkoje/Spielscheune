# Spielscheune – Rechnungsgenerator

Automatische Erstellung und Versand von PDF-Rechnungen für die Kinderbetreuung „Sonnenstunden".

---

## Funktionsweise

1. Stamm- und Bewegungsdaten werden aus `spielscheune.xlsm` gelesen (Blätter *Kundenkartei* und *Monatsabrechnung*)
2. Für jedes Kind wird geprüft, ob die Abrechnungsfrequenz eine neue Rechnung erlaubt
3. Bereits abgerechnete Einträge werden übersprungen (Duplikatschutz)
4. `vorlage.docx` wird mit den Daten befüllt und per LibreOffice in ein PDF gewandelt
5. Das PDF wird per E-Mail an die hinterlegte Adresse verschickt
6. Rechnungsnummer und Abrechnungshistorie werden persistiert

---

## Voraussetzungen

- Python 3.10+
- [LibreOffice](https://www.libreoffice.org/download/libreoffice/) (für PDF-Erzeugung)
- Microsoft Excel (für `.xlsm` mit Makro-Button)
- Strato SMTP-Zugang für `sonnenstunden@vialkowitsch.de`

---

## Einrichtung

```bash
# 1. Virtuelle Umgebung anlegen und Abhängigkeiten installieren
python3 -m venv venv
source venv/bin/activate
pip install -r requirements.txt

# 2. Passwort konfigurieren
cp .env.example .env
# .env öffnen und SMTP_PASS eintragen

# 3. Excel-Datei erstellen (einmalig)
python CreateExcel.py
```

### Makro-Button einrichten (einmalig)

1. `spielscheune.xlsx` in Excel öffnen
2. **Datei → Speichern unter → Excel-Arbeitsmappe mit Makros (.xlsm)**  
   Dateiname: `spielscheune.xlsm`
3. **Extras → Makros → Visual Basic Editor** (oder `Alt+F11`)
4. **Datei → Importieren** → `AbrechnungMakro.bas` auswählen
5. VBA-Editor schließen
6. **Rechtsklick auf den grünen Button** → *Makro zuweisen* → `StartAbrechnung` → OK

Nach dieser Einrichtung startet ein Klick auf den Button den kompletten Abrechnungslauf.

---

## Tägliche Nutzung

### Daten pflegen

Alle Stamm- und Bewegungsdaten werden direkt in `spielscheune.xlsm` gepflegt:

| Blatt | Inhalt | Ändert sich |
|---|---|---|
| **Kundenkartei** | Name, Adresse, E-Mail, Abrechnungsfrequenz | Selten (neue Kunden, Umzug) |
| **Monatsabrechnung** | Kind, Leistungszeitraum, Anzahl Essen, Preis | Jeden Monat neu |

Die Spalte **Name** in der Monatsabrechnung wird per XLOOKUP automatisch aus der Kundenkartei befüllt (hellblau hinterlegt, nicht manuell bearbeiten).

### Abrechnung starten

- **Per Button:** Grünen Button in der Monatsabrechnung klicken → Bestätigung → fertig
- **Per Terminal:** `bash start.sh`

### Ausgabe (Beispiel)

```
INFO: Spielscheune Rechnungsgenerator
INFO: Lade: spielscheune.xlsm
INFO: Datum: 19.05.2026  |  4 Datensatz/Datensätze
INFO:   ✓  Max Muster           RE-202605-0001  →  Rechnung_K123_RE-202605-0001.pdf
INFO:   ✉  Gesendet an eltern@example.de
INFO:   ⏭  Jonas               noch nicht fällig – nächste: Juli 2026
WARNING:   ⚠  Jürgen              bereits abgerechnet (RE-202604-0003) – übersprungen
INFO: Fertig: 1 erstellt, 1 bereits abgerechnet, 1 noch nicht fällig  →  rechnungen/
```

---

## Abrechnungsfrequenz

Die Frequenz pro Kunde steuert, wann die nächste Rechnung fällig ist:

| Wert | Bedeutung | Nächste Abrechnung nach |
|---|---|---|
| `monatlich` | Jeden Monat | 1 Monat |
| `vierteljährlich` | Einmal pro Quartal | 3 Monate |
| `halbjährlich` | Zweimal pro Jahr | 6 Monate |
| `jährlich` | Einmal pro Jahr | 12 Monate |

Einträge, die noch nicht fällig sind, werden mit `⏭` geloggt und **nicht** übersprungen – sie liegen einfach im nächsten Lauf automatisch bereit, sobald der Zeitraum erreicht ist.

---

## Rechnungsnummern

Format: `RE-YYYYMM-NNNN` (z. B. `RE-202605-0001`).  
Der Zähler wird in `rechnungsnummern.json` pro Monat gespeichert und läuft nicht zurück.

---

## Persistenz-Dateien

| Datei | Zweck |
|---|---|
| `bereits_abgerechnet.json` | Duplikatschutz: welche (Kunde, Kind, LZR)-Kombination wurde abgerechnet |
| `rechnungsnummern.json` | Rechnungsnummer-Zähler pro Monat |
| `abrechnungshistorie.csv` | Vollständiges Logbuch aller ausgestellten Rechnungen (append-only) |

---

## Vorlage

`vorlage.docx` enthält Platzhalter in geschweiften Klammern:

| Platzhalter | Wird ersetzt durch |
|---|---|
| `{Name}` | Name des Erziehungsberechtigten |
| `{Strasse}` | Straße und Hausnummer |
| `{PLZ}` | Postleitzahl |
| `{Ort}` | Wohnort |
| `{Kundennummer}` | Kundennummer |
| `{Abrechnungsfrequenz}` | Abrechnungsfrequenz |
| `{Kind_Vorname}` | Vorname des Kindes |
| `{LZR}` | Leistungszeitraum |
| `{Rechnungsnummer}` | Automatisch vergebene Rechnungsnummer |
| `{Rechnungsdatum}` | Datum der Rechnungserstellung |
| `{Anzahl_Essen}` | Anzahl der Mahlzeiten |
| `{Preis_pro_Essen}` | Einzelpreis in EUR |
| `{Gesamtpreis}` | Gesamtbetrag in EUR |

---

## Projektstruktur

```
Spielscheune/
├── CreateInvoice.py          # Hauptprogramm
├── CreateExcel.py            # Erstellt spielscheune.xlsx (einmalig)
├── AbrechnungMakro.bas       # VBA-Makro für den Excel-Button
├── spielscheune.xlsm         # Arbeitsmappe (Kundenkartei + Monatsabrechnung)
├── vorlage.docx              # Rechnungsvorlage
├── start.sh                  # Startskript (alternativ zum Button)
├── test_invoice.py           # Pytest-Testsuite (69 Tests)
├── requirements.txt          # Python-Abhängigkeiten
├── .env.example              # Vorlage für lokale Konfiguration
├── .env                      # Lokales Passwort (nicht im Repo)
├── .gitignore
├── bereits_abgerechnet.json  # Duplikatschutz (wird automatisch gepflegt)
├── rechnungsnummern.json     # Nummernzähler (wird automatisch gepflegt)
├── abrechnungshistorie.csv   # Logbuch aller Rechnungen (wird automatisch gepflegt)
└── rechnungen/               # Generierte PDFs (nicht im Repo)
    └── YYYY-MM/
        ├── Rechnung_K123_RE-202605-0001.docx
        └── Rechnung_K123_RE-202605-0001.pdf
```

---

## Tests ausführen

```bash
source venv/bin/activate
python -m pytest test_invoice.py -v
```

69 Tests decken ab: LZR-Parsing, Frequenzprüfung (alle 4 Stufen), Duplikatschutz, Rechnungsnummern, Abrechnungshistorie, Excel-Import und vollständige Integrationsläufe.

---

## Abhängigkeiten

| Paket | Zweck |
|---|---|
| `python-docx` | Word-Vorlage befüllen |
| `pandas` | Daten verarbeiten |
| `openpyxl` | Excel-Datei lesen |
