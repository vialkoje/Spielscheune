# Spielscheune – Rechnungsgenerator

Automatische Erstellung und Versand von PDF-Rechnungen für die Kinderbetreuung „Sonnenstunden".

---

## Funktionsweise

1. Einträge aus `abrechnungsdaten.csv` werden gelesen
2. Für jeden Eintrag wird `vorlage.docx` mit den Daten befüllt
3. Das befüllte Dokument wird per LibreOffice in ein PDF umgewandelt
4. Das PDF wird per E-Mail an die hinterlegte Adresse verschickt
5. Die Rechnungsnummer wird automatisch vergeben und gespeichert

---

## Voraussetzungen

- Python 3.10+
- [LibreOffice](https://www.libreoffice.org/download/libreoffice/) (für PDF-Erzeugung)
- Strato SMTP-Zugang für `sonnenstunden@vialkowitsch.de`

---

## Einrichtung

```bash
# 1. Repository klonen
git clone https://github.com/vialkoje/Spielscheune.git
cd Spielscheune

# 2. Virtuelle Umgebung anlegen und Abhängigkeiten installieren
python3 -m venv venv
source venv/bin/activate
pip install -r requirements.txt

# 3. Passwort konfigurieren
cp .env.example .env
# .env öffnen und SMTP_PASS eintragen
```

---

## Konfiguration

### `.env` (wird nicht ins Repo eingecheckt)

```
SMTP_PASS=dein-strato-passwort
```

### `abrechnungsdaten.csv`

Semikolon-getrennte CSV-Datei mit folgenden Spalten:

| Spalte | Beschreibung | Beispiel |
|---|---|---|
| `Name` | Name des Erziehungsberechtigten | `Max Muster` |
| `Strasse` | Straße und Hausnummer | `Musterweg 1` |
| `PLZ` | Postleitzahl | `12345` |
| `Ort` | Wohnort | `Musterstadt` |
| `Kundennummer` | Eindeutige Kundennummer | `K123` |
| `Abrechnungsfrequenz` | z. B. `monatlich`, `vierteljährlich` | `monatlich` |
| `Anzahl_Essen` | Anzahl der abzurechnenden Mahlzeiten | `20` |
| `Preis_pro_Essen` | Preis je Mahlzeit (Komma als Dezimaltrennzeichen) | `3,50` |
| `Kind_Vorname` | Vorname des Kindes | `Elias` |
| `Rechnungsemail` | E-Mail-Empfänger für die Rechnung | `eltern@example.de` |
| `LZR` | Leistungszeitraum (Monat + Jahr) | `April 2025` |

### `vorlage.docx`

Word-Dokument mit Platzhaltern in geschweiften Klammern:

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

## Starten

```bash
bash start.sh
```

Die generierten Rechnungen werden unter `rechnungen/YYYY-MM/` abgelegt:

```
rechnungen/
└── 2025-04/
    ├── Rechnung_K123_RE-202504-0001.docx
    ├── Rechnung_K123_RE-202504-0001.pdf
    ├── Rechnung_K124_RE-202504-0002.docx
    └── Rechnung_K124_RE-202504-0002.pdf
```

---

## Rechnungsnummern

Das Format ist `RE-YYYYMM-NNNN` (z. B. `RE-202504-0001`).  
Der Zähler wird in `rechnungsnummern.json` pro Monat gespeichert und läuft nicht zurück.

---

## Projektstruktur

```
Spielscheune/
├── CreateInvoice.py        # Hauptprogramm
├── vorlage.docx            # Rechnungsvorlage
├── abrechnungsdaten.csv    # Abrechnungsdaten
├── start.sh                # Startskript
├── requirements.txt        # Python-Abhängigkeiten
├── .env.example            # Vorlage für lokale Konfiguration
├── .env                    # Lokales Passwort (nicht im Repo)
├── .gitignore
└── rechnungen/             # Generierte Rechnungen (nicht im Repo)
```

---

## Abhängigkeiten

| Paket | Zweck |
|---|---|
| `python-docx` | Word-Vorlage befüllen |
| `pandas` | CSV-Datei einlesen |
