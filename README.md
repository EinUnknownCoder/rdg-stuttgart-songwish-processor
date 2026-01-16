# RDG Stuttgart Song Wish Processor

Ein Python-Tool zur Verarbeitung und Validierung von Songwünschen für Random Dance Games (RDG) in Stuttgart.

## Features

- **URL-Validierung**: Prüft YouTube-URLs auf Korrektheit
- **Artist/Title-Matching**: Vergleicht Künstler und Songtitel mit dem YouTube-Videotitel (ignoriert Groß-/Kleinschreibung und Sonderzeichen)
- **Lyric-Video-Erkennung**: Erkennt, ob es sich um ein Lyric Video handelt (warnt bei offiziellen MVs und Dance Practice Videos)
- **Dauer-Validierung**: Prüft, ob der gewünschte Songabschnitt max. 90 Sekunden lang ist
- **18+-Prüfung**: Erkennt altersbeschränkte Videos
- **Blockliste**: Unterstützung für manuell gepflegte Song-Blockliste
- **Automatische Nachrichtenerstellung**: Generiert zweisprachige Nachrichten (DE/EN) für die Kontaktaufnahme

## Installation

```bash
# Virtual Environment erstellen (Python 3.10+)
python3.11 -m venv .venv

# Aktivieren
source .venv/bin/activate  # macOS/Linux
# oder
.venv\Scripts\activate     # Windows

# Dependencies installieren
pip install -r requirements.txt
```

## Verwendung

1. **Songwish-Datei vorbereiten**: Stelle sicher, dass `songwish.xlsx` im Projektverzeichnis liegt

2. **Formular-URL anpassen** (optional): In `songwish_processor.py` die Variable `FORM_URL` anpassen:
   ```python
   FORM_URL = "https://forms.gle/YOUR_FORM_URL_HERE"
   ```

3. **Script ausführen**:
   ```bash
   source .venv/bin/activate
   python songwish_processor.py
   ```

4. **Output**: Es werden folgende Dateien erstellt:
   - `output.xlsx` - Hauptausgabe mit zwei Worksheets:
     - **Messages**: Vorgefertigte Nachrichten für die ersten 50 Anfragen
     - **Songlist**: Alle Songs mit Validierungsergebnissen
   - `blocked_songs.xlsx` - Template für die Song-Blockliste

## Eingabedatei-Format (songwish.xlsx)

Die Eingabedatei muss folgende Spalten enthalten:

| Spalte | Beschreibung |
|--------|--------------|
| Timestamp | Zeitstempel der Anfrage |
| Email Address | E-Mail-Adresse des Anfragenden |
| Sprache der Regeln | Bevorzugte Sprache (🇩🇪 Deutsch / 🇬🇧 English) |
| Bevorzugte Kommunikation | Kontaktmethode (Instagram/WhatsApp) |
| Instagram @Name | Instagram-Benutzername |
| WhatsApp Number | Telefonnummer für WhatsApp |
| YT URL | YouTube-URL des ersten Songs |
| Künstler | Künstlername des ersten Songs |
| Songname | Titel des ersten Songs |
| Start Timestamp | Startzeit (MM:SS:00) |
| End Timestamp | Endzeit (MM:SS:00) |
| Anmerkung | Zusätzliche Anmerkungen |
| YT URL.1, Künstler.1, ... | Daten für den zweiten Song |

## Ausgabedatei-Format (output.xlsx)

### Worksheet: Messages

| Spalte | Beschreibung |
|--------|--------------|
| # | Laufende Nummer |
| Contact URL | Direkter Link zur Kontaktaufnahme (Instagram/WhatsApp) |
| Message | Vorgefertigte Nachricht in der gewählten Sprache |
| Status | OK oder Fehler |
| Artist | Künstlername |
| Title | Songtitel |
| Errors | Fehlerbeschreibung (falls vorhanden) |

### Worksheet: Songlist

Gleiche Spalten wie `request.xlsx` (Songlist) plus:
- **#**: Laufende Nummerierung
- **Anmerkung**: Anmerkung aus dem Songwunsch
- **Errors**: Validierungsfehler

Zeilen mit Fehlern sind rot markiert.

## Blockliste (blocked_songs.xlsx)

Eine Excel-Datei mit zwei Spalten zur manuellen Pflege:

| Artist | Title |
|--------|-------|
| ... | ... |

Songs in dieser Liste werden automatisch abgelehnt.

## Validierungsregeln

1. **Artist/Title-Match**: Künstler und Titel müssen im YouTube-Videotitel vorkommen
   - Groß-/Kleinschreibung wird ignoriert
   - Sonderzeichen und Leerzeichen werden ignoriert (z.B. "Stray Kids" = "straykids")

2. **Lyric Video**: Das Video sollte ein Lyric Video sein
   - Warnung bei: "official mv", "official music video", "dance practice", "m/v"
   - Akzeptiert: "lyric", "lyrics", "lyric video"

3. **Dauer**: Der gewählte Abschnitt darf max. 90 Sekunden lang sein

4. **Altersbeschränkung**: 18+ Videos sind nicht erlaubt

5. **URL-Bereinigung**: Playlist-Parameter (`&list=...`) werden automatisch entfernt

## Lizenz

MIT License
