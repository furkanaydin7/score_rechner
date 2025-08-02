# Standort- und Firmen-Score-Rechner

Ein integriertes Python-Tool zur automatischen Bewertung von Firmenstandorten basierend auf verschiedenen Standort- und Firmenparametern.

## 🚀 Features

- **Automatische Berechnung** aller relevanten Metriken
- **Integration** der bestehenden Python-Skripte für:
  - ÖV-Güteklasse (aus CSV-Datei)
  - Nächste ÖV-Haltestelle
  - Nächste Autobahnauffahrt (via Overpass API)
  - Nächster Parkplatz (via OpenStreetMap)
- **Excel-Export** mit formatierter Ausgabe
- **Batch-Verarbeitung** mehrerer Firmen
- **Konfigurierbar** über JSON-Datei

## 📋 Voraussetzungen

- Python 3.8 oder höher
- Internetverbindung (für OSM/Overpass-Abfragen)

## 🛠️ Installation

### 1. Repository klonen oder Dateien herunterladen

```bash
# Erstellen Sie einen neuen Ordner
mkdir standort-score-rechner
cd standort-score-rechner
```

### 2. Virtuelle Umgebung erstellen (empfohlen)

```bash
# Windows
python -m venv venv
venv\Scripts\activate

# Linux/Mac
python3 -m venv venv
source venv/bin/activate
```

### 3. Abhängigkeiten installieren

```bash
pip install pandas numpy geopandas shapely osmnx overpy haversine pyproj openpyxl
```

Alternativ mit requirements.txt:

```bash
# requirements.txt erstellen mit folgendem Inhalt:
pandas>=1.3.0
numpy>=1.21.0
geopandas>=0.10.0
shapely>=1.8.0
osmnx>=1.1.0
overpy>=0.6
haversine>=2.5.0
pyproj>=3.2.0
openpyxl>=3.0.0

# Dann installieren:
pip install -r requirements.txt
```

## 📁 Benötigte Dateien und Projektstruktur

```
standort-score-rechner/
├── .venv/                          # Virtuelle Umgebung (nach Installation)
├── data/                           # Datenverzeichnis
│   ├── oev_qualitaet_gemeinden.csv    # ÖV-Güteklassen der Gemeinden
│   └── Betriebspunkt.csv              # ÖV-Haltestellen-Daten
├── output/                         # Ausgabeverzeichnis (wird automatisch erstellt)
├── standort_score_rechner.py       # Hauptmodul mit allen Berechnungen
├── main_script.py                  # Ausführungsskript
├── config.json                     # Konfiguration mit Firmen und Standorten
├── requirements.txt                # Python-Abhängigkeiten
└── README.md                       # Diese Datei
```

## 🔧 Konfiguration

### config.json anpassen

Die `config.json` enthält zwei Hauptbereiche:

#### Standorte (nur Rohdaten erforderlich)
```json
"standorte": {
  "Zürich": {
    "anzahl_beschaeftigte": 536980,
    "anzahl_einwohner": 427721,
    "anzahl_einpendelnde": 338458.9,
    "motorisierungsgrad": 328,
    "modal_split_auto": 18.8
  }
}
```

**Automatisch berechnet/ermittelt werden:**
- `oev_gueteklasse` → aus der CSV-Datei
- `beschaeftigte_pro_1000` → (anzahl_beschaeftigte / anzahl_einwohner) × 1000
- `einpendler_prozent` → (anzahl_einpendelnde / anzahl_beschaeftigte) × 100

#### Firmen
```json
"firmen": [
  {
    "name": "UBS",
    "adresse": "Bahnhofstrasse 45, Zürich",
    "lat": 47.37206508238767,
    "lon": 8.538280693066223,
    "mitarbeiterzahl": 23400,
    "branche": "Finanzen, Versicherungen, Beratung",
    "standort": "Zürich"
  }
]
```

### Unterstützte Branchen

- IT & Software (5 Punkte)
- Finanzen, Versicherungen, Beratung (4 Punkte)
- Verwaltung, Bildung, Gesundheitswesen, Dienstleisungen (3 Punkte)
- Industrie, Produktion & Handel (2 Punkte)
- Logistik & Transport (1 Punkt)

## 🚀 Verwendung

### Einfache Ausführung

```bash
python main_script.py
```

Dies wird:
1. Alle Firmen aus der config.json verarbeiten
2. Alle Scores berechnen
3. Eine Excel-Datei mit Zeitstempel im `output/` Ordner erstellen
4. Eine Zusammenfassung anzeigen

### Programmatische Verwendung

```python
from pathlib import Path
from standort_score_rechner import StandortScoreRechner, Standort, Firma

# Pfade definieren
oev_path = Path("data/oev_qualitaet_gemeinden.csv")
betriebspunkt_path = Path("data/Betriebspunkt.csv")

# Rechner initialisieren
rechner = StandortScoreRechner(
    oev_qualitaet_path=oev_path,
    betriebspunkt_path=betriebspunkt_path
)

# Standort definieren (mit Rohdaten)
standort = Standort(
    name="Zürich",
    anzahl_beschaeftigte=536980,
    anzahl_einwohner=427721,
    anzahl_einpendelnde=338458.9,
    motorisierungsgrad=328,
    modal_split_auto=18.8
)

# Firma definieren
firma = Firma(
    name="Beispiel AG",
    adresse="Musterstrasse 1, Zürich",
    lat=47.3769,
    lon=8.5417,
    mitarbeiterzahl=100,
    branche="IT & Software"
)

# Score berechnen
ergebnis = rechner.berechne_scores(standort, firma)

# Ergebnis anzeigen
rechner.print_ergebnis(ergebnis)

# In Excel exportieren
rechner.export_to_excel([ergebnis], "beispiel_scores.xlsx")
```

## 📊 Ausgabe

### Konsolen-Ausgabe
- Fortschrittsanzeige während der Berechnung
- Automatische Berechnungen werden angezeigt
- Zusammenfassung mit Top-Firmen
- Hinweise auf Firmen mit Verbesserungspotential

### Excel-Ausgabe
- Ein Tabellenblatt pro Firma
- Formatierte Darstellung aller Parameter
- Automatische Kategorie-Zuweisung
- Berechnete Durchschnitts- und Gesamt-Scores
- Gespeichert im `output/` Ordner mit Zeitstempel

## 🔍 Bewertungskriterien

### Standort-Parameter (je 1-5 Punkte)
1. **ÖV-Anbindungsqualität**: A=5, B=4, C=3, D=2, E=1
2. **Beschäftigte pro 1000 Einwohner**: <300=5, 300-500=4, 501-700=3, 701-900=2, >900=1
3. **Einpendler prozentual**: <40%=5, 40-50%=4, 51-60%=5, 61-70%=2, >70%=1
4. **Motorisierungsgrad**: <500=5, 500-600=4, 601-700=3, 701-800=2, >800=1
5. **Modal-Split (Autopendler)**: <40%=5, 40-50%=4, 51-60%=3, 61-70%=2, >70%=1

### Firmen-Parameter (je 1-5 Punkte)
1. **Mitarbeiterzahl**: <50=5, 50-100=4, 101-250=3, 251-500=2, >500=1
2. **ÖV-Anbindung**: <300m=5, 300-500m=4, 501-750m=3, 751-1000m=2, >1000m=1
3. **Branche**: IT=5, Finanzen=4, Verwaltung=3, Industrie=2, Logistik=1
4. **Autobahnauffahrt**: >5000m=5, 3001-5000m=4, 2001-3000m=3, 1000-2000m=2, <1000m=1
5. **Parkplatz**: >500m=5, 301-500m=4, 201-300m=3, 100-200m=2, <100m=1

### Score-Berechnung
- **Standort-Score** = Durchschnitt der Standort-Parameter
- **Firmen-Score** = Durchschnitt der Firmen-Parameter
- **Gesamt-Score** = (Standort-Score + Firmen-Score) / 2

## ⚠️ Troubleshooting

### OSM/Overpass-Fehler
- Prüfen Sie Ihre Internetverbindung
- Bei Timeout-Fehlern: Warten und erneut versuchen
- Alternative: Manuelle Werte in config.json eintragen

### Fehlende CSV-Dateien
- Stellen Sie sicher, dass `oev_qualitaet_gemeinden.csv` und `Betriebspunkt.csv` im `data/` Verzeichnis liegen
- Prüfen Sie die Spaltennamen in den CSV-Dateien:
  - `oev_qualitaet_gemeinden.csv`: Spalten `bfs_nummer`, `gemeinde`, `mean_score`
  - `Betriebspunkt.csv`: Spalten `Name`, `E`, `N`

### Installation von geopandas
Falls Probleme bei der Installation von geopandas auftreten:
```bash
# Windows (mit Anaconda)
conda install -c conda-forge geopandas

# Linux/Mac
sudo apt-get install gdal-bin libgdal-dev  # Ubuntu/Debian
pip install geopandas
```

### Fehlende Module
Das Skript prüft automatisch beim Start, ob alle benötigten Module installiert sind und gibt eine entsprechende Fehlermeldung aus.

## 📝 Hinweise zur Verwendung

- Die CSV-Dateien müssen im `data/` Verzeichnis liegen
- Die Ausgabe-Excel-Dateien werden im `output/` Verzeichnis gespeichert
- Bei der ersten Ausführung werden die Verzeichnisse automatisch erstellt
- Die ÖV-Güteklasse wird automatisch aus der CSV ermittelt
- Alle Berechnungen werden in der Konsole protokolliert