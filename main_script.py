#!/usr/bin/env python3
"""
Hauptskript für die Berechnung von Standort- und Firmenbewertungen
==================================================================

Dieses Skript liest die Konfiguration aus config.json und berechnet
automatisch alle Scores für die definierten Firmen.

Verwendung:
-----------
1. Passen Sie die config.json an Ihre Bedürfnisse an
2. Stellen Sie sicher, dass die CSV-Dateien vorhanden sind:
   - oev_qualitaet_gemeinden.csv
   - Betriebspunkt.csv
3. Führen Sie das Skript aus: python main.py
"""

import json
import sys
from pathlib import Path
from datetime import datetime

# Import der Haupt-Rechner-Klasse
from standort_score_rechner import StandortScoreRechner, Standort, Firma

def load_config(config_file: str = "config.json"):
    """Lädt die Konfiguration aus JSON-Datei"""
    try:
        with open(config_file, 'r', encoding='utf-8') as f:
            return json.load(f)
    except FileNotFoundError:
        print(f"❌ Konfigurationsdatei '{config_file}' nicht gefunden!")
        sys.exit(1)
    except json.JSONDecodeError as e:
        print(f"❌ Fehler beim Lesen der Konfiguration: {e}")
        sys.exit(1)

def main():
    """Hauptfunktion"""
    print("=" * 80)
    print("STANDORT- UND FIRMEN-SCORE-RECHNER")
    print("=" * 80)
    print(f"Startzeit: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n")

    # Konfiguration laden
    print("📁 Lade Konfiguration...")
    config = load_config()

    # Rechner initialisieren
    print("🔧 Initialisiere Rechner...")
    rechner = StandortScoreRechner()

    # Alle Ergebnisse sammeln
    alle_ergebnisse = []
    fehler_count = 0

    print(f"\n📊 Berechne Scores für {len(config['firmen'])} Firmen...\n")

    for firma_config in config['firmen']:
        try:
            print(f"🏢 Verarbeite {firma_config['name']}...")

            # Standort-Objekt erstellen
            standort_name = firma_config['standort']
            if standort_name not in config['standorte']:
                print(f"   ⚠️  Standort '{standort_name}' nicht in Konfiguration gefunden!")
                fehler_count += 1
                continue

            standort_data = config['standorte'][standort_name]

            # Erstelle Standort-Objekt (alle Berechnungen erfolgen automatisch)
            standort = Standort(
                name=standort_name,
                anzahl_beschaeftigte=standort_data['anzahl_beschaeftigte'],
                anzahl_einwohner=standort_data['anzahl_einwohner'],
                anzahl_einpendelnde=standort_data['anzahl_einpendelnde'],
                motorisierungsgrad=standort_data['motorisierungsgrad'],
                modal_split_auto=standort_data['modal_split_auto']
            )

            # Firmen-Objekt erstellen
            firma = Firma(
                name=firma_config['name'],
                adresse=firma_config['adresse'],
                lat=firma_config['lat'],
                lon=firma_config['lon'],
                mitarbeiterzahl=firma_config['mitarbeiterzahl'],
                branche=firma_config['branche']
            )

            # Score berechnen
            ergebnis = rechner.berechne_scores(standort, firma)
            alle_ergebnisse.append(ergebnis)

            # Kurzübersicht ausgeben
            print(f"   ✅ Gesamt-Score: {ergebnis['scores']['gesamt_score']} "
                  f"(Standort: {ergebnis['scores']['standort_score']}, "
                  f"Firma: {ergebnis['scores']['firmen_score']})")

        except Exception as e:
            print(f"   ❌ Fehler bei {firma_config['name']}: {str(e)}")
            fehler_count += 1

    # Ergebnisse exportieren
    if alle_ergebnisse:
        output_file = f"standort_scores_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
        print(f"\n💾 Exportiere Ergebnisse nach {output_file}...")
        rechner.export_to_excel(alle_ergebnisse, output_file)
        print(f"✅ Export erfolgreich!")

        # Zusammenfassung
        print("\n" + "=" * 80)
        print("ZUSAMMENFASSUNG")
        print("=" * 80)
        print(f"Erfolgreich verarbeitet: {len(alle_ergebnisse)} Firmen")
        print(f"Fehler: {fehler_count}")

        # Top 5 Firmen nach Gesamt-Score
        print("\n🏆 TOP 5 FIRMEN (nach Gesamt-Score):")
        sorted_results = sorted(alle_ergebnisse,
                              key=lambda x: x['scores']['gesamt_score'],
                              reverse=True)
        for i, result in enumerate(sorted_results[:5], 1):
            print(f"{i}. {result['firma']} ({result['standort']}): "
                  f"Score {result['scores']['gesamt_score']}")

        # Schlechteste 3 Firmen
        print("\n⚠️  FIRMEN MIT VERBESSERUNGSPOTENTIAL:")
        for result in sorted_results[-3:]:
            print(f"- {result['firma']} ({result['standort']}): "
                  f"Score {result['scores']['gesamt_score']}")

    else:
        print("\n❌ Keine Ergebnisse zum Exportieren!")

    print(f"\nEndzeit: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    print("=" * 80)

if __name__ == "__main__":
    # Prüfe ob notwendige Module installiert sind
    required_modules = [
        'pandas', 'numpy', 'geopandas', 'shapely',
        'osmnx', 'overpy', 'haversine', 'pyproj'
    ]

    missing_modules = []
    for module in required_modules:
        try:
            __import__(module)
        except ImportError:
            missing_modules.append(module)

    if missing_modules:
        print("❌ Fehlende Python-Module:")
        print(f"   Bitte installieren Sie: {', '.join(missing_modules)}")
        print(f"\n   pip install {' '.join(missing_modules)}")
        sys.exit(1)

    # Hauptprogramm ausführen
    main()