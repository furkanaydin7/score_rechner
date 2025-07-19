#!/usr/bin/env python3
"""
Integrierter Standort- und Firmen-Score-Rechner
================================================
Berechnet automatisch alle Scores basierend auf Standort- und Firmenparametern.
Integriert die bestehenden Python-Skripte für ÖV, Autobahn und Parkplätze.
"""

import pandas as pd
import numpy as np
from dataclasses import dataclass
from typing import Dict, Tuple, Optional, List
import json
from pathlib import Path

# Für die integrierten Funktionen
import shapely.geometry as sh
import geopandas as gpd
import osmnx as ox
import overpy
from haversine import haversine, Unit
from pyproj import Transformer

# =============================================================================
# DATENKLASSEN FÜR STRUKTURIERTE EINGABE
# =============================================================================

@dataclass
class Standort:
    """Standortdaten einer Gemeinde/Stadt"""
    name: str

    # Rohdaten (müssen angegeben werden)
    anzahl_beschaeftigte: float
    anzahl_einwohner: float
    anzahl_einpendelnde: float
    motorisierungsgrad: float
    modal_split_auto: float

    # Berechnete Werte (werden automatisch ermittelt)
    oev_gueteklasse: Optional[str] = None
    beschaeftigte_pro_1000: Optional[float] = None
    einpendler_prozent: Optional[float] = None

    def __post_init__(self):
        """Berechnet abgeleitete Werte nach der Initialisierung"""
        # Berechne beschaeftigte_pro_1000
        self.beschaeftigte_pro_1000 = (self.anzahl_beschaeftigte / self.anzahl_einwohner) * 1000
        print(f"  → Berechnet: Beschäftigte pro 1000 = {self.beschaeftigte_pro_1000:.1f}")

        # Berechne einpendler_prozent
        self.einpendler_prozent = (self.anzahl_einpendelnde / self.anzahl_beschaeftigte) * 100
        print(f"  → Berechnet: Einpendler % = {self.einpendler_prozent:.2f}%")

@dataclass
class Firma:
    """Firmendaten"""
    name: str
    adresse: str
    lat: float
    lon: float
    mitarbeiterzahl: int
    branche: str

# =============================================================================
# KATEGORIE-DEFINITIONEN UND PUNKTZUWEISUNGEN
# =============================================================================

KATEGORIEN = {
    'oev_anbindungsqualitaet': {
        'A': 5, 'B': 4, 'C': 3, 'D': 2, 'E': 1
    },
    'beschaeftigte_pro_1000': {
        '< 300': 5,
        '300–500': 4,
        '501–700': 3,
        '701–900': 2,
        '> 900': 1
    },
    'einpendler_prozent': {
        '< 40 %': 5,
        '40–50 %': 4,
        '51–60 %': 3,
        '61–70 %': 2,
        '> 70 %': 1
    },
    'motorisierungsgrad': {
        '< 500': 5,
        '500–600': 4,
        '601–700': 3,
        '701–800': 2,
        '> 800': 1
    },
    'modal_split': {
        '< 40%': 5,
        '40–50%': 4,
        '51–60%': 3,
        '61–70%': 2,
        '> 70%': 1
    },
    'mitarbeiterzahl': {
        '< 50': 5,
        '50–100': 4,
        '101–250': 3,
        '251–500': 2,
        '> 500': 1
    },
    'oev_naechste_haltestelle': {
        '< 300 m': 5,
        '300–500 m': 4,
        '501–750 m': 3,
        '751–1000 m': 2,
        '> 1000 m': 1
    },
    'branche': {
        'IT & Software': 5,
        'Finanzen, Versicherungen, Beratung': 4,
        'Verwaltung, Bildung, Gesundheitswesen, Dienstleisungen': 3,
        'Industrie, Produktion & Handel': 2,
        'Logistik & Transport': 1
    },
    'autobahn_distanz': {
        '> 5000 m': 5,
        '3001–5000 m': 4,
        '2001–3000 m': 3,
        '1000–2000 m': 2,
        '< 1000 m': 1
    },
    'parkplatz_distanz': {
        '> 500 m': 5,
        '301–500 m': 4,
        '201–300 m': 3,
        '100–200 m': 2,
        '< 100 m': 1
    }
}

# =============================================================================
# INTEGRIERTE FUNKTIONEN AUS DEN BESTEHENDEN SKRIPTEN
# =============================================================================

class TransportAnalyzer:
    """Integriert alle Transportanalyse-Funktionen"""

    def __init__(self):
        self.lv95_to_wgs84 = Transformer.from_crs(2056, 4326, always_xy=True)
        self.overpass_url = "http://overpass.osm.ch/api/interpreter"

    def get_oev_gueteklasse(self, gemeinde_name: str, csv_path: str = "oev_qualitaet_gemeinden.csv") -> Tuple[str, float]:
        """
        Holt ÖV-Güteklasse aus CSV
        Returns: (Güteklasse, mean_score)
        """
        try:
            df = pd.read_csv(csv_path, dtype={"bfs_nummer": str})
            treffer = df[df["gemeinde"].str.contains(gemeinde_name, case=False, regex=False)]

            if treffer.empty:
                raise ValueError(f"Gemeinde '{gemeinde_name}' nicht gefunden")

            row = treffer.iloc[0]
            mean = row.mean_score

            if mean >= 4.5: kl = "A"
            elif mean >= 3.5: kl = "B"
            elif mean >= 2.5: kl = "C"
            elif mean >= 1.5: kl = "D"
            else: kl = "E"

            print(f"  → ÖV-Güteklasse für {row.gemeinde}: {kl} (Score: {mean:.2f})")

            return kl, mean
        except Exception as e:
            print(f"Fehler beim Lesen der ÖV-Güteklasse: {e}")
            return "C", 3.0  # Standardwert

    def get_naechste_haltestelle(self, lat: float, lon: float, csv_path: str = "Betriebspunkt.csv") -> Tuple[str, float]:
        """
        Findet nächste ÖV-Haltestelle
        Returns: (Name, Distanz in Metern)
        """
        try:
            df = pd.read_csv(csv_path, usecols=["Name", "E", "N"])

            # LV95 zu WGS84 konvertieren
            lats, lons = [], []
            for _, row in df.iterrows():
                wgs_lon, wgs_lat = self.lv95_to_wgs84.transform(row["E"], row["N"])
                lats.append(wgs_lat)
                lons.append(wgs_lon)

            df["lat"] = lats
            df["lon"] = lons

            # Distanzen berechnen
            ziel = (lat, lon)
            df["dist"] = df.apply(
                lambda r: haversine((r["lat"], r["lon"]), ziel, unit=Unit.METERS),
                axis=1
            )

            nearest = df.nsmallest(1, "dist").iloc[0]
            return nearest["Name"], nearest["dist"]
        except Exception as e:
            print(f"Fehler bei Haltestellensuche: {e}")
            return "Unbekannt", 500.0  # Standardwert

    def get_naechste_autobahnauffahrt(self, lat: float, lon: float) -> Tuple[str, float]:
        """
        Findet nächste Autobahnauffahrt via Overpass API
        Returns: (Name, Distanz in Metern)
        """
        try:
            api = overpy.Overpass(url=self.overpass_url)
            query = f"""
            [out:json][timeout:25];
            way(around:20000,{lat},{lon})["highway"~"motorway|motorway_link"]->.allways;
            (.allways; >;);
            out body;
            """

            result = api.query(query)

            # Nodes klassifizieren
            node_lookup = {n.id: n for n in result.nodes}
            node_types = {nid: [] for nid in node_lookup}
            motorway_links = []

            for way in result.ways:
                hwy = way.tags.get("highway", "")
                for n in way.nodes:
                    node_types[n.id].append(hwy)
                if hwy == "motorway_link":
                    motorway_links.append(way)

            # Entry points finden
            entries = {}
            for w in motorway_links:
                first_id, last_id = w.nodes[0].id, w.nodes[-1].id
                if "motorway" not in node_types[first_id]:
                    entries[first_id] = w
                if "motorway" not in node_types[last_id]:
                    entries[last_id] = w

            if not entries:
                return "Keine gefunden", 10000.0

            # Nächste finden
            min_dist = float('inf')
            best_name = "Unbekannt"

            for nid, way in entries.items():
                n = node_lookup[nid]
                name = (way.tags.get("name") or way.tags.get("ref") or
                       f"Auffahrt {way.id}")
                dist = haversine((lat, lon), (n.lat, n.lon), unit=Unit.METERS)

                if dist < min_dist:
                    min_dist = dist
                    best_name = name

            return best_name, min_dist
        except Exception as e:
            print(f"Fehler bei Autobahnsuche: {e}")
            return "Unbekannt", 3000.0  # Standardwert

    def get_naechster_parkplatz(self, lat: float, lon: float) -> Tuple[str, float]:
        """
        Findet nächsten Parkplatz via OSM
        Returns: (Name, Distanz in Metern)
        """
        try:
            firm = (lat, lon)
            tags = {"amenity": "parking"}

            gdf = ox.features_from_point(firm, tags=tags, dist=1000)
            if gdf.empty:
                return "Kein Parkplatz gefunden", 1000.0

            # In metrisches CRS projizieren
            gdf = gdf.to_crs(2056)
            firm_pt = gpd.GeoSeries([sh.Point(lon, lat)], crs="EPSG:4326").to_crs(2056).iloc[0]
            gdf["dist_m"] = gdf.geometry.distance(firm_pt)

            nearest = gdf.nsmallest(1, "dist_m").iloc[0]
            name = nearest.get('name', 'Parkplatz ohne Namen')

            return name, nearest['dist_m']
        except Exception as e:
            print(f"Fehler bei Parkplatzsuche: {e}")
            return "Unbekannt", 200.0  # Standardwert

# =============================================================================
# HAUPTKLASSE FÜR SCORE-BERECHNUNG
# =============================================================================

class StandortScoreRechner:
    """Berechnet alle Scores für Standort und Firma"""

    def __init__(self):
        self.analyzer = TransportAnalyzer()

    def kategorie_zuweisen(self, wert: float, typ: str) -> str:
        """Weist basierend auf dem Wert die passende Kategorie zu"""

        if typ == 'beschaeftigte_pro_1000':
            if wert < 300: return '< 300'
            elif wert <= 500: return '300–500'
            elif wert <= 700: return '501–700'
            elif wert <= 900: return '701–900'
            else: return '> 900'

        elif typ == 'einpendler_prozent':
            if wert < 40: return '< 40 %'
            elif wert <= 50: return '40–50 %'
            elif wert <= 60: return '51–60 %'
            elif wert <= 70: return '61–70 %'
            else: return '> 70 %'

        elif typ == 'motorisierungsgrad':
            if wert < 500: return '< 500'
            elif wert <= 600: return '500–600'
            elif wert <= 700: return '601–700'
            elif wert <= 800: return '701–800'
            else: return '> 800'

        elif typ == 'modal_split':
            if wert < 40: return '< 40%'
            elif wert <= 50: return '40–50%'
            elif wert <= 60: return '51–60%'
            elif wert <= 70: return '61–70%'
            else: return '> 70%'

        elif typ == 'mitarbeiterzahl':
            if wert < 50: return '< 50'
            elif wert <= 100: return '50–100'
            elif wert <= 250: return '101–250'
            elif wert <= 500: return '251–500'
            else: return '> 500'

        elif typ == 'oev_naechste_haltestelle':
            if wert < 300: return '< 300 m'
            elif wert <= 500: return '300–500 m'
            elif wert <= 750: return '501–750 m'
            elif wert <= 1000: return '751–1000 m'
            else: return '> 1000 m'

        elif typ == 'autobahn_distanz':
            if wert < 1000: return '< 1000 m'
            elif wert <= 2000: return '1000–2000 m'
            elif wert <= 3000: return '2001–3000 m'
            elif wert <= 5000: return '3001–5000 m'
            else: return '> 5000 m'

        elif typ == 'parkplatz_distanz':
            if wert < 100: return '< 100 m'
            elif wert <= 200: return '100–200 m'
            elif wert <= 300: return '201–300 m'
            elif wert <= 500: return '301–500 m'
            else: return '> 500 m'

    def berechne_scores(self, standort: Standort, firma: Firma) -> Dict:
        """Berechnet alle Scores für eine Firma an einem Standort"""

        # ÖV-Güteklasse automatisch ermitteln
        print(f"\n🔍 Ermittle ÖV-Güteklasse für {standort.name}...")
        standort.oev_gueteklasse, _ = self.analyzer.get_oev_gueteklasse(standort.name)

        ergebnis = {
            'firma': firma.name,
            'adresse': f"{firma.adresse} ({firma.lat}, {firma.lon})",
            'standort': standort.name,
            'standort_parameter': {},
            'firmen_parameter': {},
            'scores': {}
        }

        # ===== STANDORT-PARAMETER =====

        # ÖV-Anbindungsqualität (aus CSV oder manuell)
        oev_klasse = standort.oev_gueteklasse
        oev_punkte = KATEGORIEN['oev_anbindungsqualitaet'][oev_klasse]
        ergebnis['standort_parameter']['oev_anbindungsqualitaet'] = {
            'wert': f"{oev_klasse}",
            'kategorie': oev_klasse,
            'punkte': oev_punkte
        }

        # Beschäftigte pro 1000 Einwohnende
        besch_kat = self.kategorie_zuweisen(standort.beschaeftigte_pro_1000, 'beschaeftigte_pro_1000')
        besch_punkte = KATEGORIEN['beschaeftigte_pro_1000'][besch_kat]
        ergebnis['standort_parameter']['beschaeftigte_pro_1000'] = {
            'wert': standort.beschaeftigte_pro_1000,
            'kategorie': besch_kat,
            'punkte': besch_punkte
        }

        # Einpendler prozentual
        einp_kat = self.kategorie_zuweisen(standort.einpendler_prozent, 'einpendler_prozent')
        einp_punkte = KATEGORIEN['einpendler_prozent'][einp_kat]
        ergebnis['standort_parameter']['einpendler_prozent'] = {
            'wert': f"{standort.einpendler_prozent:.2f}%",
            'kategorie': einp_kat,
            'punkte': einp_punkte
        }

        # Motorisierungsgrad
        motor_kat = self.kategorie_zuweisen(standort.motorisierungsgrad, 'motorisierungsgrad')
        motor_punkte = KATEGORIEN['motorisierungsgrad'][motor_kat]
        ergebnis['standort_parameter']['motorisierungsgrad'] = {
            'wert': standort.motorisierungsgrad,
            'kategorie': motor_kat,
            'punkte': motor_punkte
        }

        # Modal-Split
        modal_kat = self.kategorie_zuweisen(standort.modal_split_auto, 'modal_split')
        modal_punkte = KATEGORIEN['modal_split'][modal_kat]
        ergebnis['standort_parameter']['modal_split'] = {
            'wert': f"{standort.modal_split_auto}%",
            'kategorie': modal_kat,
            'punkte': modal_punkte
        }

        # ===== FIRMEN-PARAMETER =====

        # Mitarbeiterzahl
        ma_kat = self.kategorie_zuweisen(firma.mitarbeiterzahl, 'mitarbeiterzahl')
        ma_punkte = KATEGORIEN['mitarbeiterzahl'][ma_kat]
        ergebnis['firmen_parameter']['mitarbeiterzahl'] = {
            'wert': firma.mitarbeiterzahl,
            'kategorie': ma_kat,
            'punkte': ma_punkte
        }

        # ÖV-Anbindung (nächste Haltestelle)
        hs_name, hs_dist = self.analyzer.get_naechste_haltestelle(firma.lat, firma.lon)
        hs_kat = self.kategorie_zuweisen(hs_dist, 'oev_naechste_haltestelle')
        hs_punkte = KATEGORIEN['oev_naechste_haltestelle'][hs_kat]
        ergebnis['firmen_parameter']['oev_naechste_haltestelle'] = {
            'wert': f"{hs_dist:.0f} m ({hs_name})",
            'kategorie': hs_kat,
            'punkte': hs_punkte
        }

        # Branche
        branche_punkte = KATEGORIEN['branche'].get(firma.branche, 3)
        ergebnis['firmen_parameter']['branche'] = {
            'wert': firma.branche,
            'kategorie': firma.branche,
            'punkte': branche_punkte
        }

        # Distanz zur nächsten Autobahnauffahrt
        ab_name, ab_dist = self.analyzer.get_naechste_autobahnauffahrt(firma.lat, firma.lon)
        ab_kat = self.kategorie_zuweisen(ab_dist, 'autobahn_distanz')
        ab_punkte = KATEGORIEN['autobahn_distanz'][ab_kat]
        ergebnis['firmen_parameter']['autobahn_distanz'] = {
            'wert': f"{ab_dist:.0f} m ({ab_name})",
            'kategorie': ab_kat,
            'punkte': ab_punkte
        }

        # Entfernung zum nächsten Parkplatz
        pp_name, pp_dist = self.analyzer.get_naechster_parkplatz(firma.lat, firma.lon)
        pp_kat = self.kategorie_zuweisen(pp_dist, 'parkplatz_distanz')
        pp_punkte = KATEGORIEN['parkplatz_distanz'][pp_kat]
        ergebnis['firmen_parameter']['parkplatz_distanz'] = {
            'wert': f"{pp_dist:.0f} m ({pp_name})",
            'kategorie': pp_kat,
            'punkte': pp_punkte
        }

        # ===== SCORES BERECHNEN =====

        # Durchschnittlicher Standort-Score
        standort_punkte = [
            oev_punkte, besch_punkte, einp_punkte, motor_punkte, modal_punkte
        ]
        standort_score = np.mean(standort_punkte)
        ergebnis['scores']['standort_score'] = round(standort_score, 1)

        # Durchschnittlicher Firmen-Score
        firmen_punkte = [
            ma_punkte, hs_punkte, branche_punkte, ab_punkte, pp_punkte
        ]
        firmen_score = np.mean(firmen_punkte)
        ergebnis['scores']['firmen_score'] = round(firmen_score, 1)

        # Gesamt-Score
        gesamt_score = (standort_score + firmen_score) / 2
        ergebnis['scores']['gesamt_score'] = round(gesamt_score, 1)

        return ergebnis

    def export_to_excel(self, ergebnisse: List[Dict], output_file: str = "standort_scores.xlsx"):
        """Exportiert die Ergebnisse in eine Excel-Datei"""
        with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
            for ergebnis in ergebnisse:
                # DataFrame für diese Firma erstellen
                data = []

                # Header
                data.append([f"Firma: {ergebnis['firma']}, {ergebnis['adresse']}", "", "", ""])
                data.append(["", "", "", ""])

                # Standort-Parameter
                data.append([f"Standort-Parameter: {ergebnis['standort']}", "Tatsächlicher Wert", "Kategorie", "Punkte"])

                for param, info in ergebnis['standort_parameter'].items():
                    param_name = param.replace('_', ' ').title()
                    data.append([param_name, info['wert'], info['kategorie'], info['punkte']])

                data.append(["", "", "", ""])
                data.append(["", "", "Durchschnittlicher Score", ergebnis['scores']['standort_score']])
                data.append(["", "", "", ""])

                # Firmen-Parameter
                data.append(["Firmen-Parameter", "Tatsächlicher Wert", "Kategorie", "Punkte"])

                for param, info in ergebnis['firmen_parameter'].items():
                    param_name = param.replace('_', ' ').title()
                    data.append([param_name, info['wert'], info['kategorie'], info['punkte']])

                data.append(["", "", "", ""])
                data.append(["", "", "Durchschnittlicher Score", ergebnis['scores']['firmen_score']])
                data.append(["", "", "", ""])

                # Gesamt-Score
                data.append(["", "", "Gesamt-Score", ergebnis['scores']['gesamt_score']])

                # DataFrame erstellen und in Excel schreiben
                df = pd.DataFrame(data, columns=['A', 'B', 'C', 'D'])
                sheet_name = ergebnis['firma'][:31]  # Excel sheet names max 31 chars
                df.to_excel(writer, sheet_name=sheet_name, index=False, header=False)

    def print_ergebnis(self, ergebnis: Dict):
        """Gibt die Ergebnisse formatiert aus"""
        print(f"\n{'='*60}")
        print(f"FIRMA: {ergebnis['firma']}")
        print(f"Adresse: {ergebnis['adresse']}")
        print(f"Standort: {ergebnis['standort']}")
        print(f"{'='*60}")

        print("\nSTANDORT-PARAMETER:")
        print(f"{'Parameter':<35} {'Wert':<20} {'Kategorie':<15} {'Punkte':<6}")
        print("-" * 80)
        for param, info in ergebnis['standort_parameter'].items():
            param_name = param.replace('_', ' ').title()
            print(f"{param_name:<35} {str(info['wert']):<20} {info['kategorie']:<15} {info['punkte']:<6}")

        print(f"\n{'Durchschnittlicher Standort-Score:':<55} {ergebnis['scores']['standort_score']}")

        print("\n\nFIRMEN-PARAMETER:")
        print(f"{'Parameter':<35} {'Wert':<20} {'Kategorie':<15} {'Punkte':<6}")
        print("-" * 80)
        for param, info in ergebnis['firmen_parameter'].items():
            param_name = param.replace('_', ' ').title()
            print(f"{param_name:<35} {str(info['wert']):<20} {info['kategorie']:<15} {info['punkte']:<6}")

        print(f"\n{'Durchschnittlicher Firmen-Score:':<55} {ergebnis['scores']['firmen_score']}")
        print(f"\n{'GESAMT-SCORE:':<55} {ergebnis['scores']['gesamt_score']}")
        print("=" * 60)

# =============================================================================
# BEISPIEL-VERWENDUNG
# =============================================================================

if __name__ == "__main__":
    # Nur für direkte Tests - normalerweise main.py verwenden
    print("Bitte verwenden Sie 'python main.py' für die normale Ausführung.")
    print("Dieses Skript ist nur das Hauptmodul mit den Klassen.")