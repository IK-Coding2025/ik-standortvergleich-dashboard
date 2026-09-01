"""Unit-Tests für die Standortvergleich-Pipeline.

Schwerpunkte gemäß QS-Vorgaben:
- Merge-Logik: korrekte Zeilenzahl, keine Dopplungen durch den Join
- Validierung der Pflichtfilter (nace_r2, unit) vor dem Export

Ausführen (im Code-Verzeichnis):
    python test_pipeline.py -v

Hinweis: Als Hauptskript starten (nicht über 'python -m unittest'), damit
keine Bytecode-Cache-Datei (.pyc) im FileCloud-Ordner geschrieben wird.
"""

import sys
import unittest

# Keine Bytecode-Caches schreiben (FileCloud-Sync-Konflikte), siehe config.py
sys.dont_write_bytecode = True

import pandas as pd

import config
from fetch_data import (
    parse_jsonstat,
    parse_period_to_date,
    validate_mandatory_filters,
)
from merge_data import (
    bereinige_tabelle,
    merge_energie,
    merge_frames,
    merge_industrie,
)


def _sts_frame(zeilen: list) -> pd.DataFrame:
    """Baut einen synthetischen STS-DataFrame (Struktur wie Rohdaten)."""
    spalten = ["freq", "indic_bt", "nace_r2", "s_adj", "unit", "geo", "time",
               "value"]
    df = pd.DataFrame(zeilen, columns=spalten)
    df["geo_label"] = df["geo"].map({"DE": "Deutschland", "FR": "Frankreich"})
    df["time_date"] = df["time"].map(parse_period_to_date)
    return df


def _nrg_frame(zeilen: list) -> pd.DataFrame:
    """Baut einen synthetischen Energiepreis-DataFrame."""
    spalten = ["freq", "siec", "nrg_cons", "unit", "tax", "currency", "geo",
               "time", "value"]
    df = pd.DataFrame(zeilen, columns=spalten)
    df["geo_label"] = df["geo"].map({"DE": "Deutschland", "FR": "Frankreich"})
    df["time_date"] = df["time"].map(parse_period_to_date)
    return df


class TestParsePeriod(unittest.TestCase):
    """Tests für die Umwandlung von Eurostat-Perioden-Codes."""

    def test_jaehrlich(self):
        self.assertEqual(parse_period_to_date("2020"), pd.Timestamp(2020, 1, 1))

    def test_halbjaehrlich(self):
        self.assertEqual(
            parse_period_to_date("2021-S1"), pd.Timestamp(2021, 1, 1)
        )
        self.assertEqual(
            parse_period_to_date("2021-S2"), pd.Timestamp(2021, 7, 1)
        )

    def test_monatlich(self):
        self.assertEqual(
            parse_period_to_date("2022-11"), pd.Timestamp(2022, 11, 1)
        )

    def test_vierteljaehrlich(self):
        self.assertEqual(
            parse_period_to_date("2020-Q4"), pd.Timestamp(2020, 10, 1)
        )

    def test_unbekannt(self):
        self.assertTrue(pd.isna(parse_period_to_date("irgendwas")))


class TestParseJsonstat(unittest.TestCase):
    """Tests für den JSON-stat-2.0-Parser (Long-Format, Labels, Status)."""

    PAYLOAD = {
        "version": "2.0",
        "class": "dataset",
        "id": ["unit", "geo", "time"],
        "size": [2, 1, 2],
        "dimension": {
            "unit": {
                "category": {
                    "index": {"EUR": 0, "RT_PRE_EUR": 1},
                    "label": {"EUR": "Euro",
                              "RT_PRE_EUR": "Wachstumsrate"},
                }
            },
            "geo": {
                "category": {
                    "index": ["DE"],  # Listenform statt Objekt
                    "label": {"DE": "Deutschland"},
                }
            },
            "time": {
                "category": {
                    "index": {"2020": 0, "2021": 1},
                    "label": {"2020": "2020", "2021": "2021"},
                }
            },
        },
        # sparse: nur 3 von 4 Zellen belegt, eine mit Status-Flag
        "value": {"0": 40.5, "1": 41.2, "3": 2.1},
        "status": {"3": "e"},
    }

    def test_long_format_und_labels(self):
        df = parse_jsonstat(self.PAYLOAD)
        self.assertEqual(len(df), 3)  # nur belegte Zellen
        self.assertIn("geo_label", df.columns)
        self.assertTrue((df["geo_label"] == "Deutschland").all())
        # Keine Dopplungen: Dimensionskombinationen eindeutig
        self.assertFalse(df.duplicated(["unit", "geo", "time"]).any())

    def test_werte_und_status(self):
        df = parse_jsonstat(self.PAYLOAD)
        zelle = df[(df["unit"] == "RT_PRE_EUR") & (df["time"] == "2021")]
        self.assertAlmostEqual(zelle["value"].iloc[0], 2.1)
        self.assertEqual(zelle["status"].iloc[0], "e")

    def test_time_date(self):
        df = parse_jsonstat(self.PAYLOAD)
        self.assertTrue(df["time_date"].notna().all())


class TestFilterValidierung(unittest.TestCase):
    """Tests für die Pflichtfilter-Validierung vor dem Export."""

    def _df(self, nace, unit, zeit):
        return pd.DataFrame({"nace_r2": nace, "unit": unit, "time": zeit})

    def test_korrekte_filter_passieren(self):
        df = self._df(["C", "C"], ["EUR", "RT_PRE_EUR"], ["2020", "2021"])
        validate_mandatory_filters(
            df, "lc_lci_lev", config.DATASETS["lc_lci_lev"]["pflichtfilter"]
        )  # darf nicht werfen

    def test_falscher_nace_wirft(self):
        df = self._df(["C", "D"], ["EUR", "EUR"], ["2020", "2020"])
        with self.assertRaises(ValueError):
            validate_mandatory_filters(
                df, "lc_lci_lev",
                config.DATASETS["lc_lci_lev"]["pflichtfilter"],
            )

    def test_falsche_unit_wirft(self):
        df = self._df(["C"], ["PPS"], ["2020"])
        with self.assertRaises(ValueError):
            validate_mandatory_filters(
                df, "lc_lci_lev",
                config.DATASETS["lc_lci_lev"]["pflichtfilter"],
            )

    def test_zeitraum_vor_2020_wirft(self):
        df = self._df(["C"], ["EUR"], ["2019"])
        with self.assertRaises(ValueError):
            validate_mandatory_filters(
                df, "lc_lci_lev",
                config.DATASETS["lc_lci_lev"]["pflichtfilter"],
            )


class TestMergeFrames(unittest.TestCase):
    """Tests für die generische Merge-Funktion (Outer Join, one-to-one)."""

    def setUp(self):
        self.keys = config.MERGE_A_KEYS
        self.links = _sts_frame([
            ["M", "PRD", "C2221", "NSA", "I21", "DE", "2024-01", 100.0],
            ["M", "PRD", "C2221", "SCA", "I21", "DE", "2024-01", 101.0],
        ])
        self.rechts = _sts_frame([
            ["M", "PRC_PRR_DOM", "C2221", "NSA", "I21", "DE", "2024-01", 110.0],
            ["M", "PRC_PRR_DOM", "C2221", "NSA", "I21", "DE", "2024-02", 111.0],
        ])

    def test_zeilenzahl_und_keine_dopplungen(self):
        merged = merge_frames(
            self.links, self.rechts, self.keys, "value_inpr", "value_inppd"
        )
        # 2 gemeinsame/unpaarige Zeilen links + 1 unpaarige Zeile rechts
        self.assertEqual(len(merged), 3)
        self.assertFalse(merged.duplicated(self.keys).any())

    def test_gemeinsame_zeile_hat_beide_werte(self):
        merged = merge_frames(
            self.links, self.rechts, self.keys, "value_inpr", "value_inppd"
        )
        zeile = merged[(merged["s_adj"] == "NSA") & (merged["time"] == "2024-01")]
        self.assertEqual(len(zeile), 1)
        self.assertAlmostEqual(zeile["value_inpr"].iloc[0], 100.0)
        self.assertAlmostEqual(zeile["value_inppd"].iloc[0], 110.0)

    def test_outer_join_erhaelt_unpaarige_zeilen(self):
        merged = merge_frames(
            self.links, self.rechts, self.keys, "value_inpr", "value_inppd"
        )
        sca = merged[merged["s_adj"] == "SCA"]
        self.assertEqual(len(sca), 1)
        self.assertTrue(sca["value_inppd"].isna().all())
        rechts_only = merged[merged["time"] == "2024-02"]
        self.assertTrue(rechts_only["value_inpr"].isna().all())
        # Label-Auffüllung von rechts: geo_label darf nicht fehlen
        self.assertEqual(rechts_only["geo_label"].iloc[0], "Deutschland")

    def test_doppelte_schluessel_werfen(self):
        links_doppelt = pd.concat([self.links, self.links.iloc[[0]]])
        with self.assertRaises(ValueError):
            merge_frames(
                links_doppelt, self.rechts, self.keys,
                "value_inpr", "value_inppd",
            )


class TestMergeA(unittest.TestCase):
    """Tests für Merge A: Industrieproduktion + Erzeugerpreise."""

    def test_merge_mit_synthetischen_daten(self):
        inpr = _sts_frame([
            ["M", "PRD", "C2221", "NSA", "I21", "DE", "2024-01", 100.0],
            ["M", "PRD", "C2222", "NSA", "I21", "DE", "2024-01", 98.0],
            ["M", "PRD", "C2221", "CA", "I21", "DE", "2024-01", 99.0],
        ])
        inppd = _sts_frame([
            ["M", "PRC_PRR_DOM", "C2221", "NSA", "I21", "DE", "2024-01", 110.0],
            ["M", "PRC_PRR_DOM", "C2222", "NSA", "I21", "DE", "2024-01", 109.0],
        ])
        merged = merge_industrie(inpr, inppd)
        # NSA-Zeilen matchen, CA-Zeile bleibt ohne Erzeugerpreis
        self.assertEqual(len(merged), 3)
        self.assertEqual(int(merged["value_inpr"].notna().sum()), 3)
        self.assertEqual(int(merged["value_inppd"].notna().sum()), 2)
        self.assertFalse(merged.duplicated(config.MERGE_A_KEYS).any())


class TestMergeB(unittest.TestCase):
    """Tests für Merge B: Strom- + Gaspreise (disjunkte Verbrauchsbänder)."""

    def setUp(self):
        self.strom = _nrg_frame([
            ["S", "E7000", "MWH2000-19999", "KWH", "X_TAX", "EUR", "DE",
             "2024-S1", 0.19],
            ["S", "E7000", "MWH2000-19999", "KWH", "X_TAX", "EUR", "DE",
             "2024-S2", 0.20],
        ])
        self.gas = _nrg_frame([
            ["S", "G3000", "GJ10000-99999", "GJ_GCV", "X_TAX", "EUR", "DE",
             "2024-S1", 12.5],
            ["S", "G3000", "GJ10000-99999", "GJ_GCV", "X_TAX", "EUR", "FR",
             "2024-S1", 11.8],
        ])

    def test_union_ohne_zeilenmultiplikation(self):
        merged = merge_energie(self.strom, self.gas)
        # Disjunkte nrg_cons-Codes -> Union: 2 + 2 Zeilen, kein Kreuzprodukt
        self.assertEqual(len(merged), 4)
        self.assertFalse(merged.duplicated(config.MERGE_B_KEYS).any())

    def test_energietraeger_und_spalten(self):
        merged = merge_energie(self.strom, self.gas)
        self.assertIn("value_strom", merged.columns)
        self.assertIn("value_gas", merged.columns)
        self.assertEqual(
            set(merged["energietraeger"].unique()), {"Strom", "Gas"}
        )
        # Herkunft korrekt: Strom-Zeilen haben keinen Gaswert und umgekehrt
        strom_zeilen = merged[merged["energietraeger"] == "Strom"]
        self.assertTrue(strom_zeilen["value_gas"].isna().all())
        gas_zeilen = merged[merged["energietraeger"] == "Gas"]
        self.assertTrue(gas_zeilen["value_strom"].isna().all())
        # siec-Code wird für Gas-Zeilen aus der rechten Tabelle aufgefüllt
        self.assertTrue((gas_zeilen["siec"] == "G3000").all())


class TestBereinigung(unittest.TestCase):
    """Tests für die Bereinigung der finalen Tabellen."""

    def _frame(self) -> pd.DataFrame:
        return pd.DataFrame({
            "currency": ["EUR", "PPS", "NAC", "EUR"],
            "currency_label": ["Euro", "KKS", "Nationale Währung", "Euro"],
            "geo": ["DE", "DE", "DE", "FR"],
            "geo_label": ["Deutschland"] * 3 + ["Frankreich"],
            "time": ["2024-S1"] * 4,
            "time_label": ["2024-S1"] * 4,
            "value": [1.0, 2.0, 3.0, 4.0],
            "status": [None, "e", None, None],
        })

    def test_nur_euro_bleibt(self):
        df = bereinige_tabelle(self._frame())
        self.assertEqual(len(df), 2)  # nur die beiden EUR-Zeilen
        self.assertNotIn("currency", df.columns)
        self.assertIn("currency_label", df.columns)

    def test_code_spalten_entfernt_labels_bleiben(self):
        df = bereinige_tabelle(self._frame())
        self.assertNotIn("geo", df.columns)
        self.assertIn("geo_label", df.columns)

    def test_time_bleibt_time_label_entfaellt(self):
        df = bereinige_tabelle(self._frame())
        self.assertIn("time", df.columns)
        self.assertNotIn("time_label", df.columns)

    def test_spalten_ohne_label_bleiben(self):
        df = bereinige_tabelle(self._frame())
        self.assertIn("status", df.columns)
        self.assertIn("value", df.columns)

    def test_nace_code_wird_in_label_uebernommen(self):
        df = pd.DataFrame({
            "nace_r2": ["C2221"],
            "nace_r2_label": ["Herstellung von Kunststoffplatten"],
            "time": ["2024-01"],
            "value": [1.0],
        })
        out = bereinige_tabelle(df)
        self.assertNotIn("nace_r2", out.columns)
        self.assertEqual(
            out["nace_r2_label"].iloc[0],
            "C2221 – Herstellung von Kunststoffplatten",
        )


if __name__ == "__main__":
    unittest.main()
