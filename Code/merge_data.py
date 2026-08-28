"""Zusammenführung (Merge) der Eurostat-Rohdaten.

Erzeugt aus den Rohdaten-Exporten drei finale Excel-Tabellen im Output-Pfad:

1. ``merge_industrieproduktion_erzeugerpreise.xlsx``
   Merge A: sts_inpr_m + sts_inppd_m
2. ``merge_energiepreise.xlsx``
   Merge B: nrg_pc_205 + nrg_pc_203
3. ``lc_lci_lev_final.xlsx``
   Arbeitskosten (kein Merge-Partner vorgesehen, daher 1:1 finalisiert)

Alle finalen Tabellen werden zusätzlich bereinigt (siehe
``bereinige_tabelle``): nur Euro-Werte, und Code-Spalten, die durch eine
zugehörige ``*_label``-Spalte lang erklärt werden, werden entfernt.

Ausführen (nach fetch_data.py):
    python merge_data.py
"""

import json
from datetime import datetime

import pandas as pd

import config
from fetch_data import export_excel

logger = config.setup_logging()


def load_raw(dataset_code: str) -> pd.DataFrame:
    """Liest eine Rohdaten-Excel-Datei aus dem Output-Pfad ein.

    Raises:
        FileNotFoundError: Mit Hinweis auf den notwendigen Datenabruf,
            falls die Datei fehlt.
    """
    pfad = config.OUTPUT_DIR / config.DATASETS[dataset_code]["datei"]
    if not pfad.exists():
        raise FileNotFoundError(
            f"Rohdaten-Datei nicht gefunden: {pfad}. "
            f"Bitte zuerst 'python fetch_data.py' ausführen."
        )
    return pd.read_excel(pfad)


def merge_frames(
    links: pd.DataFrame,
    rechts: pd.DataFrame,
    keys: list,
    wert_links: str,
    wert_rechts: str,
) -> pd.DataFrame:
    """Führt zwei Long-Format-DataFrames per Outer Join zusammen.

    Begründung Outer Join (statt Inner Join): Die Abdeckung der beiden
    Datensätze unterscheidet sich (z. B. hat sts_inppd_m nur s_adj='NSA',
    während sts_inpr_m zusätzlich 'CA'/'SCA' liefert; die Länder- und
    Zeitraumabdeckung weicht leicht ab). Ein Inner Join würde diese
    Beobachtungen stillschweigend verwerfen; der Outer Join erhält alle
    Werte und kennzeichnet fehlende Partner durch leere Zellen (NaN).

    Die Eindeutigkeit der Schlüssel wird per ``validate='one_to_one'``
    erzwungen – eine Zeilenverdopplung durch den Join ist damit
    ausgeschlossen (Verstoß löst eine Exception aus).

    Nicht-Schlüsselspalten, die in beiden Tabellen vorkommen (z. B.
    ``geo_label``, ``siec``, ``time_date``), werden aus der linken Tabelle
    übernommen und für nur rechts vorkommende Zeilen aus der rechten
    Tabelle aufgefüllt (combine_first).

    Args:
        links/rechts: Eingabe-DataFrames mit einer ``value``-Spalte.
        keys: Gemeinsame Schlüsselspalten des Joins.
        wert_links/wert_rechts: Neue Spaltennamen der Werte zur
            Herkunftskennzeichnung (z. B. ``value_inpr``/``value_inppd``).
    """
    for name, df in (("links", links), ("rechts", rechts)):
        if df.duplicated(keys).any():
            doppelte = df[df.duplicated(keys, keep=False)][keys].head(5)
            raise ValueError(
                f"Schlüssel nicht eindeutig ({name}): Dopplungen bei\n{doppelte}"
            )

    gemeinsam = [
        spalte
        for spalte in links.columns
        if spalte in rechts.columns and spalte not in keys and spalte != "value"
    ]
    rechts_teil = rechts[keys + ["value"] + gemeinsam].rename(
        columns={"value": wert_rechts, **{c: f"{c}__rechts" for c in gemeinsam}}
    )
    merged = links.rename(columns={"value": wert_links}).merge(
        rechts_teil, on=keys, how="outer", validate="one_to_one"
    )
    for spalte in gemeinsam:
        merged[spalte] = merged[spalte].combine_first(
            merged.pop(f"{spalte}__rechts")
        )
    return merged.sort_values(keys).reset_index(drop=True)


def merge_industrie(
    inpr: pd.DataFrame | None = None, inppd: pd.DataFrame | None = None
) -> pd.DataFrame:
    """Merge A: Industrieproduktionsindex + Erzeugerpreisindex.

    Schlüssel: time, geo, nace_r2, unit sowie freq und s_adj. Die Dimension
    ``indic_bt`` ist bewusst NICHT Teil des Schlüssels: Sie ist je Datensatz
    konstant (PRD bzw. PRC_PRR_DOM) und würde als Schlüssel jeden
    tatsächlichen Treffer verhindern. ``s_adj`` gehört dagegen in den
    Schlüssel, damit saisonbereinigte Reihen nicht mit unbereinigten
    Erzeugerpreisen kombiniert werden (Erzeugerpreise liegen nur als NSA vor).
    """
    inpr = load_raw("sts_inpr_m") if inpr is None else inpr
    inppd = load_raw("sts_inppd_m") if inppd is None else inppd
    merged = merge_frames(
        inpr, inppd, config.MERGE_A_KEYS, "value_inpr", "value_inppd"
    )
    logger.info(
        "Merge A (Industrieproduktion + Erzeugerpreise): %d + %d Zeilen -> "
        "%d Zeilen (Schlüssel: %s)",
        len(inpr), len(inppd), len(merged), ", ".join(config.MERGE_A_KEYS),
    )
    return merged


def merge_energie(
    strom: pd.DataFrame | None = None, gas: pd.DataFrame | None = None
) -> pd.DataFrame:
    """Merge B: Strompreise + Gaspreise (Nicht-Haushalte).

    Schlüssel: time, geo, unit sowie die übrigen gemeinsamen Dimensionen
    freq, nrg_cons, tax und currency. Das Verbrauchsband (nrg_cons) wird
    in den Schlüssel aufgenommen, obwohl die Codierung zwischen Strom
    (MWh-Bänder) und Gas (GJ-Bänder) NICHT vergleichbar ist: Ein Weglassen
    würde jeden Strom-Datensatz mit jedem Gas-Datensatz kreuzen
    (kartesisches Produkt). Durch die disjunkten Band-Codes ergibt der
    Outer Join faktisch eine Union beider Datensätze ohne
    Zeilenmultiplikation – Strom- und Gaswerte stehen je Zeile in den
    getrennten Spalten ``value_strom``/``value_gas``.

    Die Hilfsspalte ``energietraeger`` kennzeichnet die Herkunft je Zeile.
    """
    strom = load_raw("nrg_pc_205") if strom is None else strom
    gas = load_raw("nrg_pc_203") if gas is None else gas
    merged = merge_frames(
        strom, gas, config.MERGE_B_KEYS, "value_strom", "value_gas"
    )
    hat_strom = merged["value_strom"].notna()
    hat_gas = merged["value_gas"].notna()
    merged["energietraeger"] = "Strom & Gas"
    merged.loc[hat_strom & ~hat_gas, "energietraeger"] = "Strom"
    merged.loc[~hat_strom & hat_gas, "energietraeger"] = "Gas"
    logger.info(
        "Merge B (Strom + Gas): %d + %d Zeilen -> %d Zeilen "
        "(davon %d Strom-, %d Gas-Zeilen; Schlüssel: %s)",
        len(strom), len(gas), len(merged),
        int(hat_strom.sum()), int(hat_gas.sum()),
        ", ".join(config.MERGE_B_KEYS),
    )
    return merged


def bereinige_tabelle(df: pd.DataFrame) -> pd.DataFrame:
    """Bereinigt eine finale Tabelle für die Excel-Ablage.

    1. Währungsfilter: Nur Euro-Werte bleiben erhalten (falls der Datensatz
       eine ``currency``-Dimension hat – betrifft die Energiepreise; PPS-
       und nationale Währungen werden verworfen).
    2. NACE-Darstellung: Der numerische NACE-Code wird in die Bezeichnung
       aufgenommen (``nace_r2_label`` wird zu 'C2221 – Herstellung von …'),
       damit die Kennung nach dem Entfernen der Code-Spalte sichtbar bleibt.
    3. Spaltenreduktion: Code-Spalten, die nur Abkürzungen enthalten und
       durch eine zugehörige ``*_label``-Spalte lang erklärt werden, werden
       gelöscht (die Label-Spalten bleiben erhalten). Ausnahme: ``time``
       bleibt als kompakter Zeitschlüssel bestehen (Periodenangaben wie
       '2024-S1' sind keine Abkürzungen); ``time_label`` entfällt dagegen
       als redundant, da es bei Eurostat-Zeitdimensionen identisch mit
       ``time`` ist.
    """
    if "currency" in df.columns:
        vorher = len(df)
        df = df[df["currency"] == "EUR"].copy()
        logger.info(
            "Währungsfilter EUR: %d -> %d Zeilen", vorher, len(df)
        )
    if "nace_r2" in df.columns and "nace_r2_label" in df.columns:
        df = df.copy()
        kombiniert = (
            df["nace_r2"].astype(str) + " – " + df["nace_r2_label"].astype(str)
        )
        df["nace_r2_label"] = kombiniert.where(
            df["nace_r2"].notna(), df["nace_r2_label"]
        )
    zu_loeschen = [
        spalte
        for spalte in df.columns
        if not spalte.endswith("_label")
        and f"{spalte}_label" in df.columns
        and spalte != "time"
    ]
    if "time_label" in df.columns:
        zu_loeschen.append("time_label")
    if zu_loeschen:
        logger.info("Entfernte Code-Spalten: %s", ", ".join(zu_loeschen))
        df = df.drop(columns=zu_loeschen)
    return df


def finalisiere_arbeitskosten() -> pd.DataFrame:
    """Finalisiert die dritte Tabelle: Arbeitskosten ohne Merge-Partner.

    Die Rohdaten sind bereits vollständig und gefiltert; sie werden daher
    lediglich bereinigt (siehe ``bereinige_tabelle``) und als
    ``lc_lci_lev_final.xlsx`` exportiert.
    """
    df = bereinige_tabelle(load_raw("lc_lci_lev"))
    logger.info("Arbeitskosten finalisiert: %d Zeilen", len(df))
    return df


def _schreibe_merge_metadata() -> None:
    """Ergänzt den Merge-Zeitpunkt in ``fetch_metadata.json``."""
    meta = {}
    if config.FETCH_METADATA_FILE.exists():
        try:
            meta = json.loads(
                config.FETCH_METADATA_FILE.read_text(encoding="utf-8")
            )
        except (json.JSONDecodeError, OSError):
            meta = {}
    meta["_merge"] = {
        "merge_zeitpunkt": datetime.now().isoformat(timespec="seconds"),
        "dateien": [
            config.FINAL_LC,
            config.FINAL_MERGE_INDUSTRIE,
            config.FINAL_MERGE_ENERGIE,
        ],
    }
    config.FETCH_METADATA_FILE.write_text(
        json.dumps(meta, ensure_ascii=False, indent=2), encoding="utf-8"
    )


def main() -> None:
    """Erzeugt alle drei finalen Excel-Tabellen im Output-Pfad."""
    merge_a = bereinige_tabelle(merge_industrie())
    export_excel(merge_a, config.OUTPUT_DIR / config.FINAL_MERGE_INDUSTRIE)
    merge_b = bereinige_tabelle(merge_energie())
    export_excel(merge_b, config.OUTPUT_DIR / config.FINAL_MERGE_ENERGIE)
    arbeitskosten = finalisiere_arbeitskosten()
    export_excel(arbeitskosten, config.OUTPUT_DIR / config.FINAL_LC)
    _schreibe_merge_metadata()
    logger.info("Merge abgeschlossen – 3 finale Tabellen im Output-Pfad.")


if __name__ == "__main__":
    main()
