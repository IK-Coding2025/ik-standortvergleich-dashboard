"""Abruf der Eurostat-Datensätze über die Dissemination API (JSON-stat 2.0).

Für jeden in ``config.DATASETS`` definierten Datensatz wird eine eigene
Abruffunktion aufgerufen, die das Ergebnis als Long-Format-DataFrame und als
Excel-Datei im Output-Pfad speichert. Der direkte REST-Zugriff via ``requests``
wird (statt der Pakete ``eurostat``/``pandasdmx``) verwendet, weil er volle
Kontrolle über Sprache (deutsche Eurostat-Bezeichnungen), Retry-Logik und
Logging bietet.

Ausführen:
    python fetch_data.py
"""

import sys

# Keine Bytecode-Caches schreiben (FileCloud-Sync-Konflikte), siehe config.py
sys.dont_write_bytecode = True

import json
import re
import time
from datetime import datetime
from pathlib import Path

import pandas as pd
import requests

import config

logger = config.setup_logging()


# ---------------------------------------------------------------------------
# API-Abruf mit Retry-Logik
# ---------------------------------------------------------------------------
def fetch_jsonstat(dataset_code: str, filter_params: dict) -> dict:
    """Ruft einen Eurostat-Datensatz als JSON-stat 2.0 ab.

    Args:
        dataset_code: Eurostat-Datensatzcode, z. B. ``"lc_lci_lev"``.
        filter_params: Mapping Dimensions-ID -> Liste erlaubter Codes,
            z. B. ``{"nace_r2": ["C"], "unit": ["EUR", "RT_PRE_EUR"]}``.

    Returns:
        Das JSON-stat-2.0-Dokument als ``dict``.

    Raises:
        RuntimeError: Wenn nach ``config.MAX_RETRIES`` Versuchen kein
            erfolgreicher Abruf möglich war oder ein Client-Fehler (4xx)
            vorliegt.
    """
    url = f"{config.EUROSTAT_BASE_URL}/{dataset_code}"
    # Parameter als Liste von Tupeln, damit mehrfache Werte je Dimension
    # als wiederholte Query-Parameter gesendet werden (z. B. unit=EUR&unit=...).
    params = [
        ("format", "JSON"),
        ("lang", config.LANG),
        ("sinceTimePeriod", config.START_PERIOD),
    ]
    for dim, values in filter_params.items():
        for value in values:
            params.append((dim, value))

    letzter_fehler = None
    for versuch in range(1, config.MAX_RETRIES + 1):
        try:
            antwort = requests.get(
                url, params=params, timeout=config.REQUEST_TIMEOUT
            )
        except requests.RequestException as exc:
            letzter_fehler = exc
            logger.warning(
                "%s: Abruf fehlgeschlagen (Versuch %d/%d): %s",
                dataset_code, versuch, config.MAX_RETRIES, exc,
            )
        else:
            if antwort.status_code == 200:
                return antwort.json()
            if 400 <= antwort.status_code < 500:
                # Client-Fehler (z. B. ungültiger Filter) -> Wiederholen
                # sinnlos, sofort abbrechen.
                raise RuntimeError(
                    f"{dataset_code}: API-Fehler {antwort.status_code} "
                    f"(kein Retry bei 4xx): {antwort.text[:300]}"
                )
            letzter_fehler = RuntimeError(
                f"HTTP {antwort.status_code}: {antwort.text[:300]}"
            )
            logger.warning(
                "%s: Server-Fehler %d (Versuch %d/%d)",
                dataset_code, antwort.status_code, versuch, config.MAX_RETRIES,
            )
        if versuch < config.MAX_RETRIES:
            wartezeit = config.RETRY_BACKOFF_SECONDS ** versuch
            logger.info("%s: Neuer Versuch in %d s ...", dataset_code, wartezeit)
            time.sleep(wartezeit)
    raise RuntimeError(
        f"{dataset_code}: Abruf nach {config.MAX_RETRIES} Versuchen "
        f"fehlgeschlagen. Letzter Fehler: {letzter_fehler}"
    )


# ---------------------------------------------------------------------------
# JSON-stat 2.0 -> Long-Format-DataFrame
# ---------------------------------------------------------------------------
def _codes_geordnet(category: dict) -> list:
    """Liefert die Kategorie-Codes einer Dimension in Index-Reihenfolge.

    JSON-stat 2.0 erlaubt ``index`` sowohl als Objekt (Code -> Position)
    als auch als Liste (direkte Reihenfolge); beides wird unterstützt.
    """
    index = category.get("index", {})
    if isinstance(index, dict):
        codes = [None] * len(index)
        for code, position in index.items():
            codes[position] = code
        return codes
    return list(index)


def parse_period_to_date(periode: str) -> pd.Timestamp:
    """Wandelt einen Eurostat-Perioden-Code in ein Datum (Periodenbeginn) um.

    Unterstützt jährlich (``2020``), halbjährlich (``2020-S2``),
    vierteljährlich (``2020-Q4``) und monatlich (``2020-01``).
    Nicht erkennbare Formate ergeben ``pd.NaT``.
    """
    s = str(periode)
    treffer = re.fullmatch(r"(\d{4})", s)
    if treffer:
        return pd.Timestamp(int(treffer.group(1)), 1, 1)
    treffer = re.fullmatch(r"(\d{4})-S([12])", s)
    if treffer:
        monat = 1 if treffer.group(2) == "1" else 7
        return pd.Timestamp(int(treffer.group(1)), monat, 1)
    treffer = re.fullmatch(r"(\d{4})-Q([1-4])", s)
    if treffer:
        monat = (int(treffer.group(2)) - 1) * 3 + 1
        return pd.Timestamp(int(treffer.group(1)), monat, 1)
    treffer = re.fullmatch(r"(\d{4})-(\d{2})", s)
    if treffer:
        return pd.Timestamp(int(treffer.group(1)), int(treffer.group(2)), 1)
    return pd.NaT


def parse_jsonstat(payload: dict) -> pd.DataFrame:
    """Konvertiert ein JSON-stat-2.0-Dokument in ein Long-Format-DataFrame.

    Eine Zeile je Kombination aus Dimensionen + Zeit + Wert, damit spätere
    Merges eindeutig funktionieren. Zu jeder Dimensions-Spalte wird eine
    zugehörige ``*_label``-Spalte mit den deutschen Eurostat-Bezeichnungen
    ergänzt (Abruf erfolgt mit ``lang=DE``). Zusätzlich werden die Spalten
    ``time_date`` (Periodenbeginn als Datum) und ``status`` (Eurostat-
    Datenqualitäts-Flags, z. B. 'e' = geschätzt) angehängt.
    """
    dim_ids = list(payload["id"])
    sizes = list(payload["size"])
    codes_pro_dim = []
    labels_pro_dim = []
    for dim in dim_ids:
        category = payload["dimension"][dim]["category"]
        codes_pro_dim.append(_codes_geordnet(category))
        labels_pro_dim.append(category.get("label") or {})

    values = payload.get("value", {})
    if isinstance(values, list):
        wert_items = list(enumerate(values))
    else:  # sparse Darstellung: Position -> Wert
        wert_items = [(int(pos), wert) for pos, wert in values.items()]

    status = payload.get("status", {})
    if isinstance(status, list):
        status_map = {i: s for i, s in enumerate(status) if s}
    else:
        status_map = {int(pos): s for pos, s in status.items()}

    n_dims = len(dim_ids)
    zeilen = []
    for flat_index, wert in wert_items:
        # Gemischtes Stellenwertsystem: letzte Dimension läuft am schnellsten.
        rest = flat_index
        positionen = [0] * n_dims
        for i in range(n_dims - 1, -1, -1):
            positionen[i] = rest % sizes[i]
            rest //= sizes[i]
        zeilen.append(
            [codes_pro_dim[i][positionen[i]] for i in range(n_dims)]
            + [wert, status_map.get(flat_index)]
        )

    df = pd.DataFrame(zeilen, columns=dim_ids + ["value", "status"])
    df["value"] = pd.to_numeric(df["value"], errors="coerce")

    # Deutsche Bezeichnungen je Dimension (Fallback: Code selbst)
    for dim, labels in zip(dim_ids, labels_pro_dim):
        df[f"{dim}_label"] = df[dim].map(labels).fillna(df[dim])

    if "time" in df.columns:
        df["time_date"] = df["time"].map(parse_period_to_date)

    return df.sort_values(dim_ids).reset_index(drop=True)


# ---------------------------------------------------------------------------
# Validierung und Export
# ---------------------------------------------------------------------------
def validate_mandatory_filters(
    df: pd.DataFrame, dataset_code: str, pflichtfilter: dict
) -> None:
    """Prüft vor dem Export, ob alle Pflichtfilter korrekt angewendet wurden.

    Raises:
        ValueError: Wenn eine gefilterte Dimension unerlaubte Codes enthält
            oder Zeiträume vor ``config.START_PERIOD`` auftauchen.
    """
    for dim, erlaubt in pflichtfilter.items():
        if dim not in df.columns:
            raise ValueError(
                f"{dataset_code}: Pflichtfilter-Dimension '{dim}' fehlt im "
                f"DataFrame (Spalten: {list(df.columns)})"
            )
        verstoesse = sorted(set(df[dim].unique()) - set(erlaubt))
        if verstoesse:
            raise ValueError(
                f"{dataset_code}: Pflichtfilter '{dim}' verletzt – "
                f"unerlaubte Codes: {verstoesse} (erlaubt: {erlaubt})"
            )
    if "time" in df.columns and not df.empty:
        jahre = df["time"].astype(str).str[:4].astype(int)
        if jahre.min() < int(config.START_PERIOD):
            raise ValueError(
                f"{dataset_code}: Zeitraum-Filter verletzt – frühestes Jahr "
                f"{jahre.min()} < {config.START_PERIOD}"
            )


def export_excel(df: pd.DataFrame, pfad: Path, sheet_name: str = "Daten") -> None:
    """Schreibt ein DataFrame als formatierte Excel-Datei.

    Mit fixierter Kopfzeile, Autofilter und angepassten Spaltenbreiten,
    damit die Datei auch direkt in Excel gut nutzbar ist.
    """
    with pd.ExcelWriter(pfad, engine="openpyxl") as writer:
        df.to_excel(writer, sheet_name=sheet_name, index=False)
        blatt = writer.sheets[sheet_name]
        blatt.freeze_panes = "A2"
        blatt.auto_filter.ref = blatt.dimensions
        for spalte_idx, spalte in enumerate(df.columns, start=1):
            laenge = max(
                len(str(spalte)),
                *(len(str(v)) for v in df[spalte].head(500)),
            )
            blatt.column_dimensions[
                blatt.cell(row=1, column=spalte_idx).column_letter
            ].width = min(laenge + 2, 50)


# ---------------------------------------------------------------------------
# Orchestrierung je Datensatz
# ---------------------------------------------------------------------------
def fetch_dataset(dataset_code: str, spez: dict) -> pd.DataFrame:
    """Ruft einen Datensatz ab, validiert ihn und exportiert die Rohdaten.

    Loggt Abrufzeitpunkt, Zeilenzahl und verwendete Filterparameter und
    pflegt diese Informationen in ``fetch_metadata.json`` (Grundlage für
    die Datenstand-Anzeige im Dashboard).
    """
    abruf_zeitpunkt = datetime.now()
    logger.info(
        "Starte Abruf: %s (%s) | Filter: %s | ab %s",
        dataset_code, spez["beschreibung"], spez["filter"] or "(alle Dimensionen)",
        config.START_PERIOD,
    )
    payload = fetch_jsonstat(dataset_code, spez["filter"])
    df = parse_jsonstat(payload)
    validate_mandatory_filters(df, dataset_code, spez["pflichtfilter"])

    pfad = config.OUTPUT_DIR / spez["datei"]
    export_excel(df, pfad)
    logger.info(
        "%s: %d Zeilen exportiert -> %s (Abruf: %s)",
        dataset_code, len(df), pfad,
        abruf_zeitpunkt.strftime("%Y-%m-%d %H:%M:%S"),
    )

    _aktualisiere_metadata(
        dataset_code,
        {
            "beschreibung": spez["beschreibung"],
            "abruf_zeitpunkt": abruf_zeitpunkt.isoformat(timespec="seconds"),
            "zeilen": int(len(df)),
            "filter": spez["filter"],
            "start_period": config.START_PERIOD,
            "datei": spez["datei"],
        },
    )
    return df


def _aktualisiere_metadata(dataset_code: str, eintrag: dict) -> None:
    """Schreibt/aktualisiert die Abruf-Metadaten (JSON) im Output-Ordner."""
    meta = {}
    if config.FETCH_METADATA_FILE.exists():
        try:
            meta = json.loads(
                config.FETCH_METADATA_FILE.read_text(encoding="utf-8")
            )
        except (json.JSONDecodeError, OSError):
            meta = {}
    meta[dataset_code] = eintrag
    config.FETCH_METADATA_FILE.write_text(
        json.dumps(meta, ensure_ascii=False, indent=2), encoding="utf-8"
    )


def fetch_lc_lci_lev() -> pd.DataFrame:
    """Datensatz lc_lci_lev – Arbeitskosten (Niveau)."""
    return fetch_dataset("lc_lci_lev", config.DATASETS["lc_lci_lev"])


def fetch_nrg_pc_205() -> pd.DataFrame:
    """Datensatz nrg_pc_205 – Strompreise Nicht-Haushalte."""
    return fetch_dataset("nrg_pc_205", config.DATASETS["nrg_pc_205"])


def fetch_nrg_pc_203() -> pd.DataFrame:
    """Datensatz nrg_pc_203 – Gaspreise Nicht-Haushalte."""
    return fetch_dataset("nrg_pc_203", config.DATASETS["nrg_pc_203"])


def fetch_sts_inpr_m() -> pd.DataFrame:
    """Datensatz sts_inpr_m – Industrieproduktionsindex."""
    return fetch_dataset("sts_inpr_m", config.DATASETS["sts_inpr_m"])


def fetch_sts_inppd_m() -> pd.DataFrame:
    """Datensatz sts_inppd_m – Erzeugerpreisindex Industrie."""
    return fetch_dataset("sts_inppd_m", config.DATASETS["sts_inppd_m"])


def main() -> None:
    """Führt den Abruf aller fünf Datensätze nacheinander aus."""
    config.OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    abrufe = [
        fetch_lc_lci_lev,
        fetch_nrg_pc_205,
        fetch_nrg_pc_203,
        fetch_sts_inpr_m,
        fetch_sts_inppd_m,
    ]
    fehler = []
    for abruf in abrufe:
        try:
            abruf()
        except Exception as exc:  # weitermachen, Rest trotzdem abrufen
            fehler.append(abruf.__name__)
            logger.error("%s fehlgeschlagen: %s", abruf.__name__, exc)
    if fehler:
        logger.error("Abgeschlossen mit Fehlern bei: %s", ", ".join(fehler))
        raise SystemExit(1)
    logger.info("Alle %d Datensätze erfolgreich abgerufen.", len(abrufe))


if __name__ == "__main__":
    main()
