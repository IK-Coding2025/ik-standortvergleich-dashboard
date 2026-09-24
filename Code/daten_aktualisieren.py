"""Aktualisiert sämtliche Eurostat-Daten des Standortvergleich-Dashboards.

Fragt alle in ``config.DATASETS`` definierten Datensätze (inkl. des
Außenhandels-Datensatzes ``ds_045409``, Eurostat Comext) frisch bei Eurostat
ab, führt anschließend den Merge/die Aufbereitung aus und schreibt alle
finalen Excel-Tabellen neu in den Output-Pfad. Nutzt dieselbe Logik wie
``run_pipeline.py`` (fetch_data.main() + merge_data.main()), ist aber als
eigenständiges, klar benanntes Skript für die regelmäßige Datenaktualisierung
gedacht (z. B. für eine manuelle oder geplante Ausführung durch das
Fachreferat, unabhängig von Entwicklungs-/Testzwecken).

Nach dem Lauf einfach das Dashboard neu laden bzw. neu starten, damit die
aktualisierten Daten angezeigt werden:
    streamlit run dashboard.py

Ausführen:
    python daten_aktualisieren.py
"""

import sys

# Keine Bytecode-Caches schreiben (FileCloud-Sync-Konflikte), siehe config.py
sys.dont_write_bytecode = True

import json
from datetime import datetime

import config
import fetch_data
import merge_data

logger = config.setup_logging()


def _datenstand_zusammenfassung() -> str:
    """Liest die Fetch-Metadaten und erstellt eine kurze Übersicht je Datensatz."""
    if not config.FETCH_METADATA_FILE.exists():
        return "Keine Metadaten gefunden."
    try:
        meta = json.loads(config.FETCH_METADATA_FILE.read_text(encoding="utf-8"))
    except (json.JSONDecodeError, OSError) as exc:
        return f"Metadaten nicht lesbar: {exc}"
    zeilen = []
    for schluessel, eintrag in sorted(meta.items()):
        if schluessel.startswith("_") or "abruf_zeitpunkt" not in eintrag:
            continue
        zeilen.append(
            f"  - {schluessel}: {eintrag.get('zeilen', '?')} Zeilen "
            f"(Abruf: {eintrag['abruf_zeitpunkt']})"
        )
    return "\n".join(zeilen) if zeilen else "Keine Datensätze in den Metadaten."


def main() -> None:
    start = datetime.now()
    logger.info(
        "=== Aktualisierung aller Eurostat-Daten gestartet (%s Datensätze "
        "inkl. Außenhandel ds_045409) ===",
        len(config.DATASETS),
    )
    try:
        fetch_data.main()
        merge_data.main()
    except SystemExit:
        logger.error(
            "Aktualisierung mit Fehlern beendet – siehe Log oben für Details "
            "zu den betroffenen Datensätzen. Bereits erfolgreich "
            "abgerufene/finalisierte Dateien wurden trotzdem aktualisiert."
        )
        raise
    dauer = (datetime.now() - start).total_seconds()
    logger.info("=== Aktualisierung abgeschlossen (%.0f s) ===", dauer)
    logger.info("Datenstand je Datensatz:\n%s", _datenstand_zusammenfassung())
    print(
        "\nFertig. Bitte das Dashboard neu laden bzw. neu starten, damit "
        "die aktualisierten Daten angezeigt werden:\n"
        "    streamlit run dashboard.py"
    )


if __name__ == "__main__":
    main()
