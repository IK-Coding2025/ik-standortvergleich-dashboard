"""Zentrale Pfad- und Parameterkonfiguration für die Standortvergleich-Pipeline.

Alle Module (fetch_data.py, merge_data.py, dashboard.py) beziehen ihre
Pfade und Parameter ausschließlich von hier, damit Anpassungen nur an
einer Stelle nötig sind.

Ausführungsreihenfolge der Pipeline:
    1. python fetch_data.py      -> 5 Rohdaten-Excel-Dateien
    2. python merge_data.py      -> 2 Merge-Dateien + lc_lci_lev_final.xlsx
    3. streamlit run dashboard.py
"""

import logging
import sys
from pathlib import Path

# ---------------------------------------------------------------------------
# Ablagepfade (zwingend vorgegeben)
# ---------------------------------------------------------------------------
# Lokal wird das FileCloud-Laufwerk verwendet. Beim Cloud-Deployment
# (Streamlit Community Cloud) existiert dieser Pfad nicht; dann wird relativ
# zum Repository gearbeitet – config.py liegt in <Repo>/Code/, die Ordner
# "Output Excel" und "Input_Logo" liegen daneben in <Repo>/.
_PFAD_LOKAL = Path(
    r"F:\Team Folders\IK_Server\Wirtschaft\Coding\Standortvergleich"
)
BASE_DIR = (
    _PFAD_LOKAL
    if _PFAD_LOKAL.exists()
    else Path(__file__).resolve().parent.parent
)
CODE_DIR = BASE_DIR / "Code"
OUTPUT_DIR = BASE_DIR / "Output Excel"
LOGO_DIR = BASE_DIR / "Input_Logo"

LOG_FILE = OUTPUT_DIR / "pipeline.log"
FETCH_METADATA_FILE = OUTPUT_DIR / "fetch_metadata.json"

# ---------------------------------------------------------------------------
# Eurostat Dissemination API (JSON-stat 2.0)
# ---------------------------------------------------------------------------
EUROSTAT_BASE_URL = (
    "https://ec.europa.eu/eurostat/api/dissemination/statistics/1.0/data"
)
LANG = "DE"                 # Deutsche Eurostat-Bezeichnungen (Kategorielabels)
START_PERIOD = "2020"       # Zeitreihen ab 2020
MAX_RETRIES = 3             # Max. Versuche bei API-Timeout/-Fehler
RETRY_BACKOFF_SECONDS = 2   # Basis für exponentielles Backoff
REQUEST_TIMEOUT = 120       # Sekunden

# ---------------------------------------------------------------------------
# Datensatz-Definitionen
# ---------------------------------------------------------------------------
# "filter":       Werden als Query-Parameter an die API übergeben.
# "pflichtfilter": Werden nach dem Abruf validiert (Export bricht bei
#                 Verstoß ab, damit keine falsch gefilterten Daten
#                 weiterverarbeitet werden).
DATASETS = {
    "lc_lci_lev": {
        "beschreibung": "Arbeitskosten (Niveau) – Labour Cost Index",
        "filter": {
            "nace_r2": ["C"],                    # Verarbeitendes Gewerbe
            "unit": ["EUR", "RT_PRE_EUR"],
        },
        "pflichtfilter": {
            "nace_r2": ["C"],
            "unit": ["EUR", "RT_PRE_EUR"],
        },
        "datei": "lc_lci_lev_raw.xlsx",
    },
    "nrg_pc_205": {
        "beschreibung": "Strompreise Nicht-Haushalte",
        "filter": {},                            # alle Dimensionen vollständig
        "pflichtfilter": {},
        "datei": "nrg_pc_205_raw.xlsx",
    },
    "nrg_pc_203": {
        "beschreibung": "Gaspreise Nicht-Haushalte",
        "filter": {},
        "pflichtfilter": {},
        "datei": "nrg_pc_203_raw.xlsx",
    },
    "sts_inpr_m": {
        "beschreibung": "Industrieproduktionsindex (monatlich)",
        "filter": {
            "nace_r2": ["C2221", "C2222"],       # Kunststofferzeugnisse
            "unit": ["I21", "PCH_PRE"],
        },
        "pflichtfilter": {
            "nace_r2": ["C2221", "C2222"],
            "unit": ["I21", "PCH_PRE"],
        },
        "datei": "sts_inpr_m_raw.xlsx",
    },
    "sts_inppd_m": {
        "beschreibung": "Erzeugerpreisindex Industrie, Inlandsmarkt (monatlich)",
        "filter": {
            "nace_r2": ["C2221", "C2222"],
            "unit": ["I21", "PCH_PRE"],
        },
        "pflichtfilter": {
            "nace_r2": ["C2221", "C2222"],
            "unit": ["I21", "PCH_PRE"],
        },
        "datei": "sts_inppd_m_raw.xlsx",
    },
}

# Dateinamen der finalen (gemergten) Tabellen
FINAL_LC = "lc_lci_lev_final.xlsx"
FINAL_MERGE_INDUSTRIE = "merge_industrieproduktion_erzeugerpreise.xlsx"
FINAL_MERGE_ENERGIE = "merge_energiepreise.xlsx"

# Schlüsselspalten der Merges (Dokumentation der Join-Logik in merge_data.py)
MERGE_A_KEYS = ["freq", "time", "geo", "nace_r2", "s_adj", "unit"]
MERGE_B_KEYS = ["freq", "time", "geo", "nrg_cons", "unit", "tax", "currency"]


def setup_logging() -> logging.Logger:
    """Richtet ein gemeinsames Logging für alle Pipeline-Module ein.

    Loggt auf die Konsole und in die Datei ``pipeline.log`` im Output-Ordner.
    Mehrfachaufruf ist ungefährlich (Handler werden nur einmal angelegt).
    """
    logger = logging.getLogger("standortvergleich")
    if logger.handlers:
        return logger
    logger.setLevel(logging.INFO)
    formatter = logging.Formatter(
        "%(asctime)s [%(levelname)s] %(message)s", datefmt="%Y-%m-%d %H:%M:%S"
    )
    # Windows-Konsole auf UTF-8 umstellen, damit deutsche Umlaute nicht
    # zu Kodierungsfehlern führen.
    for stream in (sys.stdout, sys.stderr):
        try:
            stream.reconfigure(encoding="utf-8", errors="replace")
        except Exception:
            pass
    console = logging.StreamHandler()
    console.setFormatter(formatter)
    logger.addHandler(console)
    try:
        OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
        file_handler = logging.FileHandler(LOG_FILE, encoding="utf-8")
        file_handler.setFormatter(formatter)
        logger.addHandler(file_handler)
    except OSError as exc:  # Logdatei ist optional, Konsole reicht im Zweifel
        logger.warning("Logdatei konnte nicht angelegt werden: %s", exc)
    return logger
