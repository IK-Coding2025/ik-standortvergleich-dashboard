"""Zentrale Pfad- und Parameterkonfiguration für die Standortvergleich-Pipeline.

Alle Module (fetch_data.py, merge_data.py, dashboard.py) beziehen ihre
Pfade und Parameter ausschließlich von hier, damit Anpassungen nur an
einer Stelle nötig sind.

Ausführungsreihenfolge der Pipeline:
    1. python fetch_data.py      -> 6 Rohdaten-Excel-Dateien
    2. python merge_data.py      -> 2 Merge-Dateien + lc_lci_lev_final.xlsx
                                  + env_waspac_final.xlsx
    3. streamlit run dashboard.py
"""

import logging
import sys
from pathlib import Path

# Keine Bytecode-Caches (.pyc) schreiben: Der Ordner liegt auf einem
# FileCloud-Team-Laufwerk; staendig neu geschriebene Cache-Dateien
# erzeugen dort sonst Sync-Konflikte. Quell- und Datendateien werden
# davon nicht beruehrt und bleiben stets aktuell synchronisiert.
sys.dont_write_bytecode = True

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
# Eurostat Comext API (DS-045409 – Außenhandel nach HS/KN, SDMX 2.1)
# ---------------------------------------------------------------------------
# Comext-Datensätze (Präfix "DS-") liegen auf einem eigenen API-Endpunkt und
# werden per SDMX-2.1-Schlüssel (statt Query-Parametern) abgefragt:
#   {COMEXT_BASE_URL}/{dataset}/{freq}.{reporter}.{partner}.{product}.{flow}.{indicators}
# Die Antwort ist im selben JSON-stat-2.0-Format wie die reguläre
# Dissemination-API, daher kann ``parse_jsonstat`` unverändert wiederverwendet
# werden.
COMEXT_BASE_URL = "https://ec.europa.eu/eurostat/api/comext/dissemination/sdmx/2.1/data"

# EU-27 (EUROSTAT-Codes) und deutsche NUTS-1-Regionen für demo_r_gind3
EU27_CODES = [
    "BE", "BG", "CZ", "DK", "DE", "EE", "IE", "EL", "ES", "FR",
    "HR", "IT", "CY", "LV", "LT", "LU", "HU", "MT", "NL", "AT",
    "PL", "PT", "RO", "SI", "SK", "FI", "SE",
]
DE_NUTS1_CODES = [
    "DE1", "DE2", "DE3", "DE4", "DE5", "DE6", "DE7", "DE8",
    "DE9", "DEA", "DEB", "DEC", "DED", "DEE", "DEF", "DEG",
]

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
            "nace_r2": ["C2221", "C2222", "C22", "C2223", "C2229",
                        "C2896", "C2016", "C1721", "C10_C11"],
            # Kunststofferzeugnisse (Details), Gummi/Kunststoff Division C22
            # (volle Länderabdeckung), Kunststoffmaschinen C2896,
            # Kunststoffe in Primärformen C2016, Papierverpackungen C1721,
            # Nahrungs-/Futtermittel + Getränke C10_C11
            "unit": ["I21", "PCH_PRE"],
        },
        "pflichtfilter": {
            "nace_r2": ["C2221", "C2222", "C22", "C2223", "C2229",
                        "C2896", "C2016", "C1721", "C10_C11"],
            "unit": ["I21", "PCH_PRE"],
        },
        "datei": "sts_inpr_m_raw.xlsx",
    },
    "sts_inppd_m": {
        "beschreibung": "Erzeugerpreisindex Industrie, Inlandsmarkt (monatlich)",
        "filter": {
            "nace_r2": ["C2221", "C2222", "C22", "C2223", "C2229",
                        "C2896", "C2016", "C1721", "C10_C11"],
            "unit": ["I21", "PCH_PRE"],
        },
        "pflichtfilter": {
            "nace_r2": ["C2221", "C2222", "C22", "C2223", "C2229",
                        "C2896", "C2016", "C1721", "C10_C11"],
            "unit": ["I21", "PCH_PRE"],
        },
        "datei": "sts_inppd_m_raw.xlsx",
    },
    "env_waspac": {
        "beschreibung": "Verpackungsabfälle und Recyclingquoten",
        "filter": {
            "waste": [
                "W1501", "W150101", "W150102", "W150103", "W150104",
                "W15010401", "W15010402", "W150107", "W150199",
            ],
            "wst_oper": [
                "GEN", "RCV_OTH", "RCV_E_PAC", "RCY", "RCY_NAT",
                "RCY_EU_FOR", "RCY_NEU",
            ],
            "unit": ["KG_HAB", "T", "PC"],
        },
        "pflichtfilter": {
            "waste": [
                "W1501", "W150101", "W150102", "W150103", "W150104",
                "W15010401", "W15010402", "W150107", "W150199",
            ],
            "wst_oper": [
                "GEN", "RCV_OTH", "RCV_E_PAC", "RCY", "RCY_NAT",
                "RCY_EU_FOR", "RCY_NEU",
            ],
            "unit": ["KG_HAB", "T", "PC"],
        },
        "start_period": "2013",
        "datei": "env_waspac_raw.xlsx",
    },
    "demo_r_gind3": {
        "beschreibung": "Bevölkerung - regionale Daten (1. Januar)",
        "filter": {
            "indic_de": ["JAN"],
            "geo": EU27_CODES + DE_NUTS1_CODES,
        },
        "pflichtfilter": {
            "indic_de": ["JAN"],
            "geo": EU27_CODES + DE_NUTS1_CODES,
        },
        "start_period": "2000",
        "datei": "demo_r_gind3_raw.xlsx",
    },
    "demo_gind": {
        "beschreibung": "Bevölkerung - EU-27 Aggregat (1. Januar)",
        "filter": {
            "geo": ["EU27_2020"],
            "indic_de": ["JAN"],
        },
        "pflichtfilter": {
            "geo": ["EU27_2020"],
            "indic_de": ["JAN"],
        },
        "start_period": "2000",
        "datei": "demo_gind_raw.xlsx",
    },
    "proj_25ndbi": {
        "beschreibung": "Bevölkerungsprojektionen 2025-2100",
        "filter": {
            "geo": EU27_CODES + ["EU27_2020"],
            "indic_de": ["JAN", "PC_Y15_64"],
            "projection": [
                "BSL", "LFRT", "LMRT", "HMIGR", "LMIGR", "NMIGR", "DCONV",
            ],
        },
        "pflichtfilter": {
            "indic_de": ["JAN", "PC_Y15_64"],
            "projection": [
                "BSL", "LFRT", "LMRT", "HMIGR", "LMIGR", "NMIGR", "DCONV",
            ],
        },
        "start_period": "2025",
        "datei": "proj_25ndbi_raw.xlsx",
    },
    "lfsa_egan22d": {
        "beschreibung": "Erwerbstätige nach NACE, Alter und Geschlecht",
        "filter": {
            "nace_r2": ["C22", "TOTAL"],
            "age": ["Y15-64"],
            "sex": ["T"],
        },
        "pflichtfilter": {
            "nace_r2": ["C22", "TOTAL"],
            "age": ["Y15-64"],
            "sex": ["T"],
        },
        "start_period": "2015",
        "datei": "lfsa_egan22d_raw.xlsx",
    },
    "jvs_q_r21": {
        "beschreibung": "Quote der offenen Stellen (Job Vacancy Rate)",
        "filter": {
            "nace_r2_1": ["C", "A-T"],
            "indic_em": ["JVR"],
            "sizeclas": ["TOTAL"],
            "s_adj": ["NSA"],
        },
        "pflichtfilter": {
            "nace_r2_1": ["C", "A-T"],
            "indic_em": ["JVR"],
            "sizeclas": ["TOTAL"],
            "s_adj": ["NSA"],
        },
        "start_period": "2016",
        "datei": "jvs_q_r21_raw.xlsx",
    },
}

# ---------------------------------------------------------------------------
# Außenhandel (DS-045409): Produktgruppen, NACE-Zuordnung, Datensatzdefinition
# ---------------------------------------------------------------------------
# Bezeichnung (Produktgruppe) -> zugehörige offizielle Warennummern
# (KN8-/HS6-/HS4-Codes). Die Bezeichnung wird im Dashboard um die Codes in
# Klammern ergänzt, z. B. "Verpackung / Folien aus LDPE (39201023, 39201024,
# 39201025)".
PRODUKTGRUPPEN = {
    "Verpackung / Folien aus LDPE": ["39201023", "39201024", "39201025"],
    "Verpackung / Folien aus HDPE": ["39201028"],
    "Verpackung / Folien aus Ethylencopolymere/Sonstige": [
        "39201040", "39201081", "39201089",
    ],
    "Verpackung / Folien aus PP": ["39202021", "39202029", "39202080"],
    "Verpackung / Folien aus PS": ["39203000"],
    "Verpackung / Folien aus PET + Polyester": [
        "39206212", "39206219", "39206900",
    ],
    "Verpackung / Folien aus Regen. Cellulose": ["39207100"],
    "Verpackung / Folien aus Celluloseacetat": ["39207380"],
    "Verpackung / Folien aus PA": ["39209200"],
    "Verpackung / Folien aus PC": ["39206100"],
    "Tüten, Beutel, Säcke usw. aus PE": ["39232100"],
    "Tüten, Beutel, Säcke usw. aus PVC": ["39232910"],
    "Tüten, Beutel, Säcke usw. aus andere Kunststoffe": ["39232990"],
    "Flaschen kleiner 2 l": ["39233010"],
    "Flaschen größer 2 l": ["39233090"],
    "andere Transportverpackungen": ["39239000"],
    "Verschlüsse": ["39235010", "39235090"],
    "Becher, Dosen, Kisten": ["39231010", "39231090"],
    "Polymere des Ethylens in Primärformen (PE, z. B. LDPE, LLDPE, HDPE)": [
        "3901",
    ],
    "Polymere des Propylens oder anderer Olefine in Primärformen (PP)": [
        "3902",
    ],
    "Polymere des Styrols in Primärformen (PS, EPS)": ["3903"],
    "Polyacetale, andere Polyether und Epoxidharze in Primärformen; "
    "Polycarbonate, Alkydharze, Polyallylester und andere Polyester in "
    "Primärformen": ["3907"],
    "Erdöl und Öle aus bituminösen Mineralien, roh": ["2709"],
    "Naphtha (Rohbenzin), Motorenbenzin & Flugbenzin, Spezialbenzine - "
    "Leichte Öle und Zubereitungen aus Erdöl oder bituminösen Mineralien, "
    "bei denen >= 90 % des Volumens „einschließlich Verluste“ bei 210 °C "
    "„nach ASTM D 86“ destillieren (ausgenommen solche, die Biodiesel "
    "enthalten)": ["271012"],
}

# Umgekehrte Zuordnung Warennummer -> Produktgruppen-Bezeichnung
WARENNUMMER_ZU_PRODUKTGRUPPE = {
    code: bezeichnung
    for bezeichnung, codes in PRODUKTGRUPPEN.items()
    for code in codes
}

# Alle abzurufenden Warennummern (Reihenfolge wie oben definiert)
AUSSENHANDEL_PRODUKTE = [
    code for codes in PRODUKTGRUPPEN.values() for code in codes
]

# NACE-Zuordnung je Warennummer-Präfix (erste Übereinstimmung gewinnt)
WZ_PRAEFIXE = [
    (("3923",), "2222"),
    (("3916", "3917", "3920", "3921"), "2221"),
    (("3901", "3902", "3903", "3907"), "2016"),
    (("2709",), "0610"),
    (("271012",), "1920"),
]

# Erklärung der NACE-Codes (WZ 2008)
WZ_LABELS = {
    "2222": "Herstellung von Verpackungsmitteln aus Kunststoffen",
    "2221": "Herstellung von Platten, Folien, Schläuchen und Profilen aus "
            "Kunststoffen",
    "2016": "Herstellung von Kunststoffen in Primärformen",
    "0610": "Gewinnung von Erdöl",
    "1920": "Mineralölverarbeitung",
}


def warennummer_zu_wz(warennummer: str) -> str | None:
    """Ordnet eine Warennummer anhand ihres Präfixes einem NACE-Code (WZ) zu.

    Gibt ``None`` zurück, wenn keine Zuordnungsregel greift.
    """
    code = str(warennummer)
    for praefixe, wz in WZ_PRAEFIXE:
        if code.startswith(praefixe):
            return wz
    return None


DATASETS["ds_045409"] = {
    "beschreibung": (
        "Außenhandel – Verpackungen und Kunststoff-Vorprodukte (Menge in "
        "100 kg), Eurostat Comext DS-045409"
    ),
    "dataset_code": "ds-045409",
    "freq": ["A"],                       # nur jährliche Werte
    "reporter": ["DE", "EU27_2020"],
    "partner": [],                       # leer = alle Partnerländer abrufen
    "product": AUSSENHANDEL_PRODUKTE,
    "flow": ["1", "2"],                  # 1 = Einfuhr/Import, 2 = Ausfuhr/Export
    "indicators": ["QUANTITY_IN_100KG"],
    "start_period": "2015",
    "datei": "ds_045409_raw.xlsx",
}

# Dateinamen der finalen (gemergten) Tabellen
FINAL_LC = "lc_lci_lev_final.xlsx"
FINAL_MERGE_INDUSTRIE = "merge_industrieproduktion_erzeugerpreise.xlsx"
FINAL_MERGE_ENERGIE = "merge_energiepreise.xlsx"
FINAL_WASTE = "env_waspac_final.xlsx"
FINAL_BEV = "merge_bevoelkerung.xlsx"
FINAL_PROJ = "proj_25ndbi_final.xlsx"
FINAL_LFSA = "lfsa_egan22d_final.xlsx"
FINAL_JVS = "jvs_q_r21_final.xlsx"
FINAL_AUSSENHANDEL = "aussenhandel_final.xlsx"

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
