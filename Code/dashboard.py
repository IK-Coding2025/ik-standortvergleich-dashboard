"""IK Standortvergleich Deutschland – Europa (Streamlit-Dashboard).

Liest die finalen Excel-Tabellen (Arbeitskosten, Energiepreise) sowie die
STS-Rohdaten (Industrieproduktion, Erzeugerpreise) aus dem Output-Pfad und
visualisiert sie in drei Tabs. Layout orientiert an
https://ikdashboard.streamlit.app/ (Kopfzeile mit Logo, Seitenleiste für
Filter, KPI-Kacheln, Charts, Footer).

Starten:
    streamlit run dashboard.py
"""

import sys

# Keine Bytecode-Caches schreiben (FileCloud-Sync-Konflikte), siehe config.py
sys.dont_write_bytecode = True

import base64
import json
import re

import pandas as pd
import plotly.graph_objects as go
import streamlit as st

import config

# IK-Unternehmensfarbe (aus dem Logo) + ergänzende Farbpalette
IK_BLAU = "#2E3192"
FARBPALETTE = [
    "#2E3192", "#E30613", "#009FE3", "#93C01F", "#F39200",
    "#662483", "#009036", "#8C8C8C", "#DA9A00", "#00B2A9",
]

st.set_page_config(
    page_title="IK Dashboard zum Standortvergleich Deutschland – Europa",
    layout="wide",
)

# Tab-Leiste (Themen-Auswahlmenü) größer und fett darstellen; der aktive
# Tab wird in IK-Blau hervorgehoben. Selektoren nutzen die stabilen
# ARIA-Rollen (tablist/tab) plus die data-testids älterer und neuerer
# Streamlit-Versionen; die Unterelemente erben Schriftgröße und -gewicht.
st.markdown(
    f"""
    <style>
    [role="tablist"] button,
    [role="tablist"] button p,
    [role="tablist"] button span,
    [role="tablist"] button div,
    [data-testid="stTabs"] button,
    [data-testid="stTabs"] button p,
    [data-testid="stTabs"] button span,
    [data-testid="stTabs"] button div,
    button[data-baseweb="tab"],
    button[data-baseweb="tab"] p,
    button[data-testid="stTab"],
    button[data-testid="stTab"] p,
    [role="tab"],
    [role="tab"] p,
    [role="tab"] span {{
        font-size: 1.25rem !important;
        font-weight: 700 !important;
    }}
    [role="tab"][aria-selected="true"],
    [role="tab"][aria-selected="true"] p,
    [role="tab"][aria-selected="true"] span,
    button[data-baseweb="tab"][aria-selected="true"],
    button[data-baseweb="tab"][aria-selected="true"] p,
    button[data-testid="stTab"][aria-selected="true"] {{
        color: {IK_BLAU} !important;
    }}
    </style>
    """,
    unsafe_allow_html=True,
)


# ---------------------------------------------------------------------------
# Hilfsfunktionen
# ---------------------------------------------------------------------------
def fmt_de(wert, nachkommastellen: int = 2) -> str:
    """Formatiert eine Zahl mit deutschen Dezimal-/Tausendertrennzeichen."""
    if wert is None or pd.isna(wert):
        return "–"
    text = f"{wert:,.{nachkommastellen}f}"
    return text.replace(",", "#").replace(".", ",").replace("#", ".")


def finde_logo() -> str | None:
    """Sucht das IK-Logo im Input-Ordner (Dateiname flexibel per Glob)."""
    for muster in ("*.png", "*.jpg", "*.jpeg"):
        treffer = sorted(config.LOGO_DIR.glob(muster))
        if treffer:
            return str(treffer[0])
    return None


def datenstand() -> str:
    """Ermittelt den letzten Abrufzeitpunkt aus den Fetch-Metadaten."""
    if not config.FETCH_METADATA_FILE.exists():
        return "unbekannt (keine Metadaten)"
    try:
        meta = json.loads(config.FETCH_METADATA_FILE.read_text(encoding="utf-8"))
    except (json.JSONDecodeError, OSError):
        return "unbekannt (Metadaten nicht lesbar)"
    zeitpunkte = [
        eintrag["abruf_zeitpunkt"]
        for schluessel, eintrag in meta.items()
        if not schluessel.startswith("_") and "abruf_zeitpunkt" in eintrag
    ]
    if not zeitpunkte:
        return "unbekannt"
    letzter = max(pd.Timestamp(z) for z in zeitpunkte)
    return letzter.strftime("%d.%m.%Y, %H:%M Uhr")


@st.cache_data
def lade_tabelle(datei: str, aenderungszeit: float) -> pd.DataFrame:
    """Liest eine finale Excel-Tabelle ein (gecacht für schnelle Ladezeiten).

    ``aenderungszeit`` (Datei-mtime) ist Teil des Cache-Schlüssels, damit
    nach einem erneuten Datenabruf/Merge automatisch die aktualisierten
    Dateien eingelesen werden statt veralteter Cache-Daten.
    """
    pfad = config.OUTPUT_DIR / datei
    if not pfad.exists():
        raise FileNotFoundError(str(pfad))
    df = pd.read_excel(pfad)
    df["time_date"] = pd.to_datetime(df["time_date"], errors="coerce")
    return df


def periode_vorjahr(periode: str) -> str | None:
    """Liefert den Perioden-Code des Vorjahres (z. B. '2021-S2' -> '2020-S2')."""
    treffer = re.fullmatch(r"(\d{4})(.*)", str(periode))
    if not treffer:
        return None
    return f"{int(treffer.group(1)) - 1}{treffer.group(2)}"


_MONATE = [
    "Januar", "Februar", "März", "April", "Mai", "Juni", "Juli",
    "August", "September", "Oktober", "November", "Dezember",
]


def periode_lang(periode: str) -> str:
    """Formatiert einen Eurostat-Perioden-Code als deutschen Langtext.

    Beispiele: ``2025-S2`` -> '2. Halbjahr 2025', ``2026-06`` ->
    'Juni 2026', ``2025`` -> '2025'.
    """
    s = str(periode)
    treffer = re.fullmatch(r"(\d{4})-S([12])", s)
    if treffer:
        return f"{treffer.group(2)}. Halbjahr {treffer.group(1)}"
    treffer = re.fullmatch(r"(\d{4})-(\d{2})", s)
    if treffer:
        return f"{_MONATE[int(treffer.group(2)) - 1]} {treffer.group(1)}"
    return s


def letzter_wert_mit_vorjahr(
    df: pd.DataFrame, wert_spalte: str
) -> tuple[float, float | None, str, str | None]:
    """Ermittelt aktuellsten Wert und Veränderung gegenüber dem Vorjahr.

    Verglichen wird bevorzugt mit der gleichen Periode des Vorjahres
    (z. B. 2025-S2 ggü. 2024-S2); falls diese fehlt, mit der jeweils
    vorherigen verfügbaren Periode.

    Returns:
        (aktuellster_wert, differenz_vorjahr, periode, vergleichsperiode)
    """
    reihe = (
        df.dropna(subset=[wert_spalte])
        .sort_values("time_date")
        .loc[lambda d: d["time_date"].notna()]
    )
    if reihe.empty:
        return None, None, None, None
    letzte = reihe.iloc[-1]
    wert_aktuell = letzte[wert_spalte]
    periode_aktuell = letzte["time"]

    vergleich = reihe[reihe["time"] == periode_vorjahr(periode_aktuell)]
    if vergleich.empty and len(reihe) > 1:
        vergleich = reihe.iloc[[-2]]  # Fallback: vorherige verfügbare Periode
    if vergleich.empty:
        return wert_aktuell, None, periode_aktuell, None
    wert_vorher = vergleich.iloc[-1][wert_spalte]
    return (
        wert_aktuell,
        wert_aktuell - wert_vorher,
        periode_aktuell,
        vergleich.iloc[-1]["time"],
    )


def kpi_kacheln(
    df: pd.DataFrame,
    wert_spalte: str,
    geos: list,
    einheit_text: str,
    nachkommastellen: int = 2,
    delta_farbe: str = "inverse",
    delta_einheit: str = "",
) -> None:
    """Zeigt je ausgewähltem Land eine KPI-Kachel (aktuellster Wert + Vorjahr).

    ``geos`` enthält die deutschen Länder-Bezeichnungen (geo_label).
    ``delta_farbe='inverse'`` (rot bei Anstieg) für Kostengrößen,
    ``'normal'`` (grün bei Anstieg) für Produktionsgrößen.
    ``delta_einheit`` ergänzt die Einheit im Delta-Text (z. B. ' EUR').
    """
    geos_mit_daten = [g for g in geos if g in set(df["geo_label"])]
    if not geos_mit_daten:
        st.info("Für die gewählte Filterkombination liegen keine Werte vor.")
        return
    spalten = st.columns(min(len(geos_mit_daten), 4))
    for spalte, geo_bez in zip(spalten, geos_mit_daten[:4]):
        teil = df[df["geo_label"] == geo_bez]
        wert, diff, periode, periode_vgl = letzter_wert_mit_vorjahr(
            teil, wert_spalte
        )
        if wert is None:
            continue
        delta_text = (
            f"{fmt_de(diff, nachkommastellen)}{delta_einheit} "
            f"ggü. {periode_vgl}"
            if diff is not None
            else None
        )
        spalte.metric(
            label=f"{geo_bez} ({periode})",
            value=f"{fmt_de(wert, nachkommastellen)} {einheit_text}",
            delta=delta_text,
            delta_color=delta_farbe,
        )
    if len(geos_mit_daten) > 4:
        st.caption(
            f"KPI-Kacheln für die ersten 4 von {len(geos_mit_daten)} "
            "ausgewählten Ländern/Regionen."
        )


def _basis_layout(
    fig: go.Figure, titel: str, legende_unten: bool = False
) -> go.Figure:
    """Einheitliches Chart-Layout (IK-Farben, Legende, Hover).

    ``legende_unten=True`` platziert die Legende horizontal unter dem
    Diagramm, damit der Graph breiter dargestellt werden kann.
    """
    legende = dict(title="")
    if legende_unten:
        legende.update(orientation="h", yanchor="top", y=-0.18, x=0)
    fig.update_layout(
        title=dict(text=titel, font=dict(color=IK_BLAU, size=16)),
        hovermode="x unified",
        legend=legende,
        height=460,
        margin=dict(l=20, r=20, t=60, b=20),
        plot_bgcolor="white",
    )
    fig.update_xaxes(showgrid=False)
    fig.update_yaxes(gridcolor="#E8E8E8")
    return fig


def linien_chart(
    df: pd.DataFrame,
    wert_spalte: str,
    titel: str,
    y_achsen_titel: str,
    gruppierung: list,
    legende_unten: bool = False,
    hover_nachkommastellen: int = 2,
) -> go.Figure:
    """Liniendiagramm mit einer Linie je Gruppenkombination.

    ``legende_unten=True`` platziert die Legende horizontal unter dem
    Diagramm (Graph wird dadurch breiter dargestellt).
    ``hover_nachkommastellen`` steuert die Genauigkeit im Hover-Tooltip
    (Energiepreise benötigen 4 Stellen, da 2 Stellen z. B. 0,0552 als
    '0,06' anzeigen würden).
    """
    fig = go.Figure()
    gruppiert = df.dropna(subset=[wert_spalte]).groupby(gruppierung, sort=True)
    mehrere_gruppen = len(df[gruppierung[-1]].unique()) > 1 if gruppierung else False
    for i, (schluessel, teil) in enumerate(gruppiert):
        if not isinstance(schluessel, tuple):
            schluessel = (schluessel,)
        teil = teil.sort_values("time_date")
        name = schluessel[0] if not mehrere_gruppen else " · ".join(
            str(s) for s in schluessel
        )
        fig.add_trace(
            go.Scatter(
                x=teil["time_date"],
                y=teil[wert_spalte],
                mode="lines+markers",
                name=name,
                line=dict(color=FARBPALETTE[i % len(FARBPALETTE)]),
                hovertemplate=(
                    "%{x|%d.%m.%Y}<br>Wert: %{y:."
                    f"{hover_nachkommastellen}"
                    "f}<extra>%{fullData.name}</extra>"
                ),
            )
        )
    fig.update_yaxes(title_text=y_achsen_titel)
    return _basis_layout(fig, titel, legende_unten)


def dual_achsen_chart(
    df_links: pd.DataFrame,
    wert_links: str,
    df_rechts: pd.DataFrame,
    wert_rechts: str,
    titel: str,
    y1_titel: str,
    y2_titel: str,
    gruppierung: list,
    prefix_links: str = "",
    prefix_rechts: str = "",
    hover_nachkommastellen: int = 2,
) -> go.Figure:
    """Liniendiagramm mit zwei y-Achsen (links/rechts).

    Linke Reihen durchgezogen, rechte Reihen gestrichelt dargestellt.
    ``prefix_links``/``prefix_rechts`` werden den Legenden-Einträgen
    vorangestellt (z. B. 'Strom: ' / 'Gas: '), damit erkennbar ist,
    welche Linie welche Größe zeigt. ``hover_nachkommastellen`` steuert
    die Genauigkeit im Hover-Tooltip.
    """
    fig = go.Figure()
    farb_index = 0
    for wert, daten, gestrichelt, y_achse, prefix in (
        (wert_links, df_links, False, "y1", prefix_links),
        (wert_rechts, df_rechts, True, "y2", prefix_rechts),
    ):
        for schluessel, teil in daten.dropna(subset=[wert]).groupby(
            gruppierung, sort=True
        ):
            if not isinstance(schluessel, tuple):
                schluessel = (schluessel,)
            teil = teil.sort_values("time_date")
            name = prefix + " · ".join(str(s) for s in schluessel)
            farbe = FARBPALETTE[farb_index % len(FARBPALETTE)]
            farb_index += 1
            fig.add_trace(
                go.Scatter(
                    x=teil["time_date"],
                    y=teil[wert],
                    mode="lines",
                    name=name,
                    yaxis=y_achse,
                    line=dict(
                        color=farbe, dash="dash" if gestrichelt else "solid"
                    ),
                    hovertemplate=(
                        "%{x|%d.%m.%Y}<br>Wert: %{y:."
                        f"{hover_nachkommastellen}"
                        "f}<extra>%{fullData.name}</extra>"
                    ),
                )
            )
    fig.update_layout(
        yaxis=dict(title=y1_titel, side="left"),
        yaxis2=dict(title=y2_titel, overlaying="y", side="right"),
    )
    return _basis_layout(fig, titel)


def optionen_label(df: pd.DataFrame, label_spalte: str) -> list:
    """Sortierte, eindeutige Bezeichnungen einer Label-Spalte."""
    return sorted(df[label_spalte].dropna().unique())


def standard_label(optionen: list, *suchtexte: str) -> list:
    """Wählt als Standard die erste Option, die einen Suchtext enthält.

    Dient der robusten Vorauswahl in den Auswahlfeldern anhand der
    deutschen Eurostat-Bezeichnungen (z. B. 'Ohne Steuern und Abgaben').
    """
    for suchtext in suchtexte:
        for option in optionen:
            if suchtext.lower() in str(option).lower():
                return [option]
    return list(optionen[:1])


# ---------------------------------------------------------------------------
# Kurzformen für Filter-Chips (lange Bezeichnungen werden in den Chips der
# Auswahlfelder sonst abgeschnitten; Dropdown und Charts zeigen weiterhin
# die vollständigen Bezeichnungen)
# ---------------------------------------------------------------------------
def label_kurz_klammer(label: str) -> str:
    """Kürzt Bezeichnungen vor einer Klammer, z. B. 'Unbereinigte Daten
    (d.h. weder …)' -> 'Unbereinigte Daten'."""
    return str(label).split(" (")[0]


def band_kurz(label: str) -> str:
    """Kurzform der Verbrauchsband-Bezeichnung für Filter-Chips.

    'Verbrauch von 2 000 MWh bis 19 999 MWh - Gruppe ID'
    -> '2 000 – 19 999 MWh (ID)'
    """
    text = re.sub(r"^Verbrauch( von)? ", "", str(label))
    text = text.replace(" bis ", " – ")
    text = re.sub(r" - Gruppe (\w+)$", r" (\1)", text)
    return text[0].upper() + text[1:] if text else text


def nace_kurz(label: str) -> str:
    """Kurzform der NACE-Bezeichnung für Filter-Chips.

    'C2221 – Herstellung von Kunststoffplatten, …'
    -> 'C2221 – Kunststoffplatten, …'
    """
    return str(label).replace("Herstellung von ", "")


def zeitraum_filter(df: pd.DataFrame, von: int, bis: int) -> pd.DataFrame:
    """Filtert ein DataFrame auf den gewählten Jahresbereich."""
    jahre = df["time_date"].dt.year
    return df[jahre.between(von, bis)]


def lesebeispiel_arbeitskosten(
    df_eur: pd.DataFrame, df_rate: pd.DataFrame
) -> None:
    """Schreibt ein dynamisches Lesebeispiel unter die Arbeitskosten-Charts.

    Erläutert anhand der aktuell gewählten Daten, wie das Niveau (linkes
    Diagramm) und die Wachstumsrate (rechtes Diagramm) zu lesen und zu
    interpretieren sind.
    """
    geos = list(dict.fromkeys(df_eur["geo_label"].dropna()))
    if not geos:
        return
    referenz = "Deutschland" if "Deutschland" in geos else geos[0]
    eu = next(
        (g for g in geos if str(g).startswith("Europäische Union")), None
    )

    wert_ref, _, periode_ref, _ = letzter_wert_mit_vorjahr(
        df_eur[df_eur["geo_label"] == referenz], "value"
    )
    if wert_ref is None:
        return
    satz_niveau = (
        f"Im Jahr {periode_ref} betrugen die Arbeitskosten je geleisteter "
        f"Arbeitsstunde im Verarbeitenden Gewerbe (NACE C) in "
        f"**{referenz}** **{fmt_de(wert_ref, 1)} EUR**"
    )
    if eu and eu != referenz:
        wert_eu, _, periode_eu, _ = letzter_wert_mit_vorjahr(
            df_eur[df_eur["geo_label"] == eu], "value"
        )
        if wert_eu is not None:
            abstand = wert_ref - wert_eu
            satz_niveau += (
                f", im Durchschnitt der **{eu}** ({periode_eu}) dagegen "
                f"**{fmt_de(wert_eu, 1)} EUR** – {referenz} liegt damit um "
                f"{fmt_de(abs(abstand), 1)} EUR bzw. "
                f"{fmt_de(abs(abstand / wert_eu * 100), 0)} % "
                f"{'über' if abstand >= 0 else 'unter'} dem EU-Durchschnitt"
            )
    satz_niveau += "."

    satz_rate = ""
    wert_rate, _, periode_rate, _ = letzter_wert_mit_vorjahr(
        df_rate[df_rate["geo_label"] == referenz], "value"
    )
    if wert_rate is not None:
        satz_rate = (
            f" Die **Wachstumsrate** (rechtes Diagramm) zeigt die "
            f"prozentuale Veränderung gegenüber der Vorperiode: Für "
            f"{referenz} beträgt sie {periode_rate} "
            f"**{fmt_de(wert_rate, 1)} %**. Eine fallende Kurve bedeutet "
            f"dabei keinen Rückgang der Arbeitskosten, sondern einen "
            f"schwächeren Anstieg gegenüber dem Vorjahr."
        )
    st.info(f"**Lesebeispiel:** {satz_niveau}{satz_rate}")


def _referenz_und_eu(geos: list) -> tuple[str | None, str | None]:
    """Wählt Referenzland (bevorzugt Deutschland) und EU-Vergleichswert."""
    if not geos:
        return None, None
    referenz = "Deutschland" if "Deutschland" in geos else geos[0]
    eu = next(
        (g for g in geos if str(g).startswith("Europäische Union")), None
    )
    return referenz, eu


def lesebeispiel_energie(
    strom_df: pd.DataFrame, gas_df: pd.DataFrame, auswahl_tax: list
) -> None:
    """Schreibt ein dynamisches Lesebeispiel unter das Energiepreis-Chart.

    Erläutert anhand der aktuellen Auswahl Strom- und Gaspreisniveau sowie
    die Zuordnung der durchgezogenen/gestrichelten Linien.
    """
    geos = list(
        dict.fromkeys(
            pd.concat([strom_df["geo_label"], gas_df["geo_label"]]).dropna()
        )
    )
    referenz, eu = _referenz_und_eu(geos)
    if referenz is None:
        return
    steuer_hinweis = (
        f" ({str(auswahl_tax[0]).lower()})" if len(auswahl_tax) == 1 else ""
    )

    teile = []
    wert_s, _, periode_s, _ = letzter_wert_mit_vorjahr(
        strom_df[strom_df["geo_label"] == referenz], "value_strom"
    )
    if wert_s is not None:
        band_s = (
            strom_df[strom_df["geo_label"] == referenz]
            .sort_values("time_date")["nrg_cons_label"]
            .iloc[-1]
        )
        satz = (
            f"Im **{periode_lang(periode_s)}** lag der Strompreis für "
            f"Nicht-Haushalte in **{referenz}**{steuer_hinweis} im "
            f"Verbrauchsband „{band_s}“ bei **{fmt_de(wert_s, 4)} EUR/kWh**"
        )
        if eu and eu != referenz:
            wert_eu, _, _, _ = letzter_wert_mit_vorjahr(
                strom_df[strom_df["geo_label"] == eu], "value_strom"
            )
            if wert_eu is not None:
                abweichung = (wert_s - wert_eu) / wert_eu * 100
                satz += (
                    f"; der EU-27-Durchschnitt lag bei "
                    f"{fmt_de(wert_eu, 4)} EUR/kWh – {referenz} liegt "
                    f"damit {fmt_de(abs(abweichung), 0)} % "
                    f"{'über' if abweichung >= 0 else 'unter'} dem "
                    f"EU-Durchschnitt"
                )
        teile.append(satz + ".")

    wert_g, _, periode_g, _ = letzter_wert_mit_vorjahr(
        gas_df[gas_df["geo_label"] == referenz], "value_gas"
    )
    if wert_g is not None:
        band_g = (
            gas_df[gas_df["geo_label"] == referenz]
            .sort_values("time_date")["nrg_cons_label"]
            .iloc[-1]
        )
        satz_g = (
            f"Der Gaspreis im Verbrauchsband „{band_g}“ betrug im "
            f"{periode_lang(periode_g)} für {referenz} "
            f"**{fmt_de(wert_g, 4)} EUR/kWh**"
        )
        if eu and eu != referenz:
            wert_g_eu, _, _, _ = letzter_wert_mit_vorjahr(
                gas_df[gas_df["geo_label"] == eu], "value_gas"
            )
            if wert_g_eu is not None:
                abweichung_g = (wert_g - wert_g_eu) / wert_g_eu * 100
                satz_g += (
                    f"; der EU-27-Durchschnitt lag bei "
                    f"{fmt_de(wert_g_eu, 4)} EUR/kWh – {referenz} liegt "
                    f"damit {fmt_de(abs(abweichung_g), 0)} % "
                    f"{'über' if abweichung_g >= 0 else 'unter'} dem "
                    f"EU-Durchschnitt"
                )
        teile.append(satz_g + ".")
    teile.append(
        "Zur Einordnung: **Durchgezogene Linien** zeigen Strompreise "
        "(linke Achse), **gestrichelte Linien** Gaspreise (rechte Achse)."
    )
    st.info(f"**Lesebeispiel:** {' '.join(teile)}")


def lesebeispiel_industrie(
    df_prod: pd.DataFrame, df_preis: pd.DataFrame, auswahl_unit: list
) -> None:
    """Schreibt ein dynamisches Lesebeispiel unter die Industrie-Charts.

    Interpretiert Indexwerte relativ zum Basisjahr (= Jahresdurchschnitt
    des Basisjahres, nicht Vorjahresmonat) bzw. – bei Auswahl der
    Veränderungsrate – die prozentuale Veränderung zur Vorperiode.
    Zieht zusätzlich einen Vergleich zum EU-27-Durchschnitt.
    """
    geos = list(dict.fromkeys(df_prod["geo_label"].dropna()))
    referenz, eu = _referenz_und_eu(geos)
    naces = list(dict.fromkeys(df_prod["nace_r2_label"].dropna()))
    if referenz is None or not naces:
        return
    nace = naces[0]

    def _letzter(df: pd.DataFrame, spalte: str, geo: str):
        """Aktuellster Wert einer NACE-Reihe für ein Gebiet."""
        return letzter_wert_mit_vorjahr(
            df[
                (df["geo_label"] == geo)
                & (df["nace_r2_label"] == nace)
            ],
            spalte,
        )

    wert_p, _, periode_p, _ = _letzter(df_prod, "value_inpr", referenz)
    if wert_p is None:
        return
    wert_e, _, periode_e, _ = _letzter(df_preis, "value_inppd", referenz)
    wert_p_eu = wert_e_eu = periode_p_eu = periode_e_eu = None
    if eu and eu != referenz:
        wert_p_eu, _, periode_p_eu, _ = _letzter(df_prod, "value_inpr", eu)
        wert_e_eu, _, periode_e_eu, _ = _letzter(df_preis, "value_inppd", eu)

    basisjahr = next(
        (
            treffer.group(1)
            for u in auswahl_unit
            if (treffer := re.search(r"(\d{4})\s*=\s*100", str(u)))
        ),
        None,
    )
    ist_index = basisjahr is not None
    if ist_index:
        # Präzise Formulierung: Das Basisjahr (= 100) bezieht sich auf den
        # Jahresdurchschnitt des Basisjahres – nicht auf einen
        # Vorjahresmonat und nicht auf einen hier berechneten Durchschnitt.
        satz = (
            f"Im **{periode_lang(periode_p)}** lag der Produktionsindex für "
            f"„{nace}“ in **{referenz}** bei **{fmt_de(wert_p, 1)}** – d. h. "
            f"die Produktion dieses Monats liegt "
            f"{fmt_de(abs(wert_p - 100), 1)} % "
            f"{'über' if wert_p >= 100 else 'unter'} dem **durchschnittlichen "
            f"Monatsniveau des Jahres {basisjahr}** (Basisjahr "
            f"{basisjahr} = 100, Jahresdurchschnitt)."
        )
        if wert_p_eu is not None:
            eu_zeit = (
                "im selben Monat" if periode_p_eu == periode_p
                else f"im {periode_lang(periode_p_eu)}"
            )
            differenz = wert_p - wert_p_eu
            satz += (
                f" Zum Vergleich erreichte der **{eu}** {eu_zeit} "
                f"{fmt_de(wert_p_eu, 1)} Punkte – {referenz} liegt damit "
                f"{fmt_de(abs(differenz), 1)} Indexpunkte "
                f"{'über' if differenz >= 0 else 'unter'} dem "
                f"EU-Durchschnitt."
            )
        if wert_e is not None:
            satz += (
                f" Der Erzeugerpreisindex stand in {referenz} im "
                f"{periode_lang(periode_e)} bei {fmt_de(wert_e, 1)}, d. h. "
                f"die Erzeugerpreise liegen {fmt_de(wert_e - 100, 1)} % "
                f"über dem Durchschnitt des Basisjahres {basisjahr}."
            )
            if wert_e_eu is not None:
                eu_zeit_e = (
                    "im selben Monat" if periode_e_eu == periode_e
                    else f"im {periode_lang(periode_e_eu)}"
                )
                satz += (
                    f" Der EU-27-Durchschnitt lag {eu_zeit_e} bei "
                    f"{fmt_de(wert_e_eu, 1)} Punkten."
                )
    else:
        satz = (
            f"Im **{periode_lang(periode_p)}** betrug die Veränderung der "
            f"Produktion („{nace}“, {referenz}) gegenüber der Vorperiode "
            f"**{fmt_de(wert_p, 1)} %**"
        )
        if wert_e is not None:
            satz += (
                f", die Veränderung der Erzeugerpreise "
                f"{fmt_de(wert_e, 1)} %"
            )
        satz += "."
        if wert_p_eu is not None:
            satz += (
                f" Im EU-27-Durchschnitt lag die Produktionsveränderung "
                f"bei {fmt_de(wert_p_eu, 1)} %."
            )
    if ist_index:
        satz += (
            " Zur Einordnung: Der **Industrieproduktionsindex** misst die "
            "mengenmäßige Entwicklung der industriellen Produktion "
            f"(Produktionsvolumen) gegenüber dem Basisjahr {basisjahr} "
            "(= 100) – Werte unter 100 bedeuten ein niedrigeres "
            f"Produktionsniveau als im Jahr {basisjahr}. "
            "Der **Erzeugerpreisindex** misst die "
            "Preisentwicklung der im Inland abgesetzten "
            "Industrieerzeugnisse (Verkaufspreise ab Werk) gegenüber "
            "demselben Basisjahr."
        )
    else:
        satz += (
            " Zur Einordnung: Beide Reihen zeigen hier die prozentuale "
            "Veränderung gegenüber der Vorperiode (Vormonat) – der "
            "Industrieproduktionsindex bildet die Mengenentwicklung der "
            "Produktion ab, der Erzeugerpreisindex die Preisentwicklung "
            "der im Inland abgesetzten Erzeugnisse."
        )
    st.info(f"**Lesebeispiel:** {satz}")


def anzeige_tabelle(df: pd.DataFrame) -> pd.DataFrame:
    """Bereitet ein DataFrame für die Datenanzeige im Dashboard vor.

    Blendet technische bzw. konstante Hilfsspalten aus (``time_date`` wird
    nur für die Charts benötigt, die Periode steht bereits in ``time``;
    ``freq_label`` ist je Datensatz konstant; ``status``-Flags werden nicht
    angezeigt) und ordnet die Spalten lesbar an: Land, Periode und Wert(e)
    zuerst. ``time`` wird als Text dargestellt, damit Jahre nicht als Zahl
    mit Tausendertrennzeichen gerendert werden.
    """
    anzeige = df.drop(
        columns=["time_date", "freq_label", "status"], errors="ignore"
    ).copy()
    if "time" in anzeige.columns:
        anzeige["time"] = anzeige["time"].astype(str)
    erste = [c for c in ("geo_label", "time") if c in anzeige.columns]
    werte = [
        c
        for c in ("value", "value_inpr", "value_inppd",
                  "value_strom", "value_gas")
        if c in anzeige.columns
    ]
    rest = [c for c in anzeige.columns if c not in erste + werte]
    return anzeige[erste + werte + rest]


def nace_label_kombiniere(df: pd.DataFrame) -> pd.DataFrame:
    """Kombiniert NACE-Code und Bezeichnung ('C2221 – Herstellung von …').

    Die Rohdaten enthalten getrennte Code-/Label-Spalten; in den finalen
    Dateien ist die Kombination bereits enthalten (merge_data.py).
    """
    if "nace_r2" in df.columns and "nace_r2_label" in df.columns:
        df = df.copy()
        kombiniert = (
            df["nace_r2"].astype(str) + " – " + df["nace_r2_label"].astype(str)
        )
        df["nace_r2_label"] = kombiniert.where(
            df["nace_r2"].notna(), df["nace_r2_label"]
        )
    return df


def ohne_code_spalten(df: pd.DataFrame) -> pd.DataFrame:
    """Entfernt Abkürzungs-Spalten, zu denen eine Label-Spalte existiert
    (für die Anzeige der Rohdaten, die noch Code-Spalten enthalten)."""
    zu_entfernen = [
        spalte
        for spalte in df.columns
        if not spalte.endswith("_label")
        and f"{spalte}_label" in df.columns
        and spalte != "time"
    ]
    return df.drop(columns=zu_entfernen, errors="ignore")


# ---------------------------------------------------------------------------
# Daten laden (mit Fehlermeldung bei fehlenden Dateien)
# ---------------------------------------------------------------------------
def lade_alle_tabellen():
    """Lädt die benötigten Tabellen oder bricht mit Fehlermeldung ab.

    Arbeitskosten und Energiepreise kommen aus den finalen (gemergten)
    Dateien. Industrieproduktion & Erzeugerpreise werden direkt aus den
    Rohdaten der beiden STS-Datensätze gelesen (vollständige Länder-
    abdeckung); die Merge-Datei bleibt separater Pipeline-Output.
    """
    dateien = [
        config.FINAL_LC,
        config.FINAL_MERGE_ENERGIE,
        config.DATASETS["sts_inpr_m"]["datei"],
        config.DATASETS["sts_inppd_m"]["datei"],
    ]
    fehlend = [d for d in dateien if not (config.OUTPUT_DIR / d).exists()]
    if fehlend:
        st.error(
            "Die folgenden Excel-Dateien wurden im Output-Pfad nicht "
            f"gefunden: {', '.join(fehlend)}\n\n"
            "Bitte zuerst den Datenabruf und den Merge ausführen:\n"
            "1. `python fetch_data.py`\n"
            "2. `python merge_data.py`"
        )
        st.stop()
    return tuple(
        lade_tabelle(d, (config.OUTPUT_DIR / d).stat().st_mtime)
        for d in dateien
    )


# ---------------------------------------------------------------------------
# Seitenaufbau
# ---------------------------------------------------------------------------
def main() -> None:
    df_arbeit, df_energie, df_inpr, df_inppd = lade_alle_tabellen()
    # Rohdaten für Tab 3 vorbereiten: NACE-Code in die Bezeichnung
    # aufnehmen und Wert-Spalten herkunftsbezogen benennen (gleiche Namen
    # wie in der bisherigen Merge-Datei)
    df_inpr = nace_label_kombiniere(df_inpr).rename(
        columns={"value": "value_inpr"}
    )
    df_inppd = nace_label_kombiniere(df_inppd).rename(
        columns={"value": "value_inppd"}
    )

    # --- Header: Logo mittig, Titel, Untertitel mit Datenstand --------------
    # Header im Stil des IK-Wirtschafts-Dashboards: kleines zentriertes
    # Logo (feste Pixelbreite, versionsunabhängig via HTML-Einbettung),
    # Titel und Unterzeile in IK-Blau, Datenstand dezent darunter
    logo = finde_logo()
    if logo:
        logo_b64 = base64.b64encode(
            open(logo, "rb").read()
        ).decode()
        st.markdown(
            "<div style='text-align:center'>"
            f"<img src='data:image/jpeg;base64,{logo_b64}' width='180'>"
            "</div>",
            unsafe_allow_html=True,
        )
    st.markdown(
        f"<h1 style='text-align:center; color:{IK_BLAU}; margin-bottom:0'>"
        "IK Dashboard zum Standortvergleich Deutschland – Europa</h1>",
        unsafe_allow_html=True,
    )
    st.markdown(
        f"<p style='text-align:center; color:{IK_BLAU}; font-size:1.4rem; "
        "font-weight:600; margin-top:0.2rem'>"
        "Kunststoffverpackungen und -folienindustrie</p>",
        unsafe_allow_html=True,
    )
    st.markdown(
        "<p style='text-align:center; color:grey'>"
        f"Datenstand (letzter Eurostat-Abruf): {datenstand()}"
        "</p>",
        unsafe_allow_html=True,
    )
    st.divider()

    # --- Sidebar: globale Filter --------------------------------------------
    st.sidebar.header("Filter")

    alle_geos = sorted(
        set(df_arbeit["geo_label"].dropna())
        | set(df_energie["geo_label"].dropna())
        | set(df_inpr["geo_label"].dropna())
        | set(df_inppd["geo_label"].dropna())
    )
    # Standardauswahl: Deutschland + EU-27
    standard_geos = [
        g for g in alle_geos
        if g == "Deutschland" or g.startswith("Europäische Union - 27")
    ]
    auswahl_geos = st.sidebar.multiselect(
        "Länder / Regionen",
        options=alle_geos,
        default=standard_geos,
        format_func=label_kurz_klammer,
    )
    if not auswahl_geos:
        st.warning("Bitte mindestens ein Land / eine Region auswählen.")
        st.stop()

    alle_jahre = pd.concat(
        [df_arbeit["time_date"], df_energie["time_date"],
         df_inpr["time_date"], df_inppd["time_date"]]
    ).dt.year
    jahr_min = max(int(alle_jahre.min()), int(config.START_PERIOD))
    jahr_max = int(alle_jahre.max())
    von_bis = st.sidebar.slider(
        "Zeitraum (Jahre)",
        min_value=jahr_min,
        max_value=jahr_max,
        value=(jahr_min, jahr_max),
    )
    st.sidebar.caption(
        "Weitere, kontextabhängige Auswahlfelder finden sich in den "
        "jeweiligen Tabs."
    )

    # --- Tabs ----------------------------------------------------------------
    tab_arbeit, tab_energie, tab_industrie = st.tabs(
        ["Arbeitskosten", "Energiepreise",
         "Industrieproduktion & Erzeugerpreise"]
    )

    # == Tab 1: Arbeitskosten =================================================
    with tab_arbeit:
        df = zeitraum_filter(
            df_arbeit[df_arbeit["geo_label"].isin(auswahl_geos)], *von_bis
        )
        # Optionen aus dem Gesamtdatensatz ableiten, damit stets alle
        # auswählbaren Ausprägungen sichtbar bleiben (inkl. der gewählten)
        lcstruct_optionen = optionen_label(df_arbeit, "lcstruct_label")
        auswahl_lcstruct = st.multiselect(
            "Arbeitskostenstruktur",
            options=lcstruct_optionen,
            default=standard_label(lcstruct_optionen, "Arbeitskosten für LCI"),
            format_func=label_kurz_klammer,
        )
        df = df[df["lcstruct_label"].isin(auswahl_lcstruct)]

        if df.empty:
            st.info("Keine Daten für die gewählte Filterkombination.")
        else:
            df_eur = df[df["unit_label"] == "Euro"]
            df_rate = df[
                df["unit_label"].str.contains("Wachstumsrate", na=False)
            ]
            spalte_eur, spalte_rate = st.columns(2)
            with spalte_eur:
                st.subheader("Arbeitskosten in Euro")
                kpi_kacheln(
                    df_eur, "value", auswahl_geos, "EUR", 1,
                    delta_einheit=" EUR",
                )
                st.plotly_chart(
                    linien_chart(
                        df_eur, "value",
                        "Arbeitskosten in Euro (NACE C – Verarbeitendes Gewerbe)",
                        "Euro",
                        ["geo_label", "lcstruct_label"],
                    ),
                    use_container_width=True,
                    config={"locale": "de"},
                )
            with spalte_rate:
                st.subheader("Wachstumsrate ggü. Vorperiode")
                kpi_kacheln(
                    df_rate, "value", auswahl_geos, "%", 1,
                    delta_einheit=" Prozentpunkte",
                )
                st.plotly_chart(
                    linien_chart(
                        df_rate, "value",
                        "Wachstumsrate gegenüber der Vorperiode (Werte in Euro)",
                        "%",
                        ["geo_label", "lcstruct_label"],
                    ),
                    use_container_width=True,
                    config={"locale": "de"},
                )
            lesebeispiel_arbeitskosten(df_eur, df_rate)
            with st.expander("Daten anzeigen"):
                st.dataframe(
                    anzeige_tabelle(df),
                    use_container_width=True,
                    hide_index=True,
                )

    # == Tab 2: Energiepreise =================================================
    with tab_energie:
        df = zeitraum_filter(
            df_energie[df_energie["geo_label"].isin(auswahl_geos)], *von_bis
        )
        # Filteroptionen aus dem Gesamtdatensatz ableiten, damit stets alle
        # auswählbaren Ausprägungen sichtbar bleiben (inkl. der gewählten).
        strom_baender = optionen_label(
            df_energie[df_energie["value_strom"].notna()], "nrg_cons_label"
        )
        gas_baender = optionen_label(
            df_energie[
                df_energie["value_gas"].notna()
                & (df_energie["unit_label"] == "Kilowattstunde")
            ],
            "nrg_cons_label",
        )
        tax_optionen = optionen_label(df_energie, "tax_label")

        # 1. Verbrauchsbänder Strom/Gas, 2. Steuerkomponente
        b1, b2 = st.columns(2)
        with b1:
            auswahl_strom_band = st.multiselect(
                "Verbrauchsband Strom",
                options=strom_baender,
                default=standard_label(strom_baender, "Gruppe ID"),
                format_func=band_kurz,
            )
        with b2:
            auswahl_gas_band = st.multiselect(
                "Verbrauchsband Gas",
                options=gas_baender,
                default=standard_label(gas_baender, "Gruppe I3"),
                format_func=band_kurz,
            )
        auswahl_tax = st.multiselect(
            "Steuerkomponente",
            options=tax_optionen,
            default=standard_label(tax_optionen, "Ohne Steuern"),
        )

        df = df[df["tax_label"].isin(auswahl_tax)]
        # Gas wird fest in Kilowattstunde betrachtet (wie Strom), daher
        # kein Auswahlfeld für die Einheit. Die Währung ist nach der
        # Bereinigung der Daten ohnehin durchgehend Euro.
        strom_df = df[
            df["value_strom"].notna()
            & df["nrg_cons_label"].isin(auswahl_strom_band)
        ]
        gas_df = df[
            df["value_gas"].notna()
            & (df["unit_label"] == "Kilowattstunde")
            & df["nrg_cons_label"].isin(auswahl_gas_band)
        ]

        if strom_df.empty and gas_df.empty:
            st.info("Keine Daten für die gewählte Filterkombination.")
        else:
            st.subheader("Kennzahlen – aktuellste Werte")
            st.markdown("**Strom**")
            kpi_kacheln(
                strom_df, "value_strom", auswahl_geos, "EUR/kWh", 4,
                delta_einheit=" EUR/kWh",
            )
            st.markdown("**Gas**")
            kpi_kacheln(
                gas_df, "value_gas", auswahl_geos, "EUR/kWh", 4,
                delta_einheit=" EUR/kWh",
            )

            st.plotly_chart(
                dual_achsen_chart(
                    strom_df, "value_strom",
                    gas_df, "value_gas",
                    "Strom- und Gaspreise Nicht-Haushalte im Zeitverlauf",
                    "Strompreis (EUR/kWh)",
                    "Gaspreis (EUR/kWh)",
                    ["geo_label", "nrg_cons_label"],
                    prefix_links="Strom: ",
                    prefix_rechts="Gas: ",
                    hover_nachkommastellen=4,
                ),
                use_container_width=True,
                config={"locale": "de"},
            )
            lesebeispiel_energie(strom_df, gas_df, auswahl_tax)
            with st.expander("Daten anzeigen"):
                st.dataframe(
                    anzeige_tabelle(pd.concat([strom_df, gas_df])).drop(
                        columns=["unit_label"], errors="ignore"
                    ),
                    use_container_width=True,
                    hide_index=True,
                )

    # == Tab 3: Industrieproduktion & Erzeugerpreise ==========================
    with tab_industrie:
        # Quelle: Rohdaten der beiden STS-Datensätze (vollständige
        # Länderabdeckung, inkl. aller von Eurostat gemeldeten Werte)
        prod_basis = zeitraum_filter(
            df_inpr[df_inpr["geo_label"].isin(auswahl_geos)], *von_bis
        )
        preis_basis = zeitraum_filter(
            df_inppd[df_inppd["geo_label"].isin(auswahl_geos)], *von_bis
        )
        # Optionen aus den Gesamtdatensätzen ableiten, damit stets alle
        # auswählbaren Ausprägungen sichtbar bleiben (inkl. der gewählten)
        f1, f2, f3 = st.columns(3)
        with f1:
            nace_optionen = sorted(
                set(df_inpr["nace_r2_label"].dropna())
                | set(df_inppd["nace_r2_label"].dropna())
            )
            auswahl_nace = st.multiselect(
                "NACE-Abschnitt",
                options=nace_optionen,
                default=nace_optionen,
                format_func=nace_kurz,
            )
        with f2:
            s_adj_optionen = optionen_label(df_inpr, "s_adj_label")
            auswahl_s_adj = st.multiselect(
                "Kalender-/Saisonbereinigung (Produktionsindex)",
                options=s_adj_optionen,
                default=standard_label(
                    s_adj_optionen, "Saison- und kalenderbereinigte"
                ),
                format_func=label_kurz_klammer,
            )
        with f3:
            unit_optionen = sorted(
                set(df_inpr["unit_label"].dropna())
                | set(df_inppd["unit_label"].dropna())
            )
            auswahl_unit = st.multiselect(
                "Maßeinheit",
                options=unit_optionen,
                default=standard_label(unit_optionen, "=100"),
            )
        st.caption(
            "Hinweis: Die Bereinigungsauswahl gilt nur für den "
            "Produktionsindex – Erzeugerpreise liegen ausschließlich "
            "unbereinigt vor. EU-Aggregate des Produktionsindex "
            "veröffentlicht Eurostat nur kalender- bzw. saisonbereinigt "
            "(keine unbereinigten Werte)."
        )
        # Die Bereinigungsauswahl gilt nur für den Produktionsindex;
        # Erzeugerpreise liegen ohnehin nur unbereinigt (NSA) vor.
        df_prod = prod_basis[
            prod_basis["nace_r2_label"].isin(auswahl_nace)
            & prod_basis["s_adj_label"].isin(auswahl_s_adj)
            & prod_basis["unit_label"].isin(auswahl_unit)
            & prod_basis["value_inpr"].notna()
        ]
        df_preis = preis_basis[
            preis_basis["nace_r2_label"].isin(auswahl_nace)
            & preis_basis["unit_label"].isin(auswahl_unit)
            & preis_basis["value_inppd"].notna()
        ]

        if df_prod.empty and df_preis.empty:
            st.info("Keine Daten für die gewählte Filterkombination.")
        else:
            if len(auswahl_unit) == 1:
                einzig = str(auswahl_unit[0])
                basis = re.search(r"(\d{4})\s*=\s*100", einzig)
                unit_kurz = (
                    f"Index ({basis.group(1)}=100)" if basis
                    else "% ggü. Vorperiode" if "Prozent" in einzig
                    else einzig
                )
            else:
                unit_kurz = "Wert"
            # Zwei getrennte Abbildungen: Produktion links, Preise rechts
            spalte_prod, spalte_preis = st.columns(2)
            with spalte_prod:
                st.subheader("Industrieproduktionsindex")
                if df_prod.empty:
                    st.info("Keine Produktionsdaten für die gewählte "
                            "Bereinigung (EU-Aggregate nur bereinigt "
                            "verfügbar).")
                else:
                    kpi_kacheln(
                        df_prod, "value_inpr", auswahl_geos, "", 2,
                        delta_farbe="normal",
                    )
                    st.plotly_chart(
                        linien_chart(
                            df_prod, "value_inpr",
                            "Industrieproduktionsindex "
                            f"({unit_kurz}) – NACE C2221/C2222",
                            unit_kurz,
                            ["geo_label", "nace_r2_label"],
                            legende_unten=True,
                        ),
                        use_container_width=True,
                        config={"locale": "de"},
                    )
            with spalte_preis:
                st.subheader("Erzeugerpreisindex (Inlandsmarkt)")
                kpi_kacheln(
                    df_preis, "value_inppd", auswahl_geos, "", 2,
                    delta_farbe="normal",
                )
                st.plotly_chart(
                    linien_chart(
                        df_preis, "value_inppd",
                        f"Erzeugerpreisindex ({unit_kurz}) – NACE C2221/C2222",
                        unit_kurz,
                        ["geo_label", "nace_r2_label"],
                        legende_unten=True,
                    ),
                    use_container_width=True,
                    config={"locale": "de"},
                )
            lesebeispiel_industrie(df_prod, df_preis, auswahl_unit)
            with st.expander("Daten anzeigen"):
                st.dataframe(
                    anzeige_tabelle(
                        ohne_code_spalten(pd.concat([df_prod, df_preis]))
                    ).rename(
                        columns={
                            "value_inpr": "value_inpr "
                            "(Industrieproduktionsindex)",
                            "value_inppd": "value_inppd "
                            "(Erzeugerpreisindex, Inlandsmarkt)",
                        }
                    ),
                    use_container_width=True,
                    hide_index=True,
                )

    # --- Footer --------------------------------------------------------------
    st.divider()
    st.caption(
        "Quelle: Eurostat – Datensätze lc_lci_lev, nrg_pc_205, nrg_pc_203, "
        "sts_inpr_m, sts_inppd_m"
    )
    st.markdown(
        "**Kontakt bei Fragen:**  \n"
        "**Referat für Wirtschaft**  \n"
        "**IK Industrieverband e.V.**  \n"
        "**Dr. Laura C. Müller**  \n"
        "**L.Mueller@Kunststoffverpackungen.de**"
    )


if __name__ == "__main__":
    main()
