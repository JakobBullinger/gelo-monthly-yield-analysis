import streamlit as st
import pandas as pd
import re
import io
import numpy as np
from datetime import datetime

# ――― Hilfsfunktionen ──────────────────────────────────────────────────────
def fmt_thousands(num: float | int, decimals: int = 0) -> str:
    """
    Gibt eine Zahl als String mit Punkt als Tausender-Separator zurück.
    Beispiel: 23722  →  '23.722'
    """
    formatted = f"{num:,.{decimals}f}"        # Standardform: 23,722
    return formatted.replace(",", ".")         # Ersetzt Komma durch Punkt

def to_excel(df: pd.DataFrame) -> bytes:
    """Schreibt ein DataFrame in eine Excel-Datei im Memory und gibt die Bytes zurück."""
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df.to_excel(writer, index=False, sheet_name='Monatsanalyse')
    return output.getvalue()

def get_staerke_klasse(d: float) -> str:
    """Ordnet einen Durchmesser einer Stärkeklasse zu."""
    if d < 100:   return "0"
    if d < 150:   return "1a"
    if d < 200:   return "1b"
    if d < 250:   return "2a"
    if d < 300:   return "2b"
    if d < 350:   return "3a"
    if d < 400:   return "3b"
    return "unbekannt"

# ――― Haupt-App ─────────────────────────────────────────────────────────────
def main() -> None:
    st.set_page_config(
        page_title="Monatsausbeute Analyse",
        layout="wide",
        initial_sidebar_state="expanded"
    )

    # — CSS: kleinere Schrift in Metrics —
    st.markdown(
        """
        <style>
        .stMetricLabel, .stMetricValue {font-size: 0.9rem !important;}
        .stMetricDelta {font-size: 0.8rem !important;}
        </style>
        """,
        unsafe_allow_html=True
    )

    # — Sidebar —
    st.sidebar.header("🔧 Einstellungen")
    st.sidebar.markdown(
        "Lade hier deine Tages-Excel-Dateien eines Monats hoch.\n\n"
        "- Akzeptiert: `.xlsx`, `.xls`\n"
        "- Dateiname muss `Ausbeuteanalyse_YYYY-MM-DD` enthalten."
    )
    uploaded = st.sidebar.file_uploader(
        "Dateien auswählen",
        type=["xlsx", "xls"],
        accept_multiple_files=True
    )

    # — Kopfbereich —
    st.title("📊 Monatsanalyse Ausbeute")
    st.markdown(
        "Diese App fasst die täglichen Ausbeute-Reports pro Auftrag und Dimension "
        "über einen oder mehrere Monate zusammen und berechnet zusätzliche Kennzahlen."
    )

    if not uploaded:
        st.warning("Bitte mindestens eine Excel-Datei hochladen.")
        return

    # — Einlesen aller Dateien —
    dfs = [pd.read_excel(f) for f in uploaded]
    df_all = pd.concat(dfs, ignore_index=True)

    # — Auftragsnummer & Cleaning —
    df_all['Auftragsnummer'] = df_all['Auftrag'].astype(str).str.extract(r'^(\d{5})')
    df_all['Auftrag_clean'] = df_all['Auftrag'].astype(str).str.extract(
        r'^(\d{5}\s*-\s*(?:[A-Za-zÄÖÜäöü]*\s*)?\d+x\d+(?:x\d+)?)'
    )[0]

    # — Trennen Gesamt- vs. Dimensionszeilen —
    df_overall = df_all[df_all['Stämme'] != 0].copy()
    df_dim     = df_all[df_all['Stämme'] == 0].copy()

    # — Aggregation Gesamt —
    agg_overall = (
        df_overall
        .groupby(['Auftragsnummer', 'Auftrag_clean'], as_index=False)
        .agg({
            'Stämme': 'sum',
            'Volumen_Eingang': 'sum',
            'Durchschn_Stämme': 'mean',
            'Teile': 'sum',
            'Laufzeit_Minuten': 'sum'
        })
        .rename(columns={'Teile': 'Teile_gesamt', 'Auftrag_clean': 'Auftrag'})
    )
    agg_overall['Durchmesser'] = np.sqrt(
        agg_overall['Volumen_Eingang'] /
        (np.pi * agg_overall['Durchschn_Stämme'] * agg_overall['Stämme'])
    ) * 20000

    # — Aggregation Dimensionen —
    dim_cols = ['Teile', 'Brutto_Volumen', 'Netto_Volumen',
                'CE', 'SF', 'SI', 'IND', 'NSI', 'Q_V', 'Ausschuss']
    grouped_dim = (
        df_dim
        .groupby(['Auftragsnummer', 'Dimension'], as_index=False)[dim_cols]
        .sum()
        .rename(columns={'Teile': 'Teile_dim'})
    )

    # — Merge & Zusatzkennzahlen —
    merged = pd.merge(grouped_dim, agg_overall, on='Auftragsnummer', how='left')
    merged['Brutto_Ausschuss'] = np.where(
        merged['Brutto_Volumen'] > 0,
        merged['Ausschuss'] / merged['Brutto_Volumen'] * 100, 0
    )
    merged['Brutto_Ausbeute'] = np.where(
        merged['Volumen_Eingang'] > 0,
        merged['Brutto_Volumen'] / merged['Volumen_Eingang'] * 100, 0
    )
    merged['Netto_Ausbeute'] = np.where(
        merged['Volumen_Eingang'] > 0,
        merged['Netto_Volumen'] / merged['Volumen_Eingang'] * 100, 0
    )

    # — Original-Layout rekonstruieren —
    final_cols = [
        'Auftrag', 'Dimension',
        'Stämme', 'Volumen_Eingang', 'Durchschn_Stämme', 'Teile_gesamt',
        'Durchmesser', 'Stärke_Klasse', 'Laufzeit_Minuten', 'Vorschub(FM/h)',
        'Brutto_Volumen', 'Brutto_Ausschuss', 'Netto_Volumen',
        'Brutto_Ausbeute', 'Netto_Ausbeute',
        'CE', 'SF', 'SI', 'IND', 'NSI', 'Q_V', 'Ausschuss'
    ]
    final_data: list[dict] = []
    merged.sort_values(['Auftragsnummer', 'Dimension'], inplace=True)

    for nr in merged['Auftragsnummer'].unique():
        o = agg_overall.loc[agg_overall['Auftragsnummer'] == nr].iloc[0]

        # Gesamt-Zeile
        row = {
            'Auftrag': o['Auftrag'],
            'Dimension': '',
            'Stämme': o['Stämme'],
            'Volumen_Eingang': o['Volumen_Eingang'],
            'Durchschn_Stämme': o['Durchschn_Stämme'],
            'Teile_gesamt': o['Teile_gesamt'],
            'Durchmesser': o['Durchmesser'],
            'Stärke_Klasse': get_staerke_klasse(o['Durchmesser']),
            'Laufzeit_Minuten': o['Laufzeit_Minuten'],
            'Vorschub(FM/h)': (o['Volumen_Eingang'] / (o['Laufzeit_Minuten'] / 60)
                               if o['Laufzeit_Minuten'] else 0)
        }
        for c in ['Brutto_Volumen', 'Brutto_Ausschuss', 'Netto_Volumen',
                  'Brutto_Ausbeute', 'Netto_Ausbeute',
                  'CE', 'SF', 'SI', 'IND', 'NSI', 'Q_V', 'Ausschuss']:
            row[c] = 0
        final_data.append(row)

        # Dimension-Zeilen
        for _, d in merged[merged['Auftragsnummer'] == nr].iterrows():
            row_d = {
                'Auftrag': d['Auftrag'],
                'Dimension': d['Dimension'],
                'Stämme': 0,
                'Volumen_Eingang': 0,
                'Durchschn_Stämme': 0,
                'Teile_gesamt': d['Teile_dim'],
                'Durchmesser': 0,
                'Stärke_Klasse': '',
                'Laufzeit_Minuten': 0,
                'Vorschub(FM/h)': 0,
                'Brutto_Volumen': d['Brutto_Volumen'],
                'Brutto_Ausschuss': d['Brutto_Ausschuss'],
                'Netto_Volumen': d['Netto_Volumen'],
                'Brutto_Ausbeute': d['Brutto_Ausbeute'],
                'Netto_Ausbeute': d['Netto_Ausbeute'],
                'CE': d['CE'], 'SF': d['SF'], 'SI': d['SI'],
                'IND': d['IND'], 'NSI': d['NSI'], 'Q_V': d['Q_V'],
                'Ausschuss': d['Ausschuss']
            }
            final_data.append(row_d)

    final_df = pd.DataFrame(final_data, columns=final_cols)

    # — Runden auf drei Nachkommastellen —
    final_df = final_df.round(3)

    # ――― Kennzahlen-Dashboard ────────────────────────────────────────────
    total_input_volume = final_df['Volumen_Eingang'].sum()
    total_brutto = final_df['Brutto_Volumen'].sum()

    # Datumsspanne & Anzahl Tage aus Dateinamen
    dates: list[datetime.date] = []
    for f in uploaded:
        m = re.search(r'(\d{4}-\d{2}-\d{2})', f.name)
        if m:
            dates.append(datetime.strptime(m.group(1), '%Y-%m-%d').date())
    if dates:
        start, end = min(dates), max(dates)
        date_range_str = f"{start.strftime('%d.%m.%y')} - {end.strftime('%d.%m.%y')}"
        filename_range = f"{start.strftime('%d_%m_%Y')}_{end.strftime('%d_%m_%Y')}"
        num_days = len(set(dates))
    else:
        date_range_str = "–"
        filename_range = "unknown"
        num_days = 0

    # — Anzeige der Metriken (mit Punkt als Tausender-Separator) —
    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Gesamt Einschnittsvolumen", f"{fmt_thousands(total_input_volume)} m³")
    c2.metric("Gesamt Brutto-Volumen",    f"{fmt_thousands(total_brutto)} m³")
    c3.metric("Daten von bis", date_range_str)
    c4.metric("Anzahl Tage", f"{num_days}")

    # ――― Detail-Tabelle & Download ───────────────────────────────────────
    with st.expander("▶️ Detailtabelle anzeigen"):
        st.dataframe(final_df, use_container_width=True)

    filename = f"monatsanalyse_{filename_range}.xlsx"
    st.download_button(
        "📥 Als Excel herunterladen",
        data=to_excel(final_df),
        file_name=filename,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        help="Lädt die fertige Monatsauswertung als Excel-Datei."
    )

# ――― App starten ──────────────────────────────────────────────────────────
if __name__ == "__main__":
    main()
