import streamlit as st
import pandas as pd
import re
import io
import numpy as np

def to_excel(df):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df.to_excel(writer, index=False, sheet_name='Monatsanalyse')
    return output.getvalue()

def get_staerke_klasse(durchmesser):
    """Ermittelt die Stärke_Klasse anhand des Durchmessers."""
    if durchmesser < 100:
        return "0"
    elif 100 <= durchmesser < 150:
        return "1a"
    elif 150 <= durchmesser < 200:
        return "1b"
    elif 200 <= durchmesser < 250:
        return "2a"
    elif 250 <= durchmesser < 300:
        return "2b"
    elif 300 <= durchmesser < 350:
        return "3a"
    elif 350 <= durchmesser < 400:
        return "3b"
    else:
        return "unbekannt"

def main():
    st.title("Monatsanalyse Ausbeute – Original‑Layout mit einheitlichem Auftragstext")
    
    uploaded_files = st.file_uploader(
        "Bitte alle Excel‑Dateien eines Monats hochladen",
        accept_multiple_files=True,
        type=["xlsx", "xls"]
    )
    if not uploaded_files:
        st.info("Bitte Dateien auswählen.")
        return

    # 1) Einlesen aller Dateien
    dfs = []
    for f in uploaded_files:
        try:
            df = pd.read_excel(f)
            dfs.append(df)
        except Exception as e:
            st.error(f"Fehler beim Einlesen von {f.name}: {e}")
            return
    if not dfs:
        st.warning("Keine Daten geladen.")
        return
    df_all = pd.concat(dfs, ignore_index=True)

    # 2) Auftragsnummer extrahieren (erste 5 Ziffern)
    df_all['Auftragsnummer'] = df_all['Auftrag'].astype(str).str.extract(r'^(\d{5})')

    # ────────────────────────────────────────────────────────────────
    # 3) Auftragstext bereinigen, damit z.B. "Produktivitätsbericht"
    #    nicht als extra Variante gewertet wird
    df_all['Auftrag_clean'] = (
        df_all['Auftrag']
        .astype(str)
        .str.replace(r'\s*Produktivitätsbericht$', '', regex=True)
        .str.strip()
    )
    # ────────────────────────────────────────────────────────────────

    # 4) Gesamtzeilen (Stämme != 0) und Dimensionszeilen (Stämme == 0) trennen
    df_overall_days = df_all[df_all['Stämme'] != 0].copy()
    df_dim          = df_all[df_all['Stämme'] == 0].copy()

    # 5) Monatliche Aggregation der Gesamtzeilen:
    #    - Summe: Stämme, Volumen_Eingang, Teile, Laufzeit_Minuten
    #    - Mittelwert: Durchschn_Stämme
    #    - Neuberechnung Durchmesser
    agg_overall = df_overall_days.groupby(
        ['Auftragsnummer','Auftrag_clean'], as_index=False
    ).agg({
        'Stämme':             'sum',
        'Volumen_Eingang':    'sum',
        'Durchschn_Stämme':   'mean',
        'Teile':              'sum',
        'Laufzeit_Minuten':   'sum'
    }).rename(columns={'Teile':'Teile_gesamt',
                      'Auftrag_clean':'Auftrag'})

    agg_overall['Durchmesser'] = np.sqrt(
        agg_overall['Volumen_Eingang'] /
        (np.pi * agg_overall['Durchschn_Stämme'] * agg_overall['Stämme'])
    ) * 20000

    # 6) Aggregation der Dimensionszeilen (pro Auftrag & Dimension)
    dim_cols = [
        'Teile','Brutto_Volumen','Netto_Volumen',
        'CE','SF','SI','IND','NSI','Q_V','Ausschuss'
    ]
    grouped_dim = df_dim.groupby(
        ['Auftragsnummer','Dimension'], as_index=False
    )[dim_cols].sum().rename(columns={'Teile':'Teile_dim'})

    # 7) Zusammenführen von Dimensions‑ und Gesamt‑Daten
    merged = pd.merge(grouped_dim, agg_overall, on='Auftragsnummer', how='left')

    # 8) Berechnung weiterer Kennzahlen pro Dimension
    merged['Brutto_Ausschuss'] = merged.apply(
        lambda r: (r['Ausschuss']/r['Brutto_Volumen']*100) if r['Brutto_Volumen'] else 0,
        axis=1
    )
    merged['Brutto_Ausbeute'] = merged.apply(
        lambda r: (r['Brutto_Volumen']/r['Volumen_Eingang']*100) if r['Volumen_Eingang'] else 0,
        axis=1
    )
    merged['Netto_Ausbeute'] = merged.apply(
        lambda r: (r['Netto_Volumen']/r['Volumen_Eingang']*100) if r['Volumen_Eingang'] else 0,
        axis=1
    )

    # 9) Rekonstruktion des Original‑Layouts
    final_cols = [
        'Auftrag','Dimension',
        'Stämme','Volumen_Eingang','Durchschn_Stämme','Teile_gesamt',
        'Durchmesser','Stärke_Klasse','Laufzeit_Minuten','Vorschub(FM/h)',
        'Brutto_Volumen','Brutto_Ausschuss','Netto_Volumen',
        'Brutto_Ausbeute','Netto_Ausbeute',
        'CE','SF','SI','IND','NSI','Q_V','Ausschuss'
    ]
    final_data = []
    merged.sort_values(['Auftragsnummer','Dimension'], inplace=True)

    for nr in merged['Auftragsnummer'].unique():
        # Gesamtzeile
        o = agg_overall.loc[agg_overall['Auftragsnummer']==nr].iloc[0]
        row = {
            'Auftrag':          o['Auftrag'],
            'Dimension':        '',
            'Stämme':           o['Stämme'],
            'Volumen_Eingang':  o['Volumen_Eingang'],
            'Durchschn_Stämme': o['Durchschn_Stämme'],
            'Teile_gesamt':     o['Teile_gesamt'],
            'Durchmesser':      o['Durchmesser'],
            'Stärke_Klasse':    get_staerke_klasse(o['Durchmesser']),
            'Laufzeit_Minuten': o['Laufzeit_Minuten'],
            'Vorschub(FM/h)':   (o['Volumen_Eingang']/(o['Laufzeit_Minuten']/60)
                                 if o['Laufzeit_Minuten'] else 0)
        }
        # setze dimensions‑Spalten auf 0
        for c in ['Brutto_Volumen','Brutto_Ausschuss','Netto_Volumen',
                  'Brutto_Ausbeute','Netto_Ausbeute',
                  'CE','SF','SI','IND','NSI','Q_V','Ausschuss']:
            row[c] = 0
        final_data.append(row)

        # Dimensionszeilen
        for _, d in merged[merged['Auftragsnummer']==nr].iterrows():
            row_d = {
                'Auftrag':          d['Auftrag'],
                'Dimension':        d['Dimension'],
                'Stämme':           0,
                'Volumen_Eingang':  0,
                'Durchschn_Stämme': 0,
                'Teile_gesamt':     d['Teile_dim'],
                'Durchmesser':      0,
                'Stärke_Klasse':    '',
                'Laufzeit_Minuten': 0,
                'Vorschub(FM/h)':   0,
                'Brutto_Volumen':   d['Brutto_Volumen'],
                'Brutto_Ausschuss': d['Brutto_Ausschuss'],
                'Netto_Volumen':    d['Netto_Volumen'],
                'Brutto_Ausbeute':  d['Brutto_Ausbeute'],
                'Netto_Ausbeute':   d['Netto_Ausbeute'],
                'CE':               d['CE'],
                'SF':               d['SF'],
                'SI':               d['SI'],
                'IND':              d['IND'],
                'NSI':              d['NSI'],
                'Q_V':              d['Q_V'],
                'Ausschuss':        d['Ausschuss']
            }
            final_data.append(row_d)

    final_df = pd.DataFrame(final_data, columns=final_cols)

    # 10) Ausgabe und Download
    st.subheader("Monatsanalyse im Original‑Layout")
    st.dataframe(final_df)

    excel_bytes = to_excel(final_df)
    st.download_button(
        "Excel herunterladen",
        data=excel_bytes,
        file_name="monatsanalyse_original_layout.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

if __name__ == "__main__":
    main()
