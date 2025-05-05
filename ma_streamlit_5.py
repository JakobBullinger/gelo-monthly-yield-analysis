import streamlit as st
import pandas as pd
import re
import io

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
    st.title("Monatsanalyse Ausbeute – Original-Layout mit Stärke_Klasse")
    
    # Upload der Dateien
    uploaded_files = st.file_uploader(
        "Bitte alle Excel-Dateien hochladen",
        accept_multiple_files=True,
        type=["xlsx", "xls"]
    )
    if not uploaded_files:
        st.info("Bitte Dateien auswählen.")
        return
    
    # --- EINLESEN & AGGREGATION (vereinfacht) ---
    # (1) Dateien einlesen
    dataframes = []
    for file in uploaded_files:
        try:
            df = pd.read_excel(file)
            dataframes.append(df)
        except Exception as e:
            st.error(f"Fehler beim Laden der Datei {file.name}: {e}")
            return
    
    if not dataframes:
        st.warning("Keine gültigen Daten gefunden.")
        return
    
    df_all = pd.concat(dataframes, ignore_index=True)
    
    # (2) Auftragsnummer extrahieren
    df_all['Auftragsnummer'] = df_all['Auftrag'].astype(str).apply(
        lambda x: re.findall(r'^\d{5}', x)[0] if re.findall(r'^\d{5}', x) else x
    )
    
    # (3) Erkennen der Gesamtzeilen vs. Dimensionszeilen
    df_overall = df_all[df_all['Stämme'] != 0].copy()
    df_dim = df_all[df_all['Stämme'] == 0].copy()
    
    # (4) Aggregation der Dimensionsdaten:
    dim_cols = [
        'Teile', 'Brutto_Volumen', 'Netto_Volumen',
        'CE', 'SF', 'SI', 'IND', 'NSI', 'Q_V', 'Ausschuss'
    ]
    grouped_dim = df_dim.groupby(['Auftragsnummer', 'Dimension'], as_index=False)[dim_cols].sum()
    grouped_dim.rename(columns={'Teile': 'Teile_dim'}, inplace=True)
    
    # (5) Gesamtwerte aus der Overall-Tabelle:
    overall_cols = [
        'Auftragsnummer', 'Auftrag', 'Stämme', 'Volumen_Eingang',
        'Durchschn_Stämme', 'Teile', 'Durchmesser', 'Laufzeit_Minuten'
    ]
    df_overall_unique = (
        df_overall[overall_cols]
        .drop_duplicates(subset=['Auftragsnummer'])
        .rename(columns={'Teile': 'Teile_gesamt'})
    )
    
    # (6) Zusammenführen der Dimensionsdaten mit den Gesamtwerten (über Auftragsnummer)
    merged = pd.merge(grouped_dim, df_overall_unique, on='Auftragsnummer', how='left')
    
    # (7) Neue Kennzahlen pro Dimension
    merged['Brutto_Ausschuss'] = merged.apply(
        lambda row: (row['Ausschuss'] / row['Brutto_Volumen'] * 100) if row['Brutto_Volumen'] else 0,
        axis=1
    )
    merged['Brutto_Ausbeute'] = merged.apply(
        lambda row: (row['Brutto_Volumen'] / row['Volumen_Eingang'] * 100) if row['Volumen_Eingang'] else 0,
        axis=1
    )
    merged['Netto_Ausbeute'] = merged.apply(
        lambda row: (row['Netto_Volumen'] / row['Volumen_Eingang'] * 100) if row['Volumen_Eingang'] else 0,
        axis=1
    )
    
    # --- REKONSTRUKTION DES ORIGINAL-FORMATS ---
    
    # Festlegen der finalen Spaltenreihenfolge.
    # Neu: "Stärke_Klasse" wird nach "Durchmesser" eingefügt.
    final_columns = [
        'Auftrag',           # Spalte 1
        'Dimension',         # Spalte 2
        'Stämme',            # Spalte 3
        'Volumen_Eingang',   # Spalte 4
        'Durchschn_Stämme',  # Spalte 5
        'Teile',             # Spalte 6
        'Durchmesser',       # Spalte 7
        'Stärke_Klasse',     # Spalte 8
        'Laufzeit_Minuten',  # Spalte 9
        'Vorschub(FM/h)',    # Spalte 10
        'Brutto_Volumen',    # Spalte 11
        'Brutto_Ausschuss',  # Spalte 12
        'Netto_Volumen',     # Spalte 13
        'Brutto_Ausbeute',   # Spalte 14
        'Netto_Ausbeute',    # Spalte 15
        'CE',                # Spalte 16
        'SF',                # Spalte 17
        'SI',                # Spalte 18
        'IND',               # Spalte 19
        'NSI',               # Spalte 20
        'Q_V',               # Spalte 21
        'Ausschuss'          # Spalte 22
    ]
    
    # Zusammenbau der finalen Zeilen:
    final_data = []
    
    # Sortierung für konsistente Reihenfolge
    merged.sort_values(by=['Auftragsnummer', 'Dimension'], inplace=True)
    auftragsnummern = merged['Auftragsnummer'].unique()
    
    for nr in auftragsnummern:
        # 1) Gesamtzeile (erste Zeile pro Auftrag)
        row_overall_src = df_overall_unique.loc[df_overall_unique['Auftragsnummer'] == nr].iloc[0]
        row_overall = {}
        row_overall['Auftrag'] = row_overall_src['Auftrag']
        row_overall['Dimension'] = ""  # Gesamtzeile, daher keine Dimensionsbezeichnung
        row_overall['Stämme'] = row_overall_src['Stämme']
        row_overall['Volumen_Eingang'] = row_overall_src['Volumen_Eingang']
        row_overall['Durchschn_Stämme'] = row_overall_src['Durchschn_Stämme']
        row_overall['Teile'] = row_overall_src['Teile_gesamt']
        row_overall['Durchmesser'] = row_overall_src['Durchmesser']
        # Berechnung der Stärke_Klasse anhand des Durchmessers
        row_overall['Stärke_Klasse'] = get_staerke_klasse(row_overall_src['Durchmesser'])
        row_overall['Laufzeit_Minuten'] = row_overall_src['Laufzeit_Minuten']
        # Berechnung des Vorschubs (FM/h) – falls Laufzeit > 0
        if row_overall_src['Laufzeit_Minuten'] and row_overall_src['Laufzeit_Minuten'] != 0:
            row_overall['Vorschub(FM/h)'] = row_overall_src['Volumen_Eingang'] / (row_overall_src['Laufzeit_Minuten'] / 60)
        else:
            row_overall['Vorschub(FM/h)'] = 0
        
        # Für die Gesamtzeile sind die dimensionsspezifischen Spalten 0
        row_overall['Brutto_Volumen'] = 0
        row_overall['Brutto_Ausschuss'] = 0
        row_overall['Netto_Volumen'] = 0
        row_overall['Brutto_Ausbeute'] = 0
        row_overall['Netto_Ausbeute'] = 0
        row_overall['CE'] = 0
        row_overall['SF'] = 0
        row_overall['SI'] = 0
        row_overall['IND'] = 0
        row_overall['NSI'] = 0
        row_overall['Q_V'] = 0
        row_overall['Ausschuss'] = 0
        
        final_data.append(row_overall)
        
        # 2) Dimensionszeilen
        sub = merged.loc[merged['Auftragsnummer'] == nr].copy()
        for _, dim_row in sub.iterrows():
            row_dim = {}
            row_dim['Auftrag'] = dim_row['Auftrag']
            row_dim['Dimension'] = dim_row['Dimension']
            # Gesamtspalten auf 0
            row_dim['Stämme'] = 0
            row_dim['Volumen_Eingang'] = 0
            row_dim['Durchschn_Stämme'] = 0
            row_dim['Teile'] = dim_row['Teile_dim']
            row_dim['Durchmesser'] = 0
            # Bei Dimensionszeilen bleibt die Stärke_Klasse leer
            row_dim['Stärke_Klasse'] = ""
            row_dim['Laufzeit_Minuten'] = 0
            row_dim['Vorschub(FM/h)'] = 0
            # Dimensionsspezifische Spalten
            row_dim['Brutto_Volumen'] = dim_row['Brutto_Volumen']
            row_dim['Brutto_Ausschuss'] = dim_row['Brutto_Ausschuss']
            row_dim['Netto_Volumen'] = dim_row['Netto_Volumen']
            row_dim['Brutto_Ausbeute'] = dim_row['Brutto_Ausbeute']
            row_dim['Netto_Ausbeute'] = dim_row['Netto_Ausbeute']
            row_dim['CE'] = dim_row['CE']
            row_dim['SF'] = dim_row['SF']
            row_dim['SI'] = dim_row['SI']
            row_dim['IND'] = dim_row['IND']
            row_dim['NSI'] = dim_row['NSI']
            row_dim['Q_V'] = dim_row['Q_V']
            row_dim['Ausschuss'] = dim_row['Ausschuss']
            
            final_data.append(row_dim)
    
    # Erzeugen des finalen DataFrames in der gewünschten Reihenfolge
    final_df = pd.DataFrame(final_data, columns=final_columns)
    
    st.subheader("Original-Layout mit berechneten Kennzahlen, Vorschub (FM/h) und Stärke_Klasse")
    st.dataframe(final_df)
    
    # Download als Excel
    excel_bytes = to_excel(final_df)
    st.download_button(
        label="Download als Excel (Original-Layout)",
        data=excel_bytes,
        file_name="monatsanalyse_original_format.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

if __name__ == "__main__":
    main()
