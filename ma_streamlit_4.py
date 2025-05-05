import streamlit as st
import pandas as pd
import re
import io

def to_excel(df):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df.to_excel(writer, index=False, sheet_name='Monatsanalyse')
    return output.getvalue()

def main():
    st.title("Monatsanalyse Ausbeute – Original-Layout (Gesamtzeile + Dimensionszeilen)")
    
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
    
    # (4) Aggregation Dimensionen (vereinfacht):
    dim_cols = [
        'Teile', 'Brutto_Volumen', 'Netto_Volumen',
        'CE', 'SF', 'SI', 'IND', 'NSI', 'Q_V', 'Ausschuss'
    ]
    grouped_dim = df_dim.groupby(['Auftragsnummer', 'Dimension'], as_index=False)[dim_cols].sum()
    grouped_dim.rename(columns={'Teile': 'Teile_dim'}, inplace=True)
    
    # (5) Holen der Gesamtwerte (erste Zeile pro Auftrag)
    overall_cols = [
        'Auftragsnummer', 'Auftrag', 'Stämme', 'Volumen_Eingang',
        'Durchschn_Stämme', 'Teile', 'Durchmesser', 'Laufzeit_Minuten'
    ]
    df_overall_unique = (
        df_overall[overall_cols]
        .drop_duplicates(subset=['Auftragsnummer'])
        .rename(columns={'Teile': 'Teile_gesamt'})
    )
    
    # (6) Mergen => dimensionale Werte + Gesamtwerte
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
    
    # --- JETZT: REKONSTRUKTION DES ORIGINAL-FORMATS ---
    
    # Diese Spalten sollen in der Endausgabe erscheinen:
    final_columns = [
        'Auftrag',           # Spalte 1
        'Dimension',         # Spalte 2
        'Stämme',            # Spalte 3
        'Volumen_Eingang',   # Spalte 4
        'Durchschn_Stämme',  # Spalte 5
        'Teile',             # Spalte 6
        'Durchmesser',       # Spalte 7
        'Laufzeit_Minuten',  # Spalte 8
        'Brutto_Volumen',    # Spalte 9
        'Brutto_Ausschuss',  # Spalte 10 (in %)
        'Netto_Volumen',     # Spalte 11
        'Brutto_Ausbeute',   # Spalte 12 (in %)
        'Netto_Ausbeute',    # Spalte 13 (in %)
        'CE',                # Spalte 14
        'SF',                # Spalte 15
        'SI',                # Spalte 16
        'IND',               # Spalte 17
        'NSI',               # Spalte 18
        'Q_V',               # Spalte 19
        'Ausschuss'          # Spalte 20 (m³)
    ]
    
    # Wir bauen nun manuell die Zeilen zusammen:
    final_data = []
    
    # Sortiere ggf. nach Auftragsnummer, damit die Reihenfolge konsistent ist
    merged.sort_values(by=['Auftragsnummer', 'Dimension'], inplace=True)
    
    # Liste aller Auftragsnummern
    auftragsnummern = merged['Auftragsnummer'].unique()
    
    for nr in auftragsnummern:
        # 1) Gesamtzeile (erste Zeile)
        #    -> entnehmen wir df_overall_unique
        row_overall_src = df_overall_unique.loc[df_overall_unique['Auftragsnummer'] == nr].iloc[0]
        
        # Dictionary für die Gesamtzeile
        row_overall = {}
        row_overall['Auftrag'] = row_overall_src['Auftrag']
        row_overall['Dimension'] = ""  # Im Originalformat bleibt diese Zeile leer
        row_overall['Stämme'] = row_overall_src['Stämme']
        row_overall['Volumen_Eingang'] = row_overall_src['Volumen_Eingang']
        row_overall['Durchschn_Stämme'] = row_overall_src['Durchschn_Stämme']
        row_overall['Teile'] = row_overall_src['Teile_gesamt']
        row_overall['Durchmesser'] = row_overall_src['Durchmesser']
        row_overall['Laufzeit_Minuten'] = row_overall_src['Laufzeit_Minuten']
        
        # Die Dimensionsspalten in der Gesamtzeile = 0 (oder leer)
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
        
        # 2) Dimensionzeilen
        sub = merged.loc[merged['Auftragsnummer'] == nr].copy()
        
        for _, dim_row in sub.iterrows():
            row_dim = {}
            row_dim['Auftrag'] = dim_row['Auftrag']
            row_dim['Dimension'] = dim_row['Dimension']  # z.B. "17x100"
            
            # Die "Gesamtspalten" auf 0
            row_dim['Stämme'] = 0
            row_dim['Volumen_Eingang'] = 0
            row_dim['Durchschn_Stämme'] = 0
            row_dim['Teile'] = dim_row['Teile_dim']  # Dimension-Teile
            row_dim['Durchmesser'] = 0
            row_dim['Laufzeit_Minuten'] = 0
            
            # Die dimensionsspezifischen Spalten aus 'merged'
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
    
    # DataFrame in der gewünschten Spaltenreihenfolge erzeugen
    final_df = pd.DataFrame(final_data, columns=final_columns)
    
    # Anzeige in Streamlit
    st.subheader("Original-Layout mit berechneten Kennzahlen")
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
