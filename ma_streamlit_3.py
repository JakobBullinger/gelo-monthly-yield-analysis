import streamlit as st
import pandas as pd
import re
import io

def to_excel(df):
    """
    Hilfsfunktion, um ein pandas DataFrame nach Excel (in Memory) zu schreiben.
    Gibt einen Byte-Stream zurück, der für den Download-Button genutzt werden kann.
    """
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df.to_excel(writer, index=False, sheet_name='Monatsanalyse')
    processed_data = output.getvalue()
    return processed_data

def main():
    st.title("Monatsanalyse Ausbeute (pro Auftrag & Dimension)")
    
    st.markdown("""
    **Vorgehen:**
    1. Lade alle Tages-Excel-Dateien eines Monats hoch.
    2. Das Tool fasst alle Dateien zu einem DataFrame zusammen.
    3. Es trennt die Gesamtzeilen (erste Zeile pro Auftrag) von den Dimensionszeilen.
    4. Pro (Auftragsnummer, Dimension) werden die Kennzahlen summiert.
    5. Die Gesamtwerte werden mit den Dimensionswerten zusammengeführt.
    6. Das Ergebnis kannst du als Excel-Datei herunterladen.
    """)

    # 1) Upload
    uploaded_files = st.file_uploader(
        "Bitte alle Excel-Dateien hochladen", 
        accept_multiple_files=True, 
        type=["xlsx", "xls"]
    )
    if not uploaded_files:
        st.info("Bitte Dateien auswählen.")
        return
    
    # 2) Einlesen & Zusammenführen
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
    
    # 3) Auftragsnummer extrahieren (z.B. erste 5 Ziffern)
    df_all['Auftragsnummer'] = df_all['Auftrag'].astype(str).apply(
        lambda x: re.findall(r'^\d{5}', x)[0] if re.findall(r'^\d{5}', x) else x
    )
    
    # 4) Trennung Overall (erste Zeile pro Auftrag) vs. Dimensionszeilen
    #    Kriterium: Overall, wenn "Stämme" != 0
    df_overall = df_all[df_all['Stämme'] != 0].copy()
    df_dim = df_all[df_all['Stämme'] == 0].copy()
    
    # 5) Aggregation pro Auftrag & Dimension
    #    Wir summieren hier auch "Teile", da es in den Dimensionszeilen befüllt ist
    dim_cols = [
        'Teile',            # weil du sagst, dass die Teile auch pro Dimension gefüllt sind
        'Brutto_Volumen',
        'Netto_Volumen',
        'CE',
        'SF',
        'SI',
        'IND',
        'NSI',
        'Q_V',
        'Ausschuss'
    ]
    grouped_dim = df_dim.groupby(['Auftragsnummer', 'Dimension'], as_index=False)[dim_cols].sum()
    
    #    Für bessere Übersicht: umbenennen "Teile" => "Teile_dim"
    grouped_dim.rename(columns={'Teile': 'Teile_dim'}, inplace=True)
    
    # 6) Die "Overall"-Tabelle enthält die Gesamtwerte
    #    (Wir behalten dort "Teile" als Gesamt-Teile => "Teile_gesamt")
    overall_cols = [
        'Auftragsnummer', 
        'Auftrag', 
        'Stämme', 
        'Volumen_Eingang', 
        'Durchschn_Stämme', 
        'Teile',            # später "Teile_gesamt"
        'Durchmesser', 
        'Laufzeit_Minuten'
    ]
    df_overall_unique = (
        df_overall[overall_cols]
        .drop_duplicates(subset=['Auftragsnummer'])
        .rename(columns={'Teile': 'Teile_gesamt'})
    )
    
    # 7) Merge (Zusammenführen) pro Auftragsnummer
    merged = pd.merge(grouped_dim, df_overall_unique, on='Auftragsnummer', how='left')
    
    # 8) Spaltenreihenfolge festlegen
    ordered_columns = [
        'Auftrag', 
        'Auftragsnummer', 
        'Dimension',
        'Stämme', 
        'Volumen_Eingang', 
        'Durchschn_Stämme', 
        'Durchmesser', 
        'Laufzeit_Minuten',
        'Teile_gesamt',     # Gesamt-Teile
        'Teile_dim',        # Teile für genau diese Dimension
        'Brutto_Volumen', 
        'Netto_Volumen', 
        'CE', 
        'SF', 
        'SI', 
        'IND', 
        'NSI', 
        'Q_V', 
        'Ausschuss'
    ]
    # Nur Spalten nehmen, die auch wirklich existieren (falls mal eine fehlt in den Daten)
    final_cols = [c for c in ordered_columns if c in merged.columns]
    merged = merged[final_cols]
    
    # 9) Darstellung in Streamlit
    st.subheader("Ergebnis: Monatsanalyse pro Auftrag & Dimension")
    st.dataframe(merged)
    
    # 10) Download als Excel
    excel_data = to_excel(merged)
    st.download_button(
        label="Download als Excel",
        data=excel_data,
        file_name="monatsanalyse_ausbeute.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

if __name__ == "__main__":
    main()
