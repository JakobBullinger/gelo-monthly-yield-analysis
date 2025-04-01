import streamlit as st
import pandas as pd
import io

st.title("Monatliche Ausbeuteanalyse")

# Debug-Modus aktivieren
debug_mode = st.checkbox("Debug Mode aktivieren", value=False)

# --- Schritt 1: Referenztabelle hochladen ---
st.subheader("Schritt 1: Referenztabelle hochladen")
dimension_file = st.file_uploader(
    "Bitte lade die Excel-Datei mit der Referenztabelle hoch (zwei Spalten, z.B. 'Dim1' und 'Dim2')",
    type=["xlsx", "xls"]
)

# --- Schritt 2: Tages-Excel-Dateien hochladen ---
st.subheader("Schritt 2: Tages-Excel-Dateien eines Monats hochladen")
daily_files = st.file_uploader(
    "Wähle mehrere Tagesdateien aus (alle haben dasselbe Format)",
    type=["xlsx", "xls"],
    accept_multiple_files=True
)

# --- Funktion: Rundholz_FM pro Auftrag berechnen (mit Debug-Ausgaben) ---
def assign_rundholz(group, debug=False):
    # Ersetze NaN in der Spalte "Dimension" durch leere Strings und trimme Leerzeichen
    group["Dimension"] = group["Dimension"].fillna("").astype(str).str.strip()
    
    if debug:
        st.write("=== Auftrag:", group["Auftrag"].iloc[0], "===")
        st.write("Gruppe vor Rundholz-FM-Zuweisung:")
        st.write(group)
    
    # Finde Zeilen, in denen "Dimension" leer ist (Gesamtzeile)
    empty_rows = group[group["Dimension"] == ""]
    if not empty_rows.empty:
        # Nehme den Wert aus "Volumen_Eingang" der ersten leeren Zeile
        total_rundholz = empty_rows.iloc[0]["Volumen_Eingang"]
    else:
        total_rundholz = 0

    # Finde alle Zeilen mit nicht-leerer "Dimension"
    non_empty = group[group["Dimension"] != ""]
    if not non_empty.empty:
        # Hauptware: letzte Zeile der Gruppe
        hauptware_index = non_empty.index[-1]
        group.loc[hauptware_index, "Rundholz_FM"] = total_rundholz
        # Alle anderen Zeilen in der Gruppe bekommen 0
        group.loc[non_empty.index[:-1], "Rundholz_FM"] = 0
    else:
        group["Rundholz_FM"] = 0
    
    if debug:
        st.write("Gruppe nach Rundholz-FM-Zuweisung:")
        st.write(group)
    
    return group

# --- Button, um die Verarbeitung zu starten ---
if st.button("Auswertung starten"):
    if dimension_file is None:
        st.warning("Bitte zuerst die Referenztabelle hochladen.")
        st.stop()
    if not daily_files:
        st.warning("Bitte mindestens eine Tagesdatei hochladen.")
        st.stop()

    # --- Referenztabelle einlesen und verarbeiten ---
    try:
        ref_df = pd.read_excel(dimension_file)
    except Exception as e:
        st.error(f"Fehler beim Einlesen der Referenztabelle: {e}")
        st.stop()

    # Annahme: Referenztabelle hat zwei Spalten, z.B. "Dim1" und "Dim2" (Werte wie "75,00")
    ref_df["Dim1"] = (
        ref_df.iloc[:, 0]
        .astype(str)
        .str.replace(",", ".")
        .astype(float)
        .astype(int)
        .astype(str)
    )
    ref_df["Dim2"] = (
        ref_df.iloc[:, 1]
        .astype(str)
        .str.replace(",", ".")
        .astype(float)
        .astype(int)
        .astype(str)
    )
    ref_df["DimensionKey"] = ref_df["Dim1"] + "x" + ref_df["Dim2"]
    # Generiere einen SortIndex (falls nicht vorhanden)
    ref_df.reset_index(inplace=True)
    ref_df.rename(columns={"index": "SortIndex"}, inplace=True)
    ref_df["SortIndex"] = ref_df["SortIndex"] + 1

    # --- Tagesdateien einlesen und zusammenführen ---
    all_dfs = []
    for f in daily_files:
        try:
            df_temp = pd.read_excel(f)
            all_dfs.append(df_temp)
        except Exception as e:
            st.error(f"Fehler beim Einlesen der Datei {f.name}: {e}")

    if not all_dfs:
        st.warning("Es konnten keine Tagesdateien eingelesen werden.")
        st.stop()

    daily_df = pd.concat(all_dfs, ignore_index=True)

    # --- Anpassung der Tagesdaten ---
    # Wir nehmen an, dass in den Tagesdaten bereits eine Spalte "Dimension" existiert.
    daily_df["Dimension"] = daily_df["Dimension"].fillna("").astype(str).str.strip()
    # Nehme an, die Spalte "Dimension" liefert bereits den richtigen Schlüssel, z.B. "43x141"
    daily_df["DimensionKey"] = daily_df["Dimension"]

    # --- Rundholz_FM pro Auftrag berechnen ---
    # Wende die Funktion auf jede Gruppe (nach "Auftrag") an, ggf. mit Debug-Ausgaben
    daily_df = daily_df.groupby("Auftrag", group_keys=False).apply(lambda g: assign_rundholz(g, debug=debug_mode))

    # --- Aggregation der Tagesdaten ---
    # Aggregiere die Kennzahlen pro DimensionKey
    agg_df = daily_df.groupby("DimensionKey", as_index=False).agg({
        "Volumen_Ausgang": "sum",
        "Brutto_Volumen": "sum",
        "Brutto_Ausschuss": "sum",
        "Netto_Volumen": "sum",
        "Brutto_Ausbeute": "sum",
        "Netto_Ausbeute": "sum",
        "CE": "sum",
        "SF": "sum",
        "SI": "sum",
        "IND": "sum",
        "NSI": "sum",
        "Q_V": "sum",
        "Ausschuss": "sum",
        "Rundholz_FM": "sum"  # neue Kennzahl
    })

    # --- Left Join: Alle Dimensionen aus der Referenztabelle übernehmen ---
    final_df = pd.merge(
        ref_df,    # Referenztabelle mit SortIndex und DimensionKey
        agg_df,    # aggregierte Tagesdaten
        on="DimensionKey",
        how="left"
    )

    # Ersetze NaN in den numerischen Kennzahl-Spalten durch 0
    numeric_cols = [
        "Volumen_Ausgang", "Brutto_Volumen", "Brutto_Ausschuss",
        "Netto_Volumen", "Brutto_Ausbeute", "Netto_Ausbeute",
        "CE", "SF", "SI", "IND", "NSI", "Q_V", "Ausschuss", "Rundholz_FM"
    ]
    for col in numeric_cols:
        if col in final_df.columns:
            final_df[col] = final_df[col].fillna(0)

    # --- Sortieren und Spaltenreihenfolge anpassen ---
    final_df.sort_values("SortIndex", inplace=True)
    cols_order = ["SortIndex", "DimensionKey"] + numeric_cols
    final_df = final_df[cols_order]

    # --- Ergebnis anzeigen ---
    st.subheader("Ergebnis-Tabelle")
    st.dataframe(final_df)

    # --- Excel-Export vorbereiten ---
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        final_df.to_excel(writer, index=False, sheet_name="Ergebnis")
    output.seek(0)

    st.download_button(
        label="Ergebnis als Excel herunterladen",
        data=output,
        file_name="Ausbeuteanalyse_Ergebnis.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
