import streamlit as st
import pandas as pd
import io

st.title("Monatliche Ausbeuteanalyse")

# 1. File Uploader für Referenztabelle (Dimensionen)
st.subheader("Schritt 1: Referenztabelle hochladen")
dimension_file = st.file_uploader(
    "Bitte lade die Excel-Datei mit den Dimensionen hoch (zwei Spalten, z.B. Dim1 und Dim2)",
    type=["xlsx", "xls"]
)

# 2. File Uploader für mehrere Tages-Exceldateien
st.subheader("Schritt 2: Tages-Exceldateien (z.B. eines Monats) hochladen")
daily_files = st.file_uploader(
    "Mehrere Dateien auswählen oder nacheinander hochladen",
    type=["xlsx", "xls"],
    accept_multiple_files=True
)

# Button, um die Verarbeitung zu starten
if st.button("Auswertung starten"):
    if dimension_file is None:
        st.warning("Bitte zuerst die Referenztabelle hochladen.")
        st.stop()
    if not daily_files:
        st.warning("Bitte mindestens eine Tagesdatei hochladen.")
        st.stop()

    # 3. Referenztabelle einlesen
    try:
        ref_df = pd.read_excel(dimension_file)
    except Exception as e:
        st.error(f"Fehler beim Einlesen der Referenztabelle: {e}")
        st.stop()

    # Beispiel: Wir nehmen an, dass Deine Referenztabelle exakt 2 Spalten hat,
    # z.B. "Dim1" und "Dim2" (numerische Werte). Passe das ggf. an.
    # Wenn sie anders heißen, z.B. "Spalte1", "Spalte2", musst Du sie umbenennen:
    # ref_df.columns = ["Dim1", "Dim2"]

    # Erzeuge einen Dimension-Key, z.B. "75.00x75.00" (als Text).
    # Oder Du lässt die Spalten separat - je nach Bedarf.
    ref_df["Dim1"] = ref_df.iloc[:, 0].astype(str).str.strip()
    ref_df["Dim2"] = ref_df.iloc[:, 1].astype(str).str.strip()
    ref_df["DimensionKey"] = ref_df["Dim1"] + "x" + ref_df["Dim2"]

    # 4. Tagesdateien einlesen und zusammenführen
    all_dfs = []
    for f in daily_files:
        try:
            # Lies jede Tagesdatei ein (Kopfzeilen anpassen, falls nötig)
            df_temp = pd.read_excel(f)

            # Beispiel: Wir gehen davon aus, dass es eine Spalte "Dimension" gibt,
            # in der sowas wie "75x75" steht. 
            # Falls die Spalte anders heißt, anpassen:
            # z.B. df_temp.rename(columns={"Dimension": "DimensionOld"}, inplace=True)

            # Falls Du Header weglassen willst, könntest Du df_temp = pd.read_excel(f, header=None)
            # und dann die Spalten manuell benennen.

            # Füge die Datei zur Liste hinzu
            all_dfs.append(df_temp)
        except Exception as e:
            st.error(f"Fehler beim Einlesen einer Tagesdatei ({f.name}): {e}")

    if not all_dfs:
        st.warning("Keine Daten eingelesen. Überprüfe die Dateien.")
        st.stop()

    # Alle Tagesdaten zu einem großen DataFrame zusammenfügen
    daily_df = pd.concat(all_dfs, ignore_index=True)

    # 5. Dimensionen in Tagesdaten ggf. anpassen
    # Beispiel: Wir nehmen an, in "daily_df" gibt es eine Spalte "Dimension",
    # in der "75x75" oder "76x96" steht. Wir erzeugen ebenfalls einen Key:
    daily_df["Dimension"] = daily_df["Dimension"].astype(str).str.strip()

    # Optional: Falls "Dimension" anders formatiert ist (z.B. "75,00 x 95,00"),
    # müsstest Du das angleichen. 
    # Hier ein Beispiel, falls Du sie splitten willst:
    # daily_df["Dim1"] = daily_df["Dimension"].str.split("x").str[0].str.strip()
    # daily_df["Dim2"] = daily_df["Dimension"].str.split("x").str[1].str.strip()
    # daily_df["DimensionKey"] = daily_df["Dim1"] + "x" + daily_df["Dim2"]

    # Für dieses Beispiel gehen wir davon aus, daily_df["Dimension"] ist schon "75x75" etc.
    daily_df["DimensionKey"] = daily_df["Dimension"]

    # 6. Aggregation der Tagesdaten (Gruppierung)
    # Beispiel: Summiere Volumen_Ausgang, Brutto_Volumen, etc.
    # Passen Sie die Spaltennamen an Dein konkretes Layout an.
    agg_df = daily_df.groupby("DimensionKey", as_index=False).agg({
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
        "Ausschuss": "sum"
    })

    # 7. Left Join: Alle Dimensionen aus ref_df + passende Kennzahlen aus agg_df
    final_df = pd.merge(
        ref_df,
        agg_df,
        on="DimensionKey",
        how="left"
    )

    # Jetzt hast Du in final_df alle Dimensionen (Sortiert in ref_df),
    # und die zugehörigen Kennzahlen (falls vorhanden), sonst NaN.

    # Optional: Ersetze NaN durch 0 in den Kennzahl-Spalten:
    numeric_cols = [
        "Brutto_Volumen", "Brutto_Ausschuss", 
        "Netto_Volumen", "Brutto_Ausbeute", "Netto_Ausbeute",
        "CE", "SF", "SI", "IND", "NSI", "Q_V", "Ausschuss"
    ]
    for col in numeric_cols:
        if col in final_df.columns:
            final_df[col] = final_df[col].fillna(0)

    # 8. Ergebnis anzeigen
    st.subheader("Ergebnis-Tabelle")
    st.dataframe(final_df)

    # 9. Download-Button für das finale Excel
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        final_df.to_excel(writer, index=False, sheet_name="Ergebnis")

    st.download_button(
        label="Ergebnis als Excel herunterladen",
        data=output.getvalue(),
        file_name="Monatsausbeute_Ausbeuteanalyse_Ergebnis.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
