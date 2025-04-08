import streamlit as st
import pandas as pd
import numpy as np
import openpyxl


# Titel der App
st.title("Verarbeitungstool Handzählungen")

# df initialisieren
df = None

# Datei-Upload
uploaded_file = st.file_uploader("Lade eine Excel-Datei hoch", type=["xlsx", "xls"])

if uploaded_file:
    # Datei als DataFrame laden
    try:
        df = pd.read_excel(uploaded_file, sheet_name="Datenbank")  # Blatt "Datenbank" laden
        st.success("Datenbank erfolgreich geladen.")
        
    except Exception as e:
        st.error(f"Fehler beim Laden der Datei: {e}")


if df is not None:

    # Benutzereingabe für Start- und Enddatum
    st.write("Wählen den gewünschten Zeitraum aus, um den Frasy-Export zu erstellen:")
    start_date = st.date_input("Startdatum", pd.to_datetime('2025-03-31'))
    end_date = st.date_input("Enddatum", pd.to_datetime('2025-08-01'))


    # Filterung des DataFrames basierend auf den ausgewählten Daten
    df = df[(df['Datum'] >= pd.Timestamp(start_date)) & (df['Datum'] <= pd.Timestamp(end_date))]

    if start_date and end_date:
        # Mapping UIC-Nummern
        uic_nummern = {"Langnau i.E., Bahnhof": "08207",
                    "Bärau, Dorf": "08990",
                    "Trubschachen, Bahnhof": "08208",
                    "Wiggen, Egghus": "72862",
                    "Escholzmatt, Bahnhof": "08210",
                    "Schüpfheim, Bahnhof": "08211"}

        # Fülle fehlende Ein- und Aussteiger mit 0
        df.loc[:, df.columns.str.startswith(('Einsteiger', 'Aussteiger'))] = \
            df.loc[:, df.columns.str.startswith(('Einsteiger', 'Aussteiger'))].fillna(0)

        # Zählen der Einträge je Datum und Kursnummer
        count_df = df.groupby(['Datum', 'Kursnummer', "Angebot"]).size().reset_index(name='Häufigkeit')

        # Angebot gemäss Erfassungen berechnen
        count_df['berechnetes_Angebot'] = count_df['Häufigkeit'] * 90

        # Zusammengehörende Zeilen zusammenfügen
        combined_df = df.groupby(['Datum', 'Kursnummer', 'Bahnhof1', 'Bahnhof2', 'Bahnhof3', 'Bahnhof4', 'Bahnhof5', 'Bahnhof6'])[
            ['Einsteiger1', 'Einsteiger2', 'Einsteiger3', 'Einsteiger4', 'Einsteiger5', 'Einsteiger6', 
            'Aussteiger1', 'Aussteiger2', 'Aussteiger3', 'Aussteiger4', 'Aussteiger5', 'Aussteiger6']
            ].sum().reset_index()

        # Angebot und berechnetes Angebot hinzufügen
        combined_df = combined_df.merge(count_df, how="left", on=["Datum", "Kursnummer"])

        # definitives Angebot bestimmen (massgebend für Frasy-Export) und Querrechnung durchführen
        # Funktion, um das Angebot zu vergleichen und die Spalten hochzurechnen
        def compare_and_update(row):
            if row['Angebot'] > row['berechnetes_Angebot']:
                row['Angebot_def'] = row['Angebot']
                # Berechnung des Faktors für die Querrechnung
                factor = row['Angebot'] / row['berechnetes_Angebot']
                # Spalten 'Einsteiger' und 'Aussteiger' hochrechnen
                for col in df.columns:
                    if col.startswith('Einsteiger') or col.startswith('Aussteiger'):
                        row[col] *= factor
            else:
                # In allen anderen Fällen (Angebot == berechnetes_Angebot oder Angebot < berechnetes_Angebot)
                row['Angebot_def'] = row['berechnetes_Angebot']
            return row

        # Anwenden der Funktion auf jede Zeile
        combined_df = combined_df.apply(compare_and_update, axis=1)


        # leerer Frasy-Export erstellen
        columns = [
            "ZUGNUMMERSCHEMA[3]", "ZUGNUMMER_NEW[6]", "DATUM[8]", "BPNUMBER[5]", "BPUIC[2]",
            "ANABCODE[1]", "ANGEBOTCLASS2[4]", "EINCLASS2[4]", "AUSCLASS2[4]", "REISENDECLASS2[4]",
            "ANGEBOTCLASS1[4]", "EINCLASS1[4]", "AUSCLASS1[4]", "REISENDECLASS1[4]", "AUSRUSTGRADCLASS2[3]",
            "AUSRUSTGRADCLASS1[3]", "ABKURZBP[5]"
        ]

        # Erstellen des leeren DataFrames
        frasy_export = pd.DataFrame(columns=columns)

        # Iteration über das bestehende DataFrame
        rows = []
        for _, row in combined_df.iterrows():
            for i in range(1, 7):  # Für jeden Bahnhof (1-6)
                new_row = {
                    "ZUGNUMMERSCHEMA[3]": "004",  # Immer 004
                    "ZUGNUMMER_NEW[6]": row["Kursnummer"], # wird noch 6-stellig gemacht
                    "DATUM[8]": row["Datum"], # wird noch umformatiert
                    "BPNUMBER[5]": "0",  # Falls eine Zuordnung möglich ist, hier anpassen
                    "BPUIC[2]": "85",  # Immer 85
                    "ANABCODE[1]": "0",  # wird später angepasst
                    "ANGEBOTCLASS2[4]": row["Angebot_def"],  # 6x gleich
                    "EINCLASS2[4]": row[f"Einsteiger{i}"],
                    "AUSCLASS2[4]": row[f"Aussteiger{i}"],
                    "REISENDECLASS2[4]": "0",  # wird später berechnet
                    "ANGEBOTCLASS1[4]": "0000",  # 1. Klasse immer 0
                    "EINCLASS1[4]": "0000",  # 1. Klasse immer 0
                    "AUSCLASS1[4]": "0000",  # 1. Klasse immer 0
                    "REISENDECLASS1[4]": "0000",  # 1. Klasse immer 0
                    "AUSRUSTGRADCLASS2[3]": "100",  # Falls nötig, hier anpassen
                    "AUSRUSTGRADCLASS1[3]": "100",  # Falls nötig, hier anpassen
                    "ABKURZBP[5]": row[f"Bahnhof{i}"]  # Der jeweilige Bahnhof
                }
                rows.append(new_row)

        # DataFrame mit allen neuen Zeilen erstellen
        frasy_export = pd.DataFrame(rows)

        # Unnötige Zeilen löschen
        frasy_export = frasy_export[frasy_export["ABKURZBP[5]"] != 0].reset_index(drop=True)

        # Umformatierung Datums-Spalte
        frasy_export["DATUM[8]"] = frasy_export["DATUM[8]"].astype(str).str.replace("-", "")

        # Hinzufügen der BPNUMBER
        frasy_export["BPNUMBER[5]"] = frasy_export["ABKURZBP[5]"].map(uic_nummern)

        # ABKURZBP mit ZZZZ ersetzen
        frasy_export["ABKURZBP[5]"] = "ZZZZ"

        # Zugnummer 6-stellig machen
        frasy_export["ZUGNUMMER_NEW[6]"] = frasy_export["ZUGNUMMER_NEW[6]"].astype(str).str.zfill(6)

        # ANABCODE ausfüllen
        frasy_export["ANABCODE[1]"] = 2  # Standardmäßig auf 2 setzen

        # Markiere die letzte Zeile pro Kombination aus ZUGNUMMER_NEW[6] und DATUM[8]
        frasy_export.loc[frasy_export.duplicated(subset=["ZUGNUMMER_NEW[6]", "DATUM[8]"], keep="last"), "ANABCODE[1]"] = 1

        # Reduktion durch ZP-Faktor
        zp_faktor = 0.950

        frasy_export['EINCLASS2[4]'] = np.ceil(frasy_export['EINCLASS2[4]'] * zp_faktor)
        frasy_export['AUSCLASS2[4]'] = np.ceil(frasy_export['AUSCLASS2[4]'] * zp_faktor)

        # Berechnung der Besetzung
        # Initialisiere die Besetzung für die erste Zeile
        frasy_export['REISENDECLASS2[4]'] = frasy_export['EINCLASS2[4]'] - frasy_export['AUSCLASS2[4]']

        # Iteriere durch den DataFrame, beginnend ab der zweiten Zeile
        for i in range(1, len(frasy_export)):
            if frasy_export.loc[i, 'DATUM[8]'] == frasy_export.loc[i-1, 'DATUM[8]'] and frasy_export.loc[i, 'ZUGNUMMER_NEW[6]'] == frasy_export.loc[i-1, 'ZUGNUMMER_NEW[6]']:
                # Wenn Datum und Zugnummer übereinstimmen, berücksichtige die REISENDECLASS2 der vorherigen Zeile
                frasy_export.loc[i, 'REISENDECLASS2[4]'] = frasy_export.loc[i, 'EINCLASS2[4]'] - frasy_export.loc[i, 'AUSCLASS2[4]'] + frasy_export.loc[i-1, 'REISENDECLASS2[4]']
            else:
                # Andernfalls berechne die Besetzung mit den aktuellen Werten
                frasy_export.loc[i, 'REISENDECLASS2[4]'] = frasy_export.loc[i, 'EINCLASS2[4]'] - frasy_export.loc[i, 'AUSCLASS2[4]'] + frasy_export.loc[i, 'REISENDECLASS2[4]']
            
            # Bei Endpunkt des Zuges die Besetzung immer auf 0 setzen
            if frasy_export.loc[i, 'ANABCODE[1]'] == 2:
                frasy_export.loc[i, 'REISENDECLASS2[4]'] = 0
        
        
        # Angebot und Daten 4-stellig machen
        frasy_export['ANGEBOTCLASS2[4]'] = frasy_export['ANGEBOTCLASS2[4]'].astype(str).str.zfill(4)
        frasy_export['EINCLASS2[4]'] = frasy_export['EINCLASS2[4]'].astype(int).astype(str).str.zfill(4)
        frasy_export['AUSCLASS2[4]'] = frasy_export['AUSCLASS2[4]'].astype(int).astype(str).str.zfill(4)
        frasy_export['REISENDECLASS2[4]'] = frasy_export['REISENDECLASS2[4]'].astype(int).astype(str).str.zfill(4)


        # Speichert den Status des Downloads in der Session
        if "downloaded" not in st.session_state:
            st.session_state.downloaded = False

        @st.cache_data
        def convert_df_to_txt(df):
            return df.to_csv(sep=';', encoding='utf-8', header=False, index=False)

        # Dateinamen erstellen
        formatted_start_date = str(start_date).replace("-", "")
        formatted_end_date = str(end_date).replace("-", "")
        dateiname = f"BLS_EV_LN_SCHH_{formatted_start_date}_{formatted_end_date}.txt"

        # Download-Button
        txt_data = convert_df_to_txt(frasy_export)
        download_clicked = st.download_button(
            label="Download der Frasy-Daten als .txt",
            data=txt_data,
            file_name=dateiname,
            mime="text/plain"
        )
    

        # Erfolgsmeldung nach dem Download
        if download_clicked:
            st.session_state.downloaded = True

        if st.session_state.downloaded:
            st.success("✅ Datei wurde erfolgreich erstellt und heruntergeladen!")
            st.balloons()
