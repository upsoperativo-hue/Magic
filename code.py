import streamlit as st
import pandas as pd
from io import BytesIO
from datetime import datetime
from openpyxl import load_workbook

# CONFIG
TEMPLATE_PATH = "TOOL_template.xlsm"   # Template Excel nel repository
SHEET_NAME = "Trackingnummern"
START_ROW = 7
COLUMN = "A"

st.set_page_config(page_title="TOOL ELEGGIBILI", page_icon="📦")
st.title("TOOL ELEGGIBILI – Weekend Eligible")

st.write("Carica uno o più file ULD Detail (.xlsx). Il tool estrarrà i tracking con WeekendEligibleVolume = 1 e li inserirà nel template TOOL.")

uploaded_files = st.file_uploader(
    "Carica i file ULD Detail",
    type=["xlsx"],
    accept_multiple_files=True
)

if uploaded_files:
    if st.button("Elabora e genera TOOL ELEGGIBILI"):
        all_tracks = []

        # 1) Estrazione tracking da tutti i file
        for file in uploaded_files:
            df = pd.read_excel(file, header=5)

            if "WeekendEligibleVolume" not in df.columns or "TrackingNumber" not in df.columns:
                st.error(f"File {file.name}: colonne 'WeekendEligibleVolume' o 'TrackingNumber' mancanti.")
                st.stop()

            df_filtered = df[df["WeekendEligibleVolume"] == 1]
            tracks = df_filtered["TrackingNumber"].dropna().astype(str).tolist()
            all_tracks.extend(tracks)

        # Rimuovi duplicati mantenendo l'ordine
        seen = set()
        unique_tracks = []
        for t in all_tracks:
            if t not in seen:
                seen.add(t)
                unique_tracks.append(t)

        if not unique_tracks:
            st.warning("Nessun tracking con WeekendEligibleVolume = 1 trovato nei file caricati.")
            st.stop()

        st.success(f"Trovati {len(unique_tracks)} tracking eleggibili.")

        # 2) Carica il TEMPLATE TOOL
        try:
            wb = load_workbook(TEMPLATE_PATH)
        except Exception as e:
            st.error(f"Impossibile aprire il template TOOL: {e}")
            st.stop()

        if SHEET_NAME not in wb.sheetnames:
            st.error(f"Nel template non esiste il foglio '{SHEET_NAME}'.")
            st.stop()

        ws = wb[SHEET_NAME]

        # 3) Pulisci colonna A da riga 7 in giù
        max_row = ws.max_row
        for r in range(START_ROW, max_row + 1):
            ws[f"{COLUMN}{r}"] = None

        # 4) Scrivi i tracking in colonna A a partire da riga 7
        row = START_ROW
        for t in unique_tracks:
            ws[f"{COLUMN}{row}"] = t
            row += 1

        # 5) Salva in memoria con nome TOOL ELEGGIBILI dd.mm.yyyy
        buffer = BytesIO()
        wb.save(buffer)
        buffer.seek(0)

        today_str = datetime.now().strftime("%d.%m.%Y")
        filename = f"TOOL ELEGGIBILI {today_str}.xlsm"

        st.download_button(
            "Scarica TOOL ELEGGIBILI",
            data=buffer.getvalue(),
            file_name=filename,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
