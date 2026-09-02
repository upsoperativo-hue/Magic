import streamlit as st
import pandas as pd
from io import BytesIO
from datetime import datetime
from openpyxl import load_workbook
import zipfile
import xml.etree.ElementTree as ET

# CONFIG
TEMPLATE_PATH = "TOOL_template.xlsm"
SHEET_NAME = "Trackingnummern"
START_ROW = 7
COLUMN = "A"

st.set_page_config(page_title="TOOL ELEGGIBILI", page_icon="📦")
st.title("TOOL ELEGGIBILI – Weekend Eligible")

st.write("Carica uno o più file ULD Detail (.xlsx). Il tool estrarrà i tracking con WeekendEligibleVolume = 1 e li inserirà nel template TOOL (.xlsm) mantenendo le macro.")

uploaded_files = st.file_uploader(
    "Carica i file ULD Detail",
    type=["xlsx"],
    accept_multiple_files=True
)

def replace_sheet_in_xlsm(original_xlsm, modified_xlsx, sheet_name):
    """
    Sostituisce il foglio XML modificato nel file XLSM originale,
    preservando macro e struttura.
    """

    # 1) Leggi il file XLSM originale come ZIP
    with zipfile.ZipFile(original_xlsm, "r") as zin:
        original_files = zin.namelist()
        zip_content = {name: zin.read(name) for name in original_files}

    # 2) Leggi il file XLSX modificato (solo foglio)
    with zipfile.ZipFile(modified_xlsx, "r") as zin2:
        modified_files = zin2.namelist()

        # Trova il foglio corrispondente
        sheet_map = {}
        workbook_xml = zin2.read("xl/workbook.xml")
        root = ET.fromstring(workbook_xml)

        # Namespace
        ns = {"ns": "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}

        sheets = root.find("ns:sheets", ns)
        for s in sheets.findall("ns:sheet", ns):
            sheet_map[s.attrib["name"]] = s.attrib[
                "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id"
            ]

        # Trova il file XML del foglio
        rels_xml = zin2.read("xl/_rels/workbook.xml.rels")
        rels_root = ET.fromstring(rels_xml)
        rels_ns = {"r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships"}

        target_sheet = None
        for rel in rels_root.findall("r:Relationship", rels_ns):
            if rel.attrib["Id"] == sheet_map[sheet_name]:
                target_sheet = "xl/" + rel.attrib["Target"]

        if target_sheet is None:
            raise ValueError("Foglio non trovato nel file XLSX modificato.")

        modified_sheet_xml = zin2.read(target_sheet)

    # 3) Ricrea il file XLSM sostituendo solo il foglio
    output = BytesIO()
    with zipfile.ZipFile(output, "w", zipfile.ZIP_DEFLATED) as zout:
        for name, data in zip_content.items():
            if name == target_sheet:
                zout.writestr(name, modified_sheet_xml)
            else:
                zout.writestr(name, data)

    output.seek(0)
    return output


if uploaded_files:
    if st.button("Elabora e genera TOOL ELEGGIBILI"):
        all_tracks = []

        # 1) Estrazione tracking da tutti i file
        for file in uploaded_files:
            df = pd.read_excel(file, header=5)  # riga 6 come intestazione

            missing = []
            if "WeekendEligibleVolume" not in df.columns:
                missing.append("WeekendEligibleVolume")
            if "TrackingNumber" not in df.columns:
                missing.append("TrackingNumber")

            if missing:
                st.warning(f"File {file.name} ignorato: colonne mancanti {', '.join(missing)}.")
                continue

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

        # 2) Carica il TEMPLATE XLSM
        try:
            wb = load_workbook(TEMPLATE_PATH, keep_vba=True)
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

        # 4) Scrivi i tracking
        row = START_ROW
        for t in unique_tracks:
            ws[f"{COLUMN}{row}"] = t
            row += 1

        # 5) Salva XLSX temporaneo
        temp_xlsx = BytesIO()
        wb.save(temp_xlsx)
        temp_xlsx.seek(0)

        # 6) Ricostruisci XLSM preservando macro
        final_file = replace_sheet_in_xlsm(TEMPLATE_PATH, temp_xlsx, SHEET_NAME)

        # 7) Nome file
        today_str = datetime.now().strftime("%d.%m.%Y")
        filename = f"TOOL ELEGGIBILI {today_str}.xlsm"

        # 8) Download
        st.download_button(
            "Scarica TOOL ELEGGIBILI",
            data=final_file.getvalue(),
            file_name=filename,
            mime="application/vnd.ms-excel.sheet.macroEnabled.12"
        )
