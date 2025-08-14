import streamlit as st
import openpyxl
from io import BytesIO

st.title("Excel Samenvoegen met Doorlopende Nummering")

uploaded_files = st.file_uploader("Selecteer Excel-bestanden om samen te voegen", accept_multiple_files=True)

if uploaded_files:
    # Laad het eerste bestand als basis
    base_wb = openpyxl.load_workbook(uploaded_files[0])
    base_ws = base_wb.active

# Bepaal de laatste waarde in kolom A
last_number = 0
for row in base_ws.iter_rows(min_row=2, max_col=1):  # ga uit van header in rij 1
    try:
        value = int(row[0].value)
        if value > last_number:
            last_number = value
    except (TypeError, ValueError):
        continue  # sla lege of niet-numerieke cellen over


    # Voeg de rest van de bestanden toe
    for f in uploaded_files[1:]:
        wb = openpyxl.load_workbook(f)
        ws = wb.active

        for row in ws.iter_rows(min_row=2):  # ga uit van header in rij 1
            last_number += 1  # tel door
            new_row = [last_number] + [cell.value for cell in row[1:]]  # behoud alles behalve kolom A
            base_ws.append(new_row)

    # Opslaan
    output = BytesIO()
    base_wb.save(output)
    output.seek(0)

    st.download_button(
        label="Download samengevoegd bestand",
        data=output,
        file_name="samengevoegd.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

