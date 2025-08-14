import streamlit as st
from openpyxl import load_workbook
from copy import copy
from io import BytesIO

st.title("Excel Samenvoeger met Doorlopende Nummering")

uploaded_files = st.file_uploader(
    "Selecteer één of meerdere Excel-bestanden (.xlsx)",
    type="xlsx",
    accept_multiple_files=True
)

if uploaded_files:
    # Laad het eerste bestand als basis
    base_file = uploaded_files[0]
    base_wb = load_workbook(filename=BytesIO(base_file.read()))
    base_ws = base_wb.active

    # Bepaal de laatste waarde in kolom A
    last_number = 0
    for row in base_ws.iter_rows(min_row=2, max_col=1):  # header in rij 1
        try:
            value = int(row[0].value)
            if value > last_number:
                last_number = value
        except (TypeError, ValueError):
            continue

    # Voeg de overige bestanden toe
    for file in uploaded_files[1:]:
        wb = load_workbook(filename=BytesIO(file.read()))
        ws = wb.active

        for i, row in enumerate(ws.iter_rows(min_row=2), start=1):
            for j, cell in enumerate(row):
                new_cell = base_ws.cell(row=base_ws.max_row + 1, column=j + 1, value=cell.value)
                if cell.has_style:
                    new_cell.font = copy(cell.font)
                    new_cell.fill = copy(cell.fill)
                    new_cell.border = copy(cell.border)
                    new_cell.alignment = copy(cell.alignment)
                    new_cell.number_format = copy(cell.number_format)
                    new_cell.protection = copy(cell.protection)
            # Update kolom A met doorlopende nummering
            base_ws.cell(row=base_ws.max_row, column=1, value=last_number + i)
        last_number += ws.max_row - 1  # update last_number voor het volgende bestand

    # Opslaan in memory en downloadknop tonen
    output = BytesIO()
    base_wb.save(output)
    st.download_button(
        "Download samengevoegd bestand",
        data=output.getvalue(),
        file_name="samengevoegd.xlsx"
    )
