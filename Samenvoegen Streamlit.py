import streamlit as st
from openpyxl import load_workbook
from copy import copy  # ✅ Gebruik de standaard copy functie
import io

st.title("📑 Excel Samenvoeger (met layout behoud)")

st.write("Upload één of meerdere `.xlsx` bestanden om samen te voegen. Layout en stijlen blijven behouden.")

uploaded_files = st.file_uploader(
    "Selecteer Excel-bestanden", type="xlsx", accept_multiple_files=True
)

if uploaded_files:
    # Laad het eerste bestand als basis
    first_file = uploaded_files[0]
    base_wb = load_workbook(filename=io.BytesIO(first_file.read()))
    base_ws = base_wb.active

    for file in uploaded_files[1:]:
        file.seek(0)  # Zorg dat we opnieuw kunnen lezen
        wb_to_merge = load_workbook(filename=io.BytesIO(file.read()))
        ws_to_merge = wb_to_merge.active

        # Vind eerste lege rij in base_ws
        if base_ws.max_row == 1 and all([cell.value is None for cell in base_ws[1]]):
            start_row = 1
        else:
            start_row = base_ws.max_row + 1

        for i, row in enumerate(ws_to_merge.iter_rows(values_only=False), start=start_row):
            for j, cell in enumerate(row, start=1):
                new_cell = base_ws.cell(row=i, column=j, value=cell.value)
                # Kopieer alle stijlen
               if cell.has_style:
                new_cell.font = copy(cell.font)
                new_cell.fill = copy(cell.fill)
                new_cell.border = copy(cell.border)
                new_cell.alignment = copy(cell.alignment)
                new_cell.number_format = copy(cell.number_format)
                new_cell.protection = copy(cell.protection)

        
        # Kopieer kolombreedtes
        for col_letter, col_dim in ws_to_merge.column_dimensions.items():
            base_ws.column_dimensions[col_letter].width = col_dim.width

        # Kopieer rijhoogtes
        for row_dim in ws_to_merge.row_dimensions:
            if ws_to_merge.row_dimensions[row_dim].height:
                base_ws.row_dimensions[row_dim + start_row - 1].height = ws_to_merge.row_dimensions[row_dim].height

    # Sla het samengevoegde bestand op
    output = io.BytesIO()
    base_wb.save(output)
    output.seek(0)

    st.download_button(
        label="📥 Download Samengevoegd Excel-bestand",
        data=output,
        file_name="samengevoegd.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

