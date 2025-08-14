import streamlit as st
import openpyxl
from io import BytesIO

st.title("📊 Excel Samenvoeger")

st.markdown("""
Upload hier meerdere `.xlsx` bestanden. De inhoud van elk volgend bestand wordt **onder de vorige** geplakt.
""")

# Bestanden uploaden
uploaded_files = st.file_uploader(
    "Selecteer één of meerdere Excel-bestanden",
    type="xlsx",
    accept_multiple_files=True
)

if uploaded_files:
    # Start met het eerste bestand als basis
    first_file = uploaded_files[0]
    wb_base = openpyxl.load_workbook(first_file)
    ws_base = wb_base.active

    # Loop door de rest van de bestanden
    for f in uploaded_files[1:]:
        wb_new = openpyxl.load_workbook(f)
        ws_new = wb_new.active

        # Vind de eerste lege rij in het basisbestand
        first_empty_row = ws_base.max_row + 1

        # Kopieer elke rij van ws_new naar ws_base
        for row_idx, row in enumerate(ws_new.iter_rows(values_only=False), start=first_empty_row):
            for col_idx, cell in enumerate(row, start=1):
                new_cell = ws_base.cell(row=row_idx, column=col_idx, value=cell.value)
                
                # Optioneel: behoud stijl van cellen
                if cell.has_style:
                    new_cell.font = cell.font
                    new_cell.fill = cell.fill
                    new_cell.border = cell.border
                    new_cell.alignment = cell.alignment
                    new_cell.number_format = cell.number_format
                    new_cell.protection = cell.protection

    # Opslaan naar een BytesIO object
    output = BytesIO()
    wb_base.save(output)
    output.seek(0)

    st.success("✅ Bestanden succesvol samengevoegd!")

    st.download_button(
        label="⬇️ Download samengevoegd bestand",
        data=output,
        file_name="samengevoegd.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
