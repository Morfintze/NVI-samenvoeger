import streamlit as st
import openpyxl
from io import BytesIO

st.title("📊 Excel Samenvoeger (Veilige versie)")

st.markdown("""
Upload hier meerdere `.xlsx` bestanden. De inhoud van elk volgend bestand wordt **onder de vorige** geplakt.
""")

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

    for f in uploaded_files[1:]:
        wb_new = openpyxl.load_workbook(f)
        ws_new = wb_new.active

        first_empty_row = ws_base.max_row + 1

        # Kopieer alleen waarden (geen stijlen)
        for row_idx, row in enumerate(ws_new.iter_rows(values_only=True), start=first_empty_row):
            for col_idx, cell_value in enumerate(row, start=1):
                ws_base.cell(row=row_idx, column=col_idx, value=cell_value)

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
