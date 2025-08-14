import pandas as pd
import tkinter as tk
from tkinter import filedialog
from openpyxl import load_workbook, Workbook
from openpyxl.styles import PatternFill, Font, Border, Alignment
from copy import copy

# Functie om de gebruiker een getal te laten invoeren
def get_number_of_files():
    while True:
        try:
            num_files = int(input("Hoeveel Excel-bestanden wil je samenvoegen? (minimaal 2): "))
            if num_files >= 2:
                return num_files
            else:
                print("Je moet minimaal 2 bestanden kiezen.")
        except ValueError:
            print("Voer een geldig getal in.")

# Functie om een Excel-bestand te selecteren via GUI
def select_excel_file(prompt):
    root = tk.Tk()
    root.withdraw()  # Verberg het hoofdvenster
    file_path = filedialog.askopenfilename(title=prompt, filetypes=[("Excel-bestanden", "*.xls;*.xlsx")])
    if not file_path:
        print("Geen bestand geselecteerd. Script wordt afgesloten.")
        exit()
    return file_path

# Functie om een werkblad inclusief opmaak te kopiëren
def copy_sheet(source_ws, target_ws):
    for row_idx, row in enumerate(source_ws.iter_rows(), start=1):
        for col_idx, cell in enumerate(row, start=1):
            new_cell = target_ws.cell(row=row_idx, column=col_idx, value=cell.value)

            # Kopieer de celopmaak
            if cell.font:
                new_cell.font = copy(cell.font)
            if cell.fill:
                new_cell.fill = copy(cell.fill)
            if cell.border:
                new_cell.border = copy(cell.border)
            if cell.alignment:
                new_cell.alignment = copy(cell.alignment)

    # Kopieer kolombreedtes en rijhoogtes
    for col_letter, col_dim in source_ws.column_dimensions.items():
        target_ws.column_dimensions[col_letter].width = col_dim.width

    for row_idx, row_dim in source_ws.row_dimensions.items():
        target_ws.row_dimensions[row_idx].height = row_dim.height

# Functie om gegevens en opmaak correct te kopiëren naar "Samengevoegd" zonder extra lege rijen
def merge_sheets(wb):
    merged_sheet = wb.create_sheet(title="Samengevoegd")
    max_col_widths = {}
    max_row_heights = {}
    current_row = 1  # Beginrij voor de samengevoegde data

    for sheet_name in wb.sheetnames:
        if sheet_name == "Samengevoegd":
            continue  # Sla het samengevoegde tabblad zelf over

        ws_temp = wb[sheet_name]

        for row in ws_temp.iter_rows():
            empty_row = all(cell.value is None for cell in row)  # Controleer of de rij leeg is
            if not empty_row:  # Alleen niet-lege rijen toevoegen
                for cell in row:
                    new_cell = merged_sheet.cell(row=current_row, column=cell.column, value=cell.value)

                    # Kopieer opmaak
                    if cell.font:
                        new_cell.font = copy(cell.font)
                    if cell.fill:
                        new_cell.fill = copy(cell.fill)
                    if cell.border:
                        new_cell.border = copy(cell.border)
                    if cell.alignment:
                        new_cell.alignment = copy(cell.alignment)

                current_row += 1  # Ga naar de volgende rij

        # Update de maximale kolombreedtes en rijhoogtes, maar vermijd extreme waarden
        for col_letter, col_dim in ws_temp.column_dimensions.items():
            width = col_dim.width
            if col_letter not in max_col_widths or (width and width < 50):  # Beperk onnodig brede kolommen
                max_col_widths[col_letter] = width

        for row_idx, row_dim in ws_temp.row_dimensions.items():
            height = row_dim.height
            if row_idx not in max_row_heights or (height and height < 50):  # Beperk onnodig hoge rijen
                max_row_heights[row_idx] = height

    # Pas de kolombreedtes toe
    for col_letter, width in max_col_widths.items():
        if width:
            merged_sheet.column_dimensions[col_letter].width = min(width, 50)  # Max breedte van 50 instellen

    # Pas de rijhoogtes toe
    for row_idx, height in max_row_heights.items():
        if height:
            merged_sheet.row_dimensions[row_idx].height = min(height, 50)  # Max hoogte van 50 instellen

    return merged_sheet

# Vraag hoeveel bestanden de gebruiker wil samenvoegen
num_files = get_number_of_files()

# Vraag de gebruiker om de Excel-bestanden te selecteren
excel_paths = [select_excel_file(f"Selecteer Excel-bestand {i+1}") for i in range(num_files)]

# Maak een nieuw, leeg Excel-bestand
wb_combined = load_workbook(excel_paths[0])  # Neem het eerste bestand als basis
while len(wb_combined.sheetnames) > 0:  # Verwijder alle bestaande tabbladen
    wb_combined.remove(wb_combined[wb_combined.sheetnames[0]])

# Kopieer elk bestand naar een apart tabblad
for i, excel_path in enumerate(excel_paths, start=1):
    wb_temp = load_workbook(excel_path)
    ws_temp = wb_temp.active  # Neem het eerste tabblad van elk bestand

    # Maak een nieuw tabblad en kopieer de inhoud en opmaak
    new_sheet = wb_combined.create_sheet(title=f"Blad{i}")
    copy_sheet(ws_temp, new_sheet)

# Verwijder eerdere versies van "Samengevoegd"
for sheet_name in wb_combined.sheetnames:
    if sheet_name == "Samengevoegd":
        del wb_combined[sheet_name]

# Maak het samengevoegde tabblad opnieuw zonder onnodige lege rijen en correcte formaten
merge_sheets(wb_combined)

# Vraag de gebruiker om een bestandsnaam en locatie voor het gecombineerde bestand
output_path = filedialog.asksaveasfilename(defaultextension=".xlsx",
                                           filetypes=[("Excel-bestanden", "*.xlsx")],
                                           title="Opslaan als")
if not output_path:
    print("Geen bestandsnaam opgegeven. Script wordt afgesloten.")
    exit()

# Sla het gecombineerde bestand op
wb_combined.save(output_path)
print(f"Bestand succesvol opgeslagen als: {output_path}")
