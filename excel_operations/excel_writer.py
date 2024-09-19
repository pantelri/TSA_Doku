import os
import shutil
from openpyxl import load_workbook
from excel_operations.excel_data_writers import (
    write_basic_data,
    write_total_data,
    write_subtotal_data,
    write_cos_data,
    write_volume_data,
    write_index_data
)

class ExcelWriter:
    def __init__(self, data_preparation):
        self.gesellschaft = data_preparation.gesellschaft
        self.account = data_preparation.account
        self.account_name = data_preparation.account_name
        self.jahr = data_preparation.jahr
        self.quartal = data_preparation.quartal
        self.spaltenueberschriften = data_preparation.spaltenueberschriften
        self.data = data_preparation.data
        
        # Erstelle den Output-Ordner, falls er nicht existiert
        output_dir = os.path.join(os.getcwd(), 'output')
        os.makedirs(output_dir, exist_ok=True)
        
        # Setze den output_path als Klassenvariable
        output_filename = f"{self.gesellschaft}_{self.account}_{self.jahr}_TSA_Doku.xlsx"
        self.output_path = os.path.join(output_dir, output_filename)

    def write_to_excel_template(self):
        workbook, sheet = self.prepare_workbook()
        
        write_basic_data(sheet, self.data)
        write_total_data(sheet, self.data, self.account)
        write_subtotal_data(sheet, self.data, self.account, self.account_name)
        write_cos_data(sheet, self.data)
        write_volume_data(sheet, self.data)
        write_index_data(sheet, self.data)
        
        self.finalize_workbook(workbook, sheet)

    def prepare_workbook(self):
        # Kopiere das Template
        template_path = os.path.join('templates', 'SAP_TSA_Template.xlsx')
        shutil.copy(template_path, self.output_path)

        # Lade die kopierte Arbeitsmappe
        workbook = load_workbook(self.output_path)
        sheet = workbook['1. Data Validation']
        return workbook, sheet

    def finalize_workbook(self, workbook, sheet):
        self.check_and_remove_empty_columns(sheet)
        self.remove_duplicate_columns(sheet)
        workbook.save(self.output_path)

    def check_and_remove_empty_columns(self, sheet):
        for col_idx in range(ord('J') - ord('A'), ord('P') - ord('A')):
            col_letter = chr(ord('A') + col_idx)
            if all(sheet[f'{col_letter}{row}'].value is None for row in range(28, 64)):
                sheet[f'{col_letter}27'].value = None
                sheet.delete_cols(col_idx + 1)
                # Verschiebe alle Zellen rechts davon um eine Spalte nach links
                for row in sheet.iter_rows(min_row=27, max_row=63, min_col=col_idx + 2):
                    for cell in row:
                        sheet.cell(row=cell.row, column=cell.column - 1, value=cell.value)
                sheet.delete_cols(sheet.max_column)

    def remove_duplicate_columns(self, sheet):
        values = {}
        for col in range(6, sheet.max_column + 1):
            value = sheet.cell(row=27, column=col).value
            if value in values:
                # Lösche die gesamte Spalte des zweiten Wertes
                sheet.delete_cols(col)
            else:
                values[value] = col

    def print_unwritten_columns(self):
        written_columns = {'Date', 'Month', 'Fiscal_Year', 'Period'}
        written_columns.update(col for col in self.data.columns if '_total' in col)
        written_columns.update(col for col in self.data.columns if f"{self.account_name}_subtotal" in col)
        written_columns.update(col for col in self.data.columns if 'Volume' in col)
        written_columns.update(col for col in self.data.columns if 'index_' in col)

        unwritten_columns = set(self.data.columns) - written_columns
        if unwritten_columns:
            print("Folgende Spalten wurden nicht in die Excel-Datei geschrieben:")
            for col in unwritten_columns:
                print(f"- {col}")
        else:
            print("Alle Spalten des DataFrames wurden in die Excel-Datei geschrieben.")

