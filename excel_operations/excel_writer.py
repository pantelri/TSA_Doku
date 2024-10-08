import os
import shutil
from openpyxl import load_workbook
from excel_data_writers import (
    write_basic_data,
    write_total_data,
    write_subtotal_data,
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
        self.remove_empty_columns(sheet)
        workbook.save(self.output_path)
        print(f"Excel-Datei wurde erstellt: {self.output_path}")
        self.print_unwritten_columns()

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

    def remove_empty_columns(self, sheet):
        for col_idx in range(ord('Q') - ord('A'), ord('G') - ord('A'), -1):
            col_letter = chr(ord('A') + col_idx)
            if sheet[f'{col_letter}27'].value is None:
                sheet.delete_cols(col_idx + 1)
