import os
import shutil
from openpyxl import load_workbook
import numpy as np
from excel_operations.excel_data_writers import (
    write_basic_data,
    write_total_data,
    write_subtotal_data,
    write_cos_data,
    write_volume_data,
    write_index_data
)
from excel_operations.planning_writer import PlanningWriter

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
        self.adjust_row_6(sheet)
        self.calculate_sums_and_averages(sheet)
        workbook.save(self.output_path)
        
        # Erstellen und Verarbeiten des PlanningWriter
        planning_writer = PlanningWriter(self.output_path)
        planning_writer.process()

    def adjust_row_6(self, sheet):
        for col in range(7, sheet.max_column + 1):  # Start from column G
            sheet.cell(row=6, column=col).value = sheet.cell(row=27, column=col).value

    def calculate_sums_and_averages(self, sheet):
        for col in range(7, sheet.max_column + 1):  # Start from column G
            header = sheet.cell(row=27, column=col).value
            is_price_index = "price index" in str(header).lower() if header else False

            for row, (start, end) in zip([7, 11, 15], [(28, 39), (40, 51), (52, 63)]):
                values = [sheet.cell(row=r, column=col).value for r in range(start, end + 1)]
                values = [v for v in values if v is not None]  # Remove None values

                if values:
                    if is_price_index:
                        result = np.mean(values)
                    else:
                        result = sum(values)

                    sheet.cell(row=row, column=col).value = result

    def check_and_remove_empty_columns(self, sheet):
        # Finde den am weitesten rechts stehenden Wert in den Zeilen 27 bis 63
        rightmost_col = 0
        for col in range(1, sheet.max_column + 1):
            if any(sheet.cell(row=row, column=col).value is not None for row in range(27, 64)):
                rightmost_col = col

        # Lösche Spalten rechts vom am weitesten rechts stehenden Wert
        if rightmost_col > 0:
            delete_start = rightmost_col + 1
            delete_end = sheet.max_column
            if delete_start <= delete_end:
                sheet.delete_cols(delete_start, delete_end - delete_start + 1)

        # Überprüfe und entferne leere Spalten
        col_idx = ord('J') - ord('A')
        while col_idx < sheet.max_column:
            col_letter = chr(ord('A') + col_idx)
            if all(sheet[f'{col_letter}{row}'].value is None for row in range(27, 64)):
                sheet.delete_cols(col_idx + 1)
            else:
                col_idx += 1

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

