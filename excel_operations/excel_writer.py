import os
import shutil
from openpyxl import load_workbook

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

        # Für leichteren Zugriff auf die einzelnen Sheets (=Tabs) in der Zukunft sollen in dieser Klasse alle Tabs leicht einzeln aufrufbar sein.
        self.workbook = None
        self.summary_sheet = None
        self.validation_sheet = None 
        self.planning_sheet = None
        self.execution_sheet = None
        self.prepare_workbook()

    def prepare_workbook(self):
        # Kopiere das Template
        template_path = os.path.join('templates', 'SAP_TSA_Template.xlsx')
        shutil.copy(template_path, self.output_path)

        # Lade die kopierte Arbeitsmappe
        self.workbook = load_workbook(self.output_path)
        self.summary_sheet = self.workbook["SAP Summary"]
        self.validation_sheet = self.workbook["1. Data Validation"]
        self.planning_sheet = self.workbook["2.1 SAP Planning"]
        self.execution_sheet = self.workbook["2.2 SAP Execution"]

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
        for col in range(1, sheet.max_column + 1):
            value = sheet.cell(row=27, column=col).value
            if value in values:
                # Lösche die gesamte Spalte des zweiten Wertes
                sheet.delete_cols(col)
            else:
                values[value] = col