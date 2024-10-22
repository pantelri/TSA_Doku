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
        self.significant_params = data_preparation.significant_params
        self.data = data_preparation.data
        
        self.workbook = None
        self.summary_sheet = None
        self.validation_sheet = None
        self.planning_sheet = None
        self.execution_sheet = None

        # Erstelle den Output-Ordner, falls er nicht existiert
        output_dir = os.path.join(os.getcwd(), 'output')
        os.makedirs(output_dir, exist_ok=True)
        
        # Setze den output_path als Klassenvariable
        output_filename = f"{self.gesellschaft}_{self.account}_{self.jahr}_TSA_Doku.xlsx"
        self.output_path = os.path.join(output_dir, output_filename)

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

    def finalize_workbook(self):
        #ToDo: m.E. im Moment nicht gebraucht, vielleicht wieder später aktivieren?
        # self.check_and_remove_empty_columns(sheet)
        # self.remove_duplicate_columns(sheet)
        workbook = self.workbook
        workbook.save(self.output_path)