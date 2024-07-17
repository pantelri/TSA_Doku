import openpyxl
import pandas as pd
from datetime import datetime

from data_processing.loaders import DatenLaden

class DataPreparation(DatenLaden):
    def __init__(self):
        super().__init__()  # Initialisiere die Superklasse, um Zugriff auf deren Variablen zu erhalten
        self.template_path = 'templates/SAP_TSA_Template.xlsx'
        self.workbook = None
        self.load_template()

    def load_template(self):
        # Lade das Template und erstelle eine Kopie
        self.workbook = openpyxl.load_workbook(self.template_path)
        self.workbook.save('working_copy.xlsx')  # Speichere die Kopie unter einem neuen Namen

    def enrich_dataframe(self):
        # Anreichern des DataFrames mit zusätzlichen Informationen
        if self.data is not None:
            # Stelle sicher, dass 'Date' als String vorliegt
            self.data['Date'] = self.data['Date'].astype(str)

            # Monatsnamen extrahieren
            self.data['Month'] = self.data['Date'].apply(
                lambda x: datetime.strptime(x, '%Y.%m').strftime('%b') if '.' in x else None
            )

            # Audit Period bestimmen
            audit_period_mapping = {'Q1': 3, 'Q2': 6, 'Q3': 9, 'Q4': 12}
            audit_period_months = audit_period_mapping[self.quartal]
            last_month = self.data['Date'].max()
            last_year = int(last_month.split('.')[0])
            last_month = int(last_month.split('.')[1])

            def determine_period(date):
                year, month = map(int, date.split('.'))
                if year == last_year and month > (last_month - audit_period_months):
                    return 'Audit Period'
                else:
                    return 'Prior Period'

            self.data['Period'] = self.data['Date'].apply(determine_period)

    
    def fill_excel(self):
        # Fülle die Zellen B28:E63 des Tabs "1. Data Validation" aus
        sheet = self.workbook['1. Data Validation']
        for index, row in self.data.iterrows():
            # Hier müsste die Logik zum Befüllen der Zellen stehen
            # Beispiel: sheet.cell(row=28 + index, column=2, value=row['Month'])
            pass

        # Speichere die Änderungen in der Arbeitskopie
        self.workbook.save('working_copy.xlsx')
