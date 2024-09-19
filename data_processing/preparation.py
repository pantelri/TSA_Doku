import pandas as pd
from datetime import datetime

from data_processing.loaders import DatenLaden
from excel_operations.excel_writer import ExcelWriter

class DataPreparation(DatenLaden):
    def __init__(self):
        super().__init__()  # Initialisiere die Superklasse, um Zugriff auf deren Variablen zu erhalten
        self.account_name = None

    def enrich_dataframe(self):
        # Anreichern des DataFrames mit zusätzlichen Informationen
        if self.data is not None:
            # Stelle sicher, dass 'Date' als String vorliegt
            self.data['Date'] = self.data['Date'].astype(str)

            # Korrigiere das Datum für Oktober -> Sonst wird Oktober als Januar behandelt
            self.data['Date'] = self.data['Date'].apply(
                lambda x: x[:-1] + '10' if x.endswith('.1') else x
            )

            # Monatsnamen extrahieren
            self.data['Month'] = self.data['Date'].apply(
                lambda x: datetime.strptime(x, '%Y.%m').strftime('%b')
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

            # Fiscal Year bestimmen
            def determine_fiscal_year(date, period):
                year, month = map(int, date.split('.'))
                if period == 'Audit Period':
                    fiscal_year = last_year % 100
                else:
                    months_difference = (last_year - year) * 12 + (last_month - month)
                    fiscal_year = (last_year - (months_difference // 12)) % 100
                return f'FY {fiscal_year:02d}'

            self.data['Fiscal_Year'] = self.data.apply(lambda row: determine_fiscal_year(row['Date'], row['Period']), axis=1)

            # Sortiere den DataFrame nach dem Datum
            self.data = self.data.sort_values('Date')

            # Speichere alle Spaltenüberschriften
            self.spaltenueberschriften = self.data.columns.tolist()

            # Finde die Spalte, die "_total" enthält und setze account_name
            total_column = next((col for col in self.data.columns if '_total' in col), None)
            if total_column:
                self.account_name = total_column.split('_total')[0]


