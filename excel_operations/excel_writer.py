import os
import shutil
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter

class ExcelWriter:
    def __init__(self, data_preparation):
        self.gesellschaft = data_preparation.gesellschaft
        self.account = data_preparation.account
        self.account_name = data_preparation.account_name
        self.jahr = data_preparation.jahr
        self.quartal = data_preparation.quartal
        self.spaltenueberschriften = data_preparation.spaltenueberschriften
        self.data = data_preparation.data

    def write_to_excel_template(self):
        # Erstelle den Output-Ordner, falls er nicht existiert
        output_dir = os.path.join(os.getcwd(), 'output')
        os.makedirs(output_dir, exist_ok=True)

        # Kopiere das Template
        template_path = os.path.join('templates', 'SAP_TSA_Template.xlsx')
        output_filename = f"{self.gesellschaft}_{self.account}_{self.jahr}_TSA_Doku.xlsx"
        output_path = os.path.join(output_dir, output_filename)
        shutil.copy(template_path, output_path)

        # Lade die kopierte Arbeitsmappe
        workbook = load_workbook(output_path)
        sheet = workbook['1. Data Validation']

        # Schreibe die Daten in die Excel-Datei
        for i, row in enumerate(self.data.itertuples(), start=28):
            sheet[f'B{i}'] = row.Date
            sheet[f'C{i}'] = row.Month
            sheet[f'D{i}'] = row.Fiscal_Year
            sheet[f'E{i}'] = row.Period

        # Fülle G27 mit dem Wert aus self.account + " total"
        sheet['G27'] = f"{self.account} total"

        # Finde die Spalte, die "_total" enthält
        total_column = next((col for col in self.data.columns if '_total' in col), None)

        if total_column:
            # Fülle G28 und darunter mit den Werten aus der "_total" Spalte
            for i, value in enumerate(self.data[total_column], start=28):
                sheet[f'G{i}'] = value

            # Finde alle Spalten mit "{self.account_name}_subtotal" im Namen
            subtotal_columns = [col for col in self.data.columns if f"{self.account_name}_subtotal" in col]

            # Überprüfe, ob es mehr als drei Subtotal-Spalten gibt
            if len(subtotal_columns) > 3:
                raise ValueError("Doku-Template wird bisher nur für max. 3 Subtotal-Spalten des Accounts_to_audit unterstützt")

            # Fülle die Zellen H27:J27 und die Spalten darunter
            for idx, col in enumerate(subtotal_columns):
                cell = chr(ord('H') + idx)  # H, I, J
                suffix = col.split(f"{self.account_name}_subtotal_")[-1]
                sheet[f'{cell}27'] = f"{self.account} {suffix}"

                # Fülle die Spalte darunter
                for i, value in enumerate(self.data[col], start=28):
                    sheet[f'{cell}{i}'] = value

        # Finde alle Spalten mit "Volume" im Namen
        volume_columns = [col for col in self.data.columns if 'Volume' in col]

        # Überprüfe, ob es mehr als drei Volume-Spalten gibt
        if len(volume_columns) > 3:
            # Füge zusätzliche Spalten ein
            sheet.insert_cols(15, len(volume_columns) - 3)

        # Fülle die Zellen M27:O27 (oder mehr) und die Spalten darunter
        for idx, col in enumerate(volume_columns):
            cell = get_column_letter(13 + idx)  # M, N, O, ...
            header = col.replace('_', ' ')
            sheet[f'{cell}27'] = header

            # Fülle die Spalte darunter
            for i, value in enumerate(self.data[col], start=28):
                sheet[f'{cell}{i}'] = value

        # Finde alle Spalten mit "index_" im Namen
        index_columns = [col for col in self.data.columns if 'index_' in col]

        # Überprüfe, ob es mehr als zwei Index-Spalten gibt
        if len(index_columns) > 2:
            # Füge zusätzliche Spalten ein
            sheet.insert_cols(18, len(index_columns) - 2)

        # Fülle die Zellen Q27:R27 (oder mehr) und die Spalten darunter
        for idx, col in enumerate(index_columns):
            cell = get_column_letter(17 + idx)  # Q, R, ...
            header = f"price {col.replace('_', ' ')}"
            sheet[f'{cell}27'] = header

            # Fülle die Spalte darunter
            for i, value in enumerate(self.data[col], start=28):
                sheet[f'{cell}{i}'] = value

        # Speichere die Änderungen
        workbook.save(output_path)
        print(f"Excel-Datei wurde erstellt: {output_path}")

        # Prüfe und gib nicht geschriebene Spalten aus
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
