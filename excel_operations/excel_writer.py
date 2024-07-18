import pandas as pd
from openpyxl import load_workbook
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.styles import Font, Alignment, Border, Side

class ExcelWriter:
    @staticmethod
    def write_to_excel_template(data, gesellschaft, account, jahr, account_name, spaltenueberschriften):
        template_path = 'Vorlage_Kontenblatt.xlsx'
        output_path = f'{gesellschaft}_{account}_{jahr}.xlsx'

        # Lade das Excel-Template
        workbook = load_workbook(template_path)
        sheet = workbook.active

        # Schreibe die Kopfzeile
        sheet['B1'] = f'{gesellschaft} - {account_name}'
        sheet['B2'] = f'Kontenblatt {account}'
        sheet['B3'] = f'Geschäftsjahr {jahr}'

        # Schreibe die Daten
        for r_idx, row in enumerate(dataframe_to_rows(data, index=False, header=True), 6):
            for c_idx, value in enumerate(row, 2):
                sheet.cell(row=r_idx, column=c_idx, value=value)

        # Formatierung
        for row in sheet[f'B6:K{sheet.max_row}']:
            for cell in row:
                cell.border = Border(left=Side(style='thin'), 
                                     right=Side(style='thin'), 
                                     top=Side(style='thin'), 
                                     bottom=Side(style='thin'))
                cell.alignment = Alignment(horizontal='center', vertical='center')

        # Formatiere die Überschriften
        for cell in sheet[6]:
            cell.font = Font(bold=True)

        # Speichere die Datei
        workbook.save(output_path)
