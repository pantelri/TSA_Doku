from openpyxl_add_ons.related_cells import get_cell_name
from .planning_writer_functions import create_parameter_plot
import os
from openpyxl.drawing.image import Image

class Planning():
    def __init__(self, ExcelWriter):

        self.planning_sheet = ExcelWriter.planning_sheet
        self.output_dir = ExcelWriter.output_dir

        self.account = ExcelWriter.account
        self.account_name = ExcelWriter.account_name
        self.jahr = ExcelWriter.jahr
        self.significant_params = ExcelWriter.significant_params
        self.data = ExcelWriter.data

        self.starting_points = ["B4", "F4", "B35", "F35"]

    def fill_worksheet(self):
        for i, param in enumerate(self.significant_params):
            cell = self.starting_points[i]
            self.planning_sheet[cell] = f"Development of {param}"
            self.fill_param_table(cell, param)
            self.create_param_plot(param)
            self.insert_param_plot(cell, param)

    def fill_param_table(self, start_cell, param):
        # Gruppiere die Daten nach Fiscal_Year und summiere den Parameter
        grouped_data = self.data.groupby('Fiscal_Year')[param].sum().sort_index()

        # Startpunkt für die Tabelle (3 Zeilen unter dem Startzellenpunkt)
        current_row = int(start_cell[1:]) + 3

        # Schreibe die Überschriften
        self.planning_sheet[get_cell_name(start_cell, 0, 2)] = "Fiscal Year"
        self.planning_sheet[get_cell_name(start_cell, 1, 2)] = "Annual total"
        self.planning_sheet[get_cell_name(start_cell, 2, 2)] = "Change compared to PY"

        prev_value = None
        for fiscal_year, value in grouped_data.items():
            # Schreibe Fiscal Year
            self.planning_sheet[get_cell_name(start_cell, 0, current_row - int(start_cell[1:]))] = f"FY {fiscal_year[-2:]}"
            
            # Schreibe Summe oder Durchschnitt
            if param.startswith("index_"):
                self.planning_sheet[get_cell_name(start_cell, 1, current_row - int(start_cell[1:]))] = value / 12
            else:
                self.planning_sheet[get_cell_name(start_cell, 1, current_row - int(start_cell[1:]))] = value
            
            # Berechne und schreibe Veränderung zum Vorjahr
            if prev_value is not None:
                change = value - prev_value
                change = change / prev_value
                self.planning_sheet[get_cell_name(start_cell, 2, current_row - int(start_cell[1:]))] = change
            
            prev_value = value
            current_row += 1

    def create_param_plot(self, param):
        create_parameter_plot(self.data, param, self.output_dir)

    def insert_param_plot(self, start_cell, param):
        img_path = os.path.join(self.output_dir, f"{param}_plot.png")
        img = Image(img_path)
        
        # Berechne die Zelle, in die das Bild eingefügt werden soll (8 Zeilen unter start_cell)
        insert_cell = get_cell_name(start_cell, 0, 8)
        
        # Füge das Bild ein
        self.planning_sheet.add_image(img, insert_cell)
