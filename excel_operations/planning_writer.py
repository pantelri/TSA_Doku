from openpyxl_add_ons.related_cells import get_cell_name

class Planning():
    def __init__(self, ExcelWriter):

        self.planning_sheet = ExcelWriter.planning_sheet

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

    def fill_param_table(self, start_cell, param):
        # Gruppiere die Daten nach Fiscal_Year und summiere den Parameter
        grouped_data = self.data.groupby('Fiscal_Year')[param].sum().sort_index()

        # Startpunkt für die Tabelle (3 Zeilen unter dem Startzellenpunkt)
        current_row = int(start_cell[1:]) + 3
        current_col = ord(start_cell[0]) - ord('A')

        # Schreibe die Überschriften
        self.planning_sheet[get_cell_name(start_cell, 0, 3)] = "Fiscal Year"
        self.planning_sheet[get_cell_name(start_cell, 1, 3)] = "Sum"
        self.planning_sheet[get_cell_name(start_cell, 2, 3)] = "Change vs PY"

        prev_value = None
        for fiscal_year, value in grouped_data.items():
            # Schreibe Fiscal Year
            self.planning_sheet[get_cell_name(start_cell, 0, current_row - int(start_cell[1:]))] = f"FY {fiscal_year[-2:]}"
            
            # Schreibe Summe
            self.planning_sheet[get_cell_name(start_cell, 1, current_row - int(start_cell[1:]))] = value
            
            # Berechne und schreibe Veränderung zum Vorjahr
            if prev_value is not None:
                change = value - prev_value
                self.planning_sheet[get_cell_name(start_cell, 2, current_row - int(start_cell[1:]))] = change
            
            prev_value = value
            current_row += 1
