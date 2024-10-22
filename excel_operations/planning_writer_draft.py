from openpyxl.chart import BarChart, Reference
from openpyxl.styles import PatternFill
from openpyxl.chart import BarChart, Reference
from openpyxl.drawing.colors import ColorChoice, RGBPercent
from openpyxl.chart.shapes import GraphicalProperties

from excel_operations.excel_writer import ExcelWriter

class Planning(ExcelWriter):
    def __init__(self):
        super().__init__()

    def fill_worksheet(self):
        pass

    def create_column_chart(self):
        # Erstellen des Diagramms
        chart = BarChart()
        chart.type = "col"
        chart.style = 10
        chart.title = "Werte aus Spalte G"
        chart.y_axis.title = 'Wert'
        chart.x_axis.title = 'Datum'

        # Daten zum Diagramm hinzufügen
        data = Reference(self.planning_sheet, min_col=2, min_row=14, max_row=49, max_col=2)
        cats = Reference(self.planning_sheet, min_col=1, min_row=14, max_row=49)
        chart.add_data(data, titles_from_data=True)
        chart.set_categories(cats)

        # Farben für die Balken festlegen
        series = chart.series[0]
        series.graphicalProperties = GraphicalProperties()
        series.graphicalProperties.solidFill = ColorChoice(srgbClr=RGBPercent(31, 51, 74))  # Blaue Farbe für die ersten 24 Balken

        # Die letzten 12 Balken rot färben
        for i in range(24, 36):
            pt = GraphicalProperties()
            pt.solidFill = ColorChoice(srgbClr=RGBPercent(75, 31, 30))  # Rote Farbe für die letzten 12 Balken
            series.dPt.append(pt)

        # Diagramm zum Arbeitsblatt hinzufügen
        self.planning_sheet.add_chart(chart, "B13")

    def write_data_to_planning_sheet(self):
        # Daten aus Spalte G und B in den Planning-Sheet schreiben
        for i, row in enumerate(range(28, 64), start=14):
            date = self.data_validation_sheet.cell(row=row, column=2).value
            value = self.data_validation_sheet.cell(row=row, column=7).value
            self.planning_sheet.cell(row=i, column=1, value=date)
            self.planning_sheet.cell(row=i, column=2, value=value)

            # Farbliche Hervorhebung der letzten 12 Werte
            if i >= 38:  # Die letzten 12 Zeilen (52-63 im Originalsheet)
                self.planning_sheet.cell(row=i, column=1).fill = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")
                self.planning_sheet.cell(row=i, column=2).fill = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")

    def process(self):
        self.write_data_to_planning_sheet()
        self.create_column_chart()
        self.workbook.save(self.workbook.path)
