from excel_operations.validation_writer_functions import (
    write_basic_data,
    write_total_data,
    write_subtotal_data,
    write_volume_data,
    write_index_data
)
from openpyxl.styles import PatternFill

class Validation():
    def __init__(self, ExcelWriter):
        self.validation_sheet = ExcelWriter.validation_sheet
        self.data = ExcelWriter.data
        self.account = ExcelWriter.account
        self.account_name = ExcelWriter.account_name

    def fill_worksheet(self):
        write_basic_data(self.validation_sheet, self.data)
        write_total_data(self.validation_sheet, self.data, self.account)
        write_subtotal_data(self.validation_sheet, self.data, self.account, self.account_name)
        write_volume_data(self.validation_sheet, self.data)
        write_index_data(self.validation_sheet, self.data)
        self.remove_yellow_columns()

    def remove_yellow_columns(self):
        if self.validation_sheet is None:
            return

        yellow_fill = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")
        columns_to_delete = []

        for cell in self.validation_sheet[27]:
            if cell.fill == yellow_fill:
                columns_to_delete.append(cell.column_letter)

        for column_letter in reversed(columns_to_delete):
            self.validation_sheet.delete_cols(self.validation_sheet[column_letter].column)

