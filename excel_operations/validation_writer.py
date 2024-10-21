from excel_operations.validation_writer_functions import (
    write_basic_data,
    write_total_data,
    write_subtotal_data,
    write_volume_data,
    write_index_data
)
from openpyxl.utils import get_column_letter

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
        self.copy_row_27_to_row_6()
        self.remove_empty_columns()

    def copy_row_27_to_row_6(self):
        if self.validation_sheet is None:
            return

        for col in range(7, 27):  # G to Z
            col_letter = get_column_letter(col)
            self.validation_sheet[f'{col_letter}6'] = self.validation_sheet[f'{col_letter}27'].value

    def remove_empty_columns(self):
        if self.validation_sheet is None:
            return

        columns_to_delete = []

        for cell in self.validation_sheet[27]:
            if cell.column_letter not in ['A', 'F'] and cell.value is None:
                columns_to_delete.append(cell.column)

        for column_index in reversed(columns_to_delete):
            self.validation_sheet.delete_cols(column_index)

