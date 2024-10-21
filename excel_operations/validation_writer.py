from excel_operations.validation_writer_functions import (
    write_basic_data,
    write_total_data,
    write_subtotal_data,
    write_volume_data,
    write_index_data
)

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

