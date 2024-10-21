from excel_operations.excel_writer import ExcelWriter
from excel_operations.validation_writer_functions import (
    write_basic_data,
    write_total_data,
    write_subtotal_data,
    write_volume_data,
    write_index_data
)

class Validation(ExcelWriter):
    def __init__(self, data_preparation):
        super().__init__(data_preparation)

# Optionales todo: Eigentlich brauchen wir nicht eine separate Funktion, um alle verschiedenenen Parameter einzeln einzutragen. 
# Eigentlich können wir aus dem Dataframe auch die "IndexAll" und "IndexE" columns entfernen und den Rest direkt reinschreiben. 
    def fill_worksheet(self):
        write_basic_data(self.validation_sheet,self.data)
        write_total_data(self.validation_sheet, self.data, self.account)
        write_subtotal_data(self.validation_sheet, self.data, self.account, self.account_name)
        write_volume_data(self.validation_sheet, self.data)
        write_index_data(self.validation_sheet, self.data)