from excel_operations.excel_writer import ExcelWriter

class Summary(ExcelWriter):
    def __init__(self, data_preparation):
        super().__init__(data_preparation)
        self.gesellschaft = data_preparation.gesellschaft

    def fill_worksheet(self):
        self.summary_sheet['D7'] = self.gesellschaft
        self.summary_sheet['D8'] = f"{self.jahr}_{self.quartal}"
        self.summary_sheet['H12'] = f"{self.account} {self.account_name}"