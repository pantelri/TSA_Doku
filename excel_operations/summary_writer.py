from excel_operations.excel_writer import ExcelWriter

class Summary(ExcelWriter):
    def __init__(self):
        super().__init__()

    def fill_worksheet(self):
        self.summary_sheet['D7'] = self.gesellschaft
        self.summary_sheet['D8'] = f"{self.jahr}_{self.quartal}"
        self.summary_sheet['H12'] = f"{self.account} {self.account_name}"