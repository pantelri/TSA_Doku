from excel_operations.excel_writer import ExcelWriter

class Summary(ExcelWriter):
    def __init__(self, data_preparation):
        super().__init__(data_preparation)
        self.gesellschaft = data_preparation.gesellschaft
        self.summary_sheet = None

    def fill_worksheet(self):
        if self.workbook is None:
            self.prepare_workbook()
        
        self.summary_sheet = self.workbook.create_sheet("Summary")
        
        if self.summary_sheet:
            self.summary_sheet['D7'] = self.gesellschaft
            self.summary_sheet['D8'] = f"{self.jahr}_{self.quartal}"
            self.summary_sheet['H12'] = f"{self.account} {self.account_name}"
        else:
            print("Fehler: Summary-Sheet konnte nicht erstellt werden.")
