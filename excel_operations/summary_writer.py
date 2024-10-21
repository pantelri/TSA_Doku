class Summary():
    def __init__(self, ExcelWriter):

        self.summary_sheet = ExcelWriter.summary_sheet

        self.gesellschaft = ExcelWriter.gesellschaft
        self.account = ExcelWriter.account
        self.account_name = ExcelWriter.account_name
        self.jahr = ExcelWriter.jahr
        self.quartal = ExcelWriter.quartal

    def fill_worksheet(self):

        self.summary_sheet['D7'] = self.gesellschaft
        self.summary_sheet['D8'] = f"{self.jahr}_{self.quartal}"
        self.summary_sheet['H12'] = f"{self.account} {self.account_name}"

