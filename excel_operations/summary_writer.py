from openpyxl import load_workbook

class SummaryWriter:
    def __init__(self, data_preparation):
        self.gesellschaft = data_preparation.gesellschaft
        self.account = data_preparation.account
        self.account_name = data_preparation.account_name
        self.jahr = data_preparation.jahr
        self.quartal = data_preparation.quartal

    def write_summary(self, excel_path):
        workbook = load_workbook(excel_path)
        sheet = workbook['SAP Summary']

        # Schreibe Gesellschaft in Zelle D7
        sheet['D7'] = self.gesellschaft

        # Schreibe Jahr und Quartal in Zelle D8
        sheet['D8'] = f"{self.jahr}_{self.quartal}"

        # Schreibe Account und Account Name in Zelle H12
        sheet['H12'] = f"{self.account} {self.account_name}"

        workbook.save(excel_path)
        print("Summary information has been written to the SAP Summary tab.")
