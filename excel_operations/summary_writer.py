from openpyxl import load_workbook

class SummaryWriter:
    def __init__(self, data_preparation, excel_writer):
        self.gesellschaft = data_preparation.gesellschaft
        self.account = data_preparation.account
        self.account_name = data_preparation.account_name
        self.jahr = data_preparation.jahr
        self.quartal = data_preparation.quartal
        self.output_path = excel_writer.output_path

    def write_summary(self):
        workbook = load_workbook(self.output_path)
        sheet = workbook['SAP Summary']

        # Schreibe Gesellschaft in Zelle D7
        sheet['D7'] = self.gesellschaft

        # Schreibe Jahr und Quartal in Zelle D8
        sheet['D8'] = f"{self.jahr}_{self.quartal}"

        # Schreibe Account und Account Name in Zelle H12
        sheet['H12'] = f"{self.account} {self.account_name}"

        workbook.save(self.output_path)
        print(f"Summary information has been written to the SAP Summary tab in {self.output_path}")
