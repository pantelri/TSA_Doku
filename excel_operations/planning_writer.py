class Planning():
    def __init__(self, ExcelWriter):

        self.planning_sheet = ExcelWriter.planning_sheet

        self.account = ExcelWriter.account
        self.account_name = ExcelWriter.account_name
        self.jahr = ExcelWriter.jahr
        self.significant_params = ExcelWriter.significant_params
        self.data = ExcelWriter.data

        self.starting_points = ["B4", "F4", "B35", "F35"]

    def fill_worksheet(self):
        for cell in self.starting_points:
            for i, param in enumerate(self.significant_params):
                self.planning_sheet[cell] = f"Development of {param}"
                if i == len(self.significant_params) - 1:
                    break
