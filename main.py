from data_processing.preparation import DataPreparation
from excel_operations.excel_writer import ExcelWriter
from excel_operations.summary_writer import Summary
from excel_operations.validation_writer import Validation

def main():
    data_preparation = DataPreparation()
    data_preparation.load_and_validate_data()
    data_preparation.enrich_dataframe()

    excel_writer = ExcelWriter(data_preparation)

    tab1 = Summary(ExcelWriter)
    tab1.fill_worksheet()

    tab2 = Validation(data_preparation)
    tab2.fill_worksheet()

    # Speichern der Arbeitsmappe
    excel_writer.finalize_workbook()

if __name__ == "__main__":
    main()


### GOAL MAIN: 
# data_preparation = DataPreparation()
# data_preparation.load_and_validate_data()
# data_preparation.enrich_dataframe()

# excel_writer = ExcelWriter(data_preparation)
# excel_writer.prepare_workbook()

# tab1 = Summary()
# tab1.fill_worksheet()

# tab2 = Validation()
# tab2.fill_worksheet()

# tab2 = Planning()
# tab2.fill_worksheet()

#todo: --> 
# tab3 = Evaluation()
# tab3.fill_worksheet()

# excel_writer.finalize_workbook()
