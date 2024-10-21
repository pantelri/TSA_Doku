from data_processing.preparation import DataPreparation
from excel_operations.excel_writer import ExcelWriter
from excel_operations.summary_writer import Summary
from excel_operations.validation_writer import Validation

def main():
    data_preparation = DataPreparation()
    data_preparation.load_and_validate_data()
    data_preparation.enrich_dataframe()

    excel_writer = ExcelWriter(data_preparation)
    excel_writer.prepare_workbook()

    tab1 = Summary(data_preparation)
    tab1.fill_worksheet()

    tab2 = Validation(data_preparation)
    tab2.fill_worksheet()

    # data_preparation = DataPreparation()
    # data_preparation.load_and_validate_data()
    # # data_preparation.print_klassenvariablen()
    # # data_preparation.print_dataframe()
    # data_preparation.enrich_dataframe()
    # # data_preparation.print_dataframe()
    # # data_preparation.print_klassenvariablen()

    # excel_writer = ExcelWriter(data_preparation)
    # #Todo as of 19.09./14:06
    # excel_writer.prepare_workbook()
    # excel_writer.write_to_excel_template()
    # print(f"Excel-Datei wurde im output-Ordner erstellt")
    # excel_writer.print_unwritten_columns()

    # summary_writer = Summary(data_preparation, excel_writer)
    # summary_writer.write_summary()

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

# tab3 = Evaluation()
# tab3.fill_worksheet()

# excel_writer.finalize_workbook()