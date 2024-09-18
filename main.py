from data_processing.preparation import DataPreparation
from excel_operations.excel_writer import ExcelWriter
from excel_operations.summary_writer import SummaryWriter

def main():
    data_preparation = DataPreparation()
    data_preparation.load_and_validate_data()
    data_preparation.print_klassenvariablen()
    data_preparation.print_dataframe()
    data_preparation.enrich_dataframe()
    data_preparation.print_dataframe()
    data_preparation.print_klassenvariablen()

    excel_writer = ExcelWriter(data_preparation)
    excel_writer.write_to_excel_template()

    summary_writer = SummaryWriter(data_preparation, excel_writer)
    summary_writer.write_summary()

if __name__ == "__main__":
    main()
