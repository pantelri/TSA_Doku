from data_processing.preparation import DataPreparation
from excel_operations.excel_writer import ExcelWriter
from excel_operations.summary_writer import Summary
from excel_operations.validation_writer import Validation
from excel_operations.planning_writer import Planning

def main():
    data_preparation = DataPreparation()
    data_preparation.load_and_validate_data()
    data_preparation.enrich_dataframe()

    excel_writer = ExcelWriter(data_preparation)

    tab1 = Summary(excel_writer)
    tab1.fill_worksheet()

    tab2 = Validation(excel_writer)
    tab2.fill_worksheet()

    tab3 = Planning(excel_writer)
    tab3.fill_worksheet()

    # Speichern der Arbeitsmappe
    excel_writer.finalize_workbook()

if __name__ == "__main__":
    main()