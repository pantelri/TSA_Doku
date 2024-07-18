from data_processing.preparation import DataPreparation

def main():
    data_preparation = DataPreparation()
    data_preparation.load_and_validate_data()
    data_preparation.print_klassenvariablen()
    data_preparation.print_dataframe()
    data_preparation.enrich_dataframe()
    data_preparation.print_dataframe()
    data_preparation.print_klassenvariablen()
    data_preparation.write_to_excel_template()

if __name__ == "__main__":
    main()
