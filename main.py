from data_processing.loaders import DatenLaden
from data_processing.preparation import DataPreparation


from data_processing.preparation import DataPreparation

def main():
    data_prep = DataPreparation()
    data_prep.load_and_validate_data()
    data_prep.enrich_dataframe()
    data_prep.write_to_excel_template()
    daten_analyse = DatenLaden()
    # Gib Informationen über den DataFrame aus
    daten_analyse.print_klassenvariablen()
    daten_analyse.print_dataframe()
    data_preparation = DataPreparation()
    # Reicher den DataFrame mit weiteren Informationen an
    data_preparation.enrich_dataframe()
    data_preparation.print_dataframe()
    data_preparation.print_klassenvariablen()
    data_preparation.write_to_excel_template()
    
if __name__ == "__main__":
    main()
