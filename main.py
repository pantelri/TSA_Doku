from data_processing.loaders import DatenLaden
from data_processing.preparation import DataPreparation


def main():
    daten_analyse = DatenLaden()
    # Gib Informationen über den DataFrame aus
    daten_analyse.print_klassenvariablen()
    daten_analyse.print_dataframe()
    data_preparation = DataPreparation()
    # Reicher den DataFrame mit weiteren Informationen an
    data_preparation.enrich_dataframe()
    data_preparation.print_dataframe()
    
if __name__ == "__main__":
    main()