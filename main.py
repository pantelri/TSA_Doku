from data_processing.loaders import DatenLaden

def main():
    daten_analyse = DatenLaden()
    # Gib Informationen über den DataFrame aus
    daten_analyse.print_klassenvariablen()
    daten_analyse.print_dataframe()


    
if __name__ == "__main__":
    main()