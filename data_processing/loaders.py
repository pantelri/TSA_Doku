import pandas as pd
import os
import re

class DatenLaden:
    def __init__(self):
        self.data = None
        self.bestellnummer = None
        self.gesellschaft = None
        self.account = None
        self.jahr = None
        self.quartal = None
        self.spaltenueberschriften = None
        self.load_and_validate_data()

    def load_and_validate_data(self):
        try:
            # Finde die Datei, die mit "for_spotfire" beginnt
            input_dir = os.path.join(os.getcwd(), 'input')
            for file in os.listdir(input_dir):
                if file.startswith("for_spotfire"):
                    file_path = os.path.join(input_dir, file)
                    break
            else:
                raise FileNotFoundError("Keine Datei mit dem Präfix 'for_spotfire' gefunden.")

            # Extrahiere relevante Informationen aus dem Dateinamen
            pattern = r"for_spotfire_(.*?)_(.*?)_([A-Z]{2})_(\d{4})_(Q\d)_(.*?).xlsx"
            match = re.search(pattern, file)
            if match:
                self.bestellnummer = match.group(1)
                self.gesellschaft = match.group(2)
                self.account = match.group(3)  #Muss aktuell 2 Buchstaben haben. Bisher Fehlermeldung für Accounts mit einem Buchstaben.
                self.jahr = match.group(4)
                self.quartal = match.group(5)

            else:
                raise ValueError("Der Dateiname entspricht nicht dem erwarteten Muster.")

            # Lese die Excel-Datei ein
            self.data = pd.read_excel(file_path)

            # Überprüfe, ob die erforderlichen Spalten vorhanden sind
            if self.data.columns[0] != "Date" or self.data.columns[1] != "IndexAll":
                raise ValueError("Die ersten beiden Spalten müssen 'Date' und 'IndexAll' heißen.")

            if "IndexE" not in self.data.columns:
                raise ValueError("Die Spalte 'IndexE' fehlt in der Excel-Liste.")

            # Speichere alle Spaltenüberschriften
            self.spaltenueberschriften = self.data.columns.tolist()

        except FileNotFoundError as e:
            print(e)
        except ValueError as e:
            print(e)

    def print_klassenvariablen(self):
        print(f"Bestellnummer: {self.bestellnummer}")
        print(f"Gesellschaft: {self.gesellschaft}")
        print(f"Account: {self.account}")
        print(f"Jahr: {self.jahr}")
        print(f"Quartal: {self.quartal}")
        print(f"Spaltenüberschriften: {self.spaltenueberschriften}")

    def print_dataframe(self):
        print(f"""Dataframe Shape
              {self.data.shape}
              Dataframe Ausgabe
              {self.data}""")