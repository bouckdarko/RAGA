# ExcelFileHandler.py

import pandas as pd


class ExcelFileHandler:
    def __init__(self, file_path):
        """
        Classe pour gérer les fichiers Excel en utilisant pandas.

        Args:
            file_path (str): Chemin vers le fichier Excel.

        Attributes:
            file_path (str): Chemin vers le fichier Excel.

        Methods:
            read_excel(sheet_name: str = None) -> pd.DataFrame:
                Lit le fichier Excel et renvoie un DataFrame.

            write_to_excel(df: pd.DataFrame, sheet_name: str = "Sheet1") -> None:
                Écrit un DataFrame dans le fichier Excel.

            get_sheet_names() -> List[str]:
                Renvoie la liste des noms de toutes les feuilles du fichier Excel.

            describe_columns(sheet_name: str = None) -> None:
                Affiche les informations sur les colonnes du fichier Excel.
        """
        self.file_path = file_path

    def read_excel(self, sheet_name=None):
        """
        Lit le fichier Excel et renvoie un DataFrame.

        Args:
            sheet_name (str, optional): Nom de la feuille à lire. Si None, lit la première feuille.

        Returns:
            pd.DataFrame: DataFrame contenant les données du fichier Excel.
        """
        try:
            if sheet_name:
                df = pd.read_excel(self.file_path, sheet_name=sheet_name)
            else:
                df = pd.read_excel(self.file_path)
            return df
        except FileNotFoundError:
            print(f"Le fichier {self.file_path} n'a pas été trouvé.")
            return None

    def write_to_excel(self, df, sheet_name="Sheet1"):
        """
        Écrit un DataFrame dans le fichier Excel.

        Args:
            df (pd.DataFrame): DataFrame à écrire.
            sheet_name (str, optional): Nom de la feuille. Par défaut, "Sheet1".
        """
        try:
            with pd.ExcelWriter(self.file_path, mode="a", engine="openpyxl") as writer:
                df.to_excel(writer, sheet_name=sheet_name, index=False)
            print(f"Données écrites dans {self.file_path}.")
        except PermissionError:
            print(
                f"Impossible d'écrire dans {self.file_path}. Vérifiez les permissions."
            )

    def get_sheet_names(self):
        """
        Renvoie la liste des noms de toutes les feuilles du fichier Excel.

        Returns:
            list: Liste des noms de feuilles.
        """
        try:
            xls = pd.ExcelFile(self.file_path)
            return xls.sheet_names
        except FileNotFoundError:
            print(f"Le fichier {self.file_path} n'a pas été trouvé.")
            return []

    def describe_columns(self, sheet_name=None):
        """
        Affiche les informations sur les colonnes du fichier Excel.

        Args:
            sheet_name (str, optional): Nom de la feuille. Si None, lit la première feuille.
        """
        df = self.read_excel(sheet_name)
        if df is not None:
            print(df.info())

    def write_value_to_cell(self, sheet_name, row, column, value):
        """
        Écrit une valeur dans une cellule spécifiée du fichier Excel.

        Args:
            sheet_name (str): Nom de la feuille.
            row (int): Numéro de ligne (commence à 1).
            column (str): Nom de la colonne (par exemple, "A", "B", etc.).
            value: Valeur à écrire dans la cellule.

        Exemple d'utilisation :
        excel_handler = ExcelFileHandler("chemin/vers/votre/fichier.xlsx")
        excel_handler.write_value_to_cell(sheet_name="Feuille1", row=2, column="B", value=42)
        """
        try:
            xls = pd.ExcelFile(self.file_path)
            if sheet_name in xls.sheet_names:
                df = pd.read_excel(self.file_path, sheet_name=sheet_name)
                df.at[row - 1, column] = value
                with pd.ExcelWriter(
                    self.file_path, mode="a", engine="openpyxl"
                ) as writer:
                    df.to_excel(writer, sheet_name=sheet_name, index=False)
                print(
                    f"Valeur '{value}' écrite dans la cellule {column}{row} de la feuille '{sheet_name}'."
                )
            else:
                print(f"La feuille '{sheet_name}' n'existe pas dans le fichier Excel.")
        except PermissionError:
            print(
                f"Impossible d'écrire dans {self.file_path}. Vérifiez les permissions."
            )

    def read_cell_value(self, sheet_name, row, column):
        """
        Lit le contenu d'une cellule spécifiée du fichier Excel.

        Args:
            sheet_name (str): Nom de la feuille.
            row (int): Numéro de ligne (commence à 1).
            column (str): Nom de la colonne (par exemple, "A", "B", etc.).

        Returns:
            str: Contenu de la cellule.
        """
        try:
            xls = pd.ExcelFile(self.file_path)
            if sheet_name in xls.sheet_names:
                df = pd.read_excel(self.file_path, sheet_name=sheet_name)
                value = df.at[row - 1, column]
                return str(value)
            else:
                print(f"La feuille '{sheet_name}' n'existe pas dans le fichier Excel.")
                return None
        except FileNotFoundError:
            print(f"Le fichier {self.file_path} n'a pas été trouvé.")
            return None