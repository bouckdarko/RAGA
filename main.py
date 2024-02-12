# main.py

# Importez les modules que vous avez créés
from src.ExcelFileHandler import ExcelFileHandler
from src.chemin_absolu import chemin_absolu

def main():

    # Demandez à l'utilisateur de saisir le nom du fichier Excel
    filename = input("Entrez le nom du fichier Excel : ")

    # Créez le chemin absolu vers le fichier Excel
    chemin_absolu = chemin_absolu(filename)
    if chemin_absolu:
        print(f"Chemin absolu vers '{filename}': {chemin_absolu}")

        # Créez une instance de la classe ExcelFileHandler
        excel_handler = ExcelFileHandler(chemin_absolu)

        # Utilisez les méthodes de la classe ExcelFileHandler
        excel_handler.read_excel(sheet_name="Feuille1")
        excel_handler.write_to_excel(df, sheet_name="Feuille2")
        excel_handler.describe_columns(sheet_name="Feuille1")

    else:
        print("Le fichier n'a pas été trouvé. Vérifiez le nom et l'emplacement du fichier.")


if __name__ == "__main__":
    main()
