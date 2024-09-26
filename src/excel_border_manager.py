"""
Module excel_border_manager

Ce module permet de gérer les bordures des cellules dans un fichier Excel en utilisant openpyxl. 
Il offre des fonctionnalités pour appliquer des bordures à des cellules individuelles, des colonnes, 
des lignes ou des plages de cellules, avec une grande flexibilité pour personnaliser le style, la couleur et l'épaisseur des bordures.

Fonctionnalités :
----------------
1. Créer une bordure avec un style, une couleur et une épaisseur spécifiques.
2. Appliquer des bordures à une cellule, une colonne, une ligne ou une plage de cellules.
3. Gérer les bordures avec différents styles (solide, pointillé, double, etc.) et couleurs en hexadécimal.

Utilisation :
-------------
1. Créer une bordure :
    border = create_border(style="solid", color="000000")  # Bordure noire solide

2. Appliquer une bordure à une cellule :
    apply_border_to_cell("mon_fichier.xlsx", "A1", border)

3. Appliquer une bordure à une colonne :
    apply_border_to_column("mon_fichier.xlsx", "B", border)

4. Appliquer une bordure à une ligne :
    apply_border_to_row("mon_fichier.xlsx", 2, border)

5. Appliquer une bordure à une plage de cellules :
    apply_border_to_range("mon_fichier.xlsx", "A", "G", 1, 5, border)
    
Exemples de styles de bordure :
-------------------------------
- solid  : Bordure continue
- dashed : Bordure en pointillés
- double : Bordure double
- thick  : Bordure épaisse
- thin   : Bordure fine

Auteur : Bouck DARKO
Version : 1.0
"""

import openpyxl
from openpyxl.styles import Border, Side


def create_border(style: str, color: str = "000000") -> Border:
    """
    Crée une bordure avec le style et la couleur spécifiés.

    Args:
        style (str): Le style de la bordure (solid, dashed, double, etc.).
        color (str): La couleur de la bordure (en hexadécimal, par défaut noir).

    Returns:
        Border: Un objet Border à utiliser dans les cellules.
    """
    side = Side(style=style, color=color)
    return Border(left=side, right=side, top=side, bottom=side)


def apply_border_to_cell(file_path: str, cell_ref: str, border: Border, sheets: list = None) -> None:
    """
    Applique une bordure à une cellule spécifique.

    Args:
        file_path (str): Le chemin vers le fichier Excel.
        cell_ref (str): La référence de la cellule (ex: 'A1').
        border (Border): L'objet Border à appliquer.
        sheets (list): La liste des feuilles où effectuer la modification (par défaut, feuille active).
    """
    wb = openpyxl.load_workbook(file_path)
    sheet_names = sheets if sheets else [wb.active.title]

    for sheet_name in sheet_names:
        sheet = wb[sheet_name]
        cell = sheet[cell_ref]
        cell.border = border

    wb.save(file_path)
    print(f"La bordure a été appliquée à la cellule {cell_ref} dans les feuilles {sheet_names}.")


def apply_border_to_column(file_path: str, column_header: str, border: Border, sheets: list = None) -> None:
    """
    Applique une bordure à toutes les cellules d'une colonne spécifiée.

    Args:
        file_path (str): Le chemin vers le fichier Excel.
        column_header (str): L'entête de la colonne à modifier.
        border (Border): L'objet Border à appliquer.
        sheets (list): La liste des feuilles où effectuer la modification (par défaut, feuille active).
    """
    wb = openpyxl.load_workbook(file_path)
    sheet_names = sheets if sheets else [wb.active.title]

    for sheet_name in sheet_names:
        sheet = wb[sheet_name]
        column_index = get_column_index(sheet, column_header)

        for row in range(2, sheet.max_row + 1):
            cell = sheet.cell(row=row, column=column_index)
            cell.border = border

    wb.save(file_path)
    print(f"La bordure a été appliquée à la colonne '{column_header}' dans les feuilles {sheet_names}.")


def apply_border_to_row(file_path: str, row_number: int, border: Border, sheets: list = None) -> None:
    """
    Applique une bordure à toutes les cellules d'une ligne spécifiée.

    Args:
        file_path (str): Le chemin vers le fichier Excel.
        row_number (int): Le numéro de la ligne à modifier.
        border (Border): L'objet Border à appliquer.
        sheets (list): La liste des feuilles où effectuer la modification (par défaut, feuille active).
    """
    wb = openpyxl.load_workbook(file_path)
    sheet_names = sheets if sheets else [wb.active.title]

    for sheet_name in sheet_names:
        sheet = wb[sheet_name]

        for column in range(1, sheet.max_column + 1):
            cell = sheet.cell(row=row_number, column=column)
            cell.border = border

    wb.save(file_path)
    print(f"La bordure a été appliquée à la ligne {row_number} dans les feuilles {sheet_names}.")


def apply_border_to_range(file_path: str, start_column: str, end_column: str, start_row: int, end_row: int, 
                          border: Border, sheets: list = None) -> None:
    """
    Applique une bordure à une plage de colonnes et de lignes spécifiée.

    Args:
        file_path (str): Le chemin vers le fichier Excel.
        start_column (str): La première colonne de la plage (ex: 'A').
        end_column (str): La dernière colonne de la plage (ex: 'G').
        start_row (int): La première ligne de la plage.
        end_row (int): La dernière ligne de la plage.
        border (Border): L'objet Border à appliquer.
        sheets (list): La liste des feuilles où effectuer la modification (par défaut, feuille active).
    """
    wb = openpyxl.load_workbook(file_path)
    sheet_names = sheets if sheets else [wb.active.title]

    for sheet_name in sheet_names:
        sheet = wb[sheet_name]
        start_col_idx = openpyxl.utils.column_index_from_string(start_column)
        end_col_idx = openpyxl.utils.column_index_from_string(end_column)

        for col_idx in range(start_col_idx, end_col_idx + 1):
            for row in range(start_row, end_row + 1):
                cell = sheet.cell(row=row, column=col_idx)
                cell.border = border

    wb.save(file_path)
    print(f"La bordure a été appliquée de {start_column}{start_row} à {end_column}{end_row} dans les feuilles {sheet_names}.")


def get_column_index(sheet, column_header: str) -> int:
    """
    Retourne l'index de la colonne pour l'entête donné.

    Args:
        sheet (Worksheet): La feuille Excel à analyser.
        column_header (str): L'entête de la colonne.

    Returns:
        int: L'index de la colonne correspondant à l'entête.
    """
    for cell in sheet[1]:
        if cell.value == column_header:
            return cell.column
    raise ValueError(f"Colonne avec l'entête '{column_header}' non trouvée.")
