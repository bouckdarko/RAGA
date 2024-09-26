"""
Module excel_color_manager

Ce module permet de gérer les couleurs de fond et les couleurs de texte dans les cellules d'un fichier Excel, 
ainsi que d'appliquer des styles de texte tels que gras, italique, souligné et barré. 
Il offre également des fonctionnalités pour appliquer des couleurs aux lignes, colonnes et plages de cellules 
et inclut une fonction utilitaire pour retrouver l'index d'une colonne à partir de son entête.

Fonctionnalités :
----------------
1. Appliquer des couleurs de fond aux cellules.
2. Appliquer des couleurs de fond aux colonnes, lignes ou plages de cellules.
3. Appliquer des couleurs de texte.
4. Appliquer des formats de texte (gras, italique, souligné, barré) aux cellules, colonnes, lignes ou plages.

Auteur : Bouck DARKO
Version : 1.0
"""

import openpyxl
from openpyxl.styles import PatternFill, Font


def apply_color_to_cell(file_path: str, cell_ref: str, color: str, sheets: list = None) -> None:
    """
    Applique une couleur de fond à une cellule spécifique.

    Args:
        file_path (str): Le chemin vers le fichier Excel.
        cell_ref (str): La référence de la cellule (ex: 'A1').
        color (str): La couleur au format hexadécimal (ex: 'FFFF00' pour jaune).
        sheets (list): La liste des feuilles où effectuer la modification (par défaut, feuille active).
    """
    wb = openpyxl.load_workbook(file_path)
    sheet_names = sheets if sheets else [wb.active.title]

    fill = PatternFill(start_color=color, end_color=color, fill_type="solid")

    for sheet_name in sheet_names:
        sheet = wb[sheet_name]
        cell = sheet[cell_ref]
        cell.fill = fill

    wb.save(file_path)
    print(f"La cellule {cell_ref} dans les feuilles {sheet_names} a été colorée en {color}.")


def apply_color_to_column(file_path: str, column_header: str, color: str, only_non_empty: bool = True, sheets: list = None) -> None:
    """
    Applique une couleur de fond à toutes les cellules d'une colonne spécifiée.

    Args:
        file_path (str): Le chemin vers le fichier Excel.
        column_header (str): L'entête de la colonne à colorer.
        color (str): La couleur au format hexadécimal (ex: 'FFFF00' pour jaune).
        only_non_empty (bool): Si True, la couleur est appliquée uniquement aux cellules non vides.
        sheets (list): La liste des feuilles où effectuer la modification (par défaut, feuille active).
    """
    wb = openpyxl.load_workbook(file_path)
    sheet_names = sheets if sheets else [wb.active.title]

    fill = PatternFill(start_color=color, end_color=color, fill_type="solid")

    for sheet_name in sheet_names:
        sheet = wb[sheet_name]
        column_index = get_column_index(sheet, column_header)

        for row in range(2, sheet.max_row + 1):
            cell = sheet.cell(row=row, column=column_index)
            if not only_non_empty or (cell.value and str(cell.value).strip()):
                cell.fill = fill

    wb.save(file_path)
    print(f"Les cellules de la colonne '{column_header}' dans les feuilles {sheet_names} ont été colorées en {color}.")


def apply_color_to_row(file_path: str, row_number: int, color: str, only_non_empty: bool = True, sheets: list = None) -> None:
    """
    Applique une couleur de fond à toutes les cellules d'une ligne spécifiée.

    Args:
        file_path (str): Le chemin vers le fichier Excel.
        row_number (int): Le numéro de la ligne à colorer.
        color (str): La couleur au format hexadécimal (ex: 'FFFF00' pour jaune).
        only_non_empty (bool): Si True, la couleur est appliquée uniquement aux cellules non vides.
        sheets (list): La liste des feuilles où effectuer la modification (par défaut, feuille active).
    """
    wb = openpyxl.load_workbook(file_path)
    sheet_names = sheets if sheets else [wb.active.title]

    fill = PatternFill(start_color=color, end_color=color, fill_type="solid")

    for sheet_name in sheet_names:
        sheet = wb[sheet_name]

        for column in range(1, sheet.max_column + 1):
            cell = sheet.cell(row=row_number, column=column)
            if not only_non_empty or (cell.value and str(cell.value).strip()):
                cell.fill = fill

    wb.save(file_path)
    print(f"Les cellules de la ligne {row_number} dans les feuilles {sheet_names} ont été colorées en {color}.")


def apply_color_to_range(file_path: str, start_column: str, end_column: str, color: str, 
                         start_row: int, end_row: int, only_non_empty: bool = True, sheets: list = None) -> None:
    """
    Applique une couleur de fond à une plage de colonnes et de lignes spécifiée.

    Args:
        file_path (str): Le chemin vers le fichier Excel.
        start_column (str): La première colonne de la plage (ex: 'A').
        end_column (str): La dernière colonne de la plage (ex: 'G').
        color (str): La couleur au format hexadécimal (ex: 'FFFF00' pour jaune).
        start_row (int): La première ligne de la plage.
        end_row (int): La dernière ligne de la plage.
        only_non_empty (bool): Si True, la couleur est appliquée uniquement aux cellules non vides.
        sheets (list): La liste des feuilles où effectuer la modification (par défaut, feuille active).
    """
    wb = openpyxl.load_workbook(file_path)
    sheet_names = sheets if sheets else [wb.active.title]

    fill = PatternFill(start_color=color, end_color=color, fill_type="solid")

    for sheet_name in sheet_names:
        sheet = wb[sheet_name]
        start_col_idx = openpyxl.utils.column_index_from_string(start_column)
        end_col_idx = openpyxl.utils.column_index_from_string(end_column)

        for col_idx in range(start_col_idx, end_col_idx + 1):
            for row in range(start_row, end_row + 1):
                cell = sheet.cell(row=row, column=col_idx)
                if not only_non_empty or (cell.value and str(cell.value).strip()):
                    cell.fill = fill

    wb.save(file_path)
    print(f"Les cellules de {start_column}{start_row} à {end_column}{end_row} dans les feuilles {sheet_names} ont été colorées en {color}.")


def apply_text_color_to_column(file_path: str, column_header: str, font_color: str, sheets: list = None) -> None:
    """
    Applique une couleur au texte dans une colonne entière.

    Args:
        file_path (str): Le chemin vers le fichier Excel.
        column_header (str): L'entête de la colonne à colorer.
        font_color (str): La couleur du texte au format hexadécimal (ex: 'FF0000' pour rouge).
        sheets (list): La liste des feuilles où effectuer la modification (par défaut, feuille active).
    """
    wb = openpyxl.load_workbook(file_path)
    sheet_names = sheets if sheets else [wb.active.title]

    for sheet_name in sheet_names:
        sheet = wb[sheet_name]
        column_index = get_column_index(sheet, column_header)

        for row in range(2, sheet.max_row + 1):
            cell = sheet.cell(row=row, column=column_index)
            cell.font = Font(color=font_color)

    wb.save(file_path)
    print(f"Le texte de la colonne '{column_header}' dans les feuilles {sheet_names} a été modifié en couleur {font_color}.")


def apply_text_color_to_row(file_path: str, row_number: int, font_color: str, sheets: list = None) -> None:
    """
    Applique une couleur au texte dans une ligne entière.

    Args:
        file_path (str): Le chemin vers le fichier Excel.
        row_number (int): Le numéro de la ligne à colorer.
        font_color (str): La couleur du texte au format hexadécimal (ex: 'FF0000' pour rouge).
        sheets (list): La liste des feuilles où effectuer la modification (par défaut, feuille active).
    """
    wb = openpyxl.load_workbook(file_path)
    sheet_names = sheets if sheets else [wb.active.title]

    for sheet_name in sheet_names:
        sheet = wb[sheet_name]

        for column in range(1, sheet.max_column + 1):
            cell = sheet.cell(row=row_number, column=column)
            cell.font = Font(color=font_color)

    wb.save(file_path)
    print(f"Le texte de la ligne {row_number} dans les feuilles {sheet_names} a été modifié en couleur {font_color}.")


def apply_text_color_to_range(file_path: str, start_column: str, end_column: str, font_color: str, 
                              start_row: int, end_row: int, sheets: list = None) -> None:
    """
    Applique une couleur au texte dans une plage de colonnes et de lignes spécifiée.

    Args:
        file_path (str): Le chemin vers le fichier Excel.
        start_column (str): La première colonne de la plage (ex: 'A').
        end_column (str): La dernière colonne de la plage (ex: 'G').
        font_color (str): La couleur du texte au format hexadécimal (ex: 'FF0000' pour rouge).
        start_row (int): La première ligne de la plage.
        end_row (int): La dernière ligne de la plage.
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
                cell.font = Font(color=font_color)

    wb.save(file_path)
    print(f"Le texte des cellules de {start_column}{start_row} à {end_column}{end_row} dans les feuilles {sheet_names} a été modifié en couleur {font_color}.")


def apply_text_format_to_column(file_path: str, column_header: str, bold: bool = False, italic: bool = False, 
                                underline: bool = False, strike: bool = False, sheets: list = None) -> None:
    """
    Applique des formats de texte à une colonne entière (gras, italique, souligné, barré).

    Args:
        file_path (str): Le chemin vers le fichier Excel.
        column_header (str): L'entête de la colonne à formater.
        bold (bool): Si True, applique le gras au texte.
        italic (bool): Si True, applique l'italique au texte.
        underline (bool): Si True, applique le soulignement au texte.
        strike (bool): Si True, applique le texte barré.
        sheets (list): La liste des feuilles où effectuer la modification (par défaut, feuille active).
    """
    wb = openpyxl.load_workbook(file_path)
    sheet_names = sheets if sheets else [wb.active.title]

    for sheet_name in sheet_names:
        sheet = wb[sheet_name]
        column_index = get_column_index(sheet, column_header)

        for row in range(2, sheet.max_row + 1):
            cell = sheet.cell(row=row, column=column_index)
            cell.font = Font(bold=bold, italic=italic, underline="single" if underline else None, strike=strike)

    wb.save(file_path)
    print(f"Le format du texte de la colonne '{column_header}' dans les feuilles {sheet_names} a été modifié (gras={bold}, italique={italic}, souligné={underline}, barré={strike}).")


def apply_text_format_to_row(file_path: str, row_number: int, bold: bool = False, italic: bool = False, 
                             underline: bool = False, strike: bool = False, sheets: list = None) -> None:
    """
    Applique des formats de texte à une ligne entière (gras, italique, souligné, barré).

    Args:
        file_path (str): Le chemin vers le fichier Excel.
        row_number (int): Le numéro de la ligne à formater.
        bold (bool): Si True, applique le gras au texte.
        italic (bool): Si True, applique l'italique au texte.
        underline (bool): Si True, applique le soulignement au texte.
        strike (bool): Si True, applique le texte barré.
        sheets (list): La liste des feuilles où effectuer la modification (par défaut, feuille active).
    """
    wb = openpyxl.load_workbook(file_path)
    sheet_names = sheets if sheets else [wb.active.title]

    for sheet_name in sheet_names:
        sheet = wb[sheet_name]

        for column in range(1, sheet.max_column + 1):
            cell = sheet.cell(row=row_number, column=column)
            cell.font = Font(bold=bold, italic=italic, underline="single" if underline else None, strike=strike)

    wb.save(file_path)
    print(f"Le format du texte de la ligne {row_number} dans les feuilles {sheet_names} a été modifié (gras={bold}, italique={italic}, souligné={underline}, barré={strike}).")


def apply_text_format_to_range(file_path: str, start_column: str, end_column: str, bold: bool = False, italic: bool = False, 
                               underline: bool = False, strike: bool = False, start_row: int = None, end_row: int = None, sheets: list = None) -> None:
    """
    Applique des formats de texte à une plage de colonnes et de lignes spécifiée (gras, italique, souligné, barré).

    Args:
        file_path (str): Le chemin vers le fichier Excel.
        start_column (str): La première colonne de la plage (ex: 'A').
        end_column (str): La dernière colonne de la plage (ex: 'G').
        bold (bool): Si True, applique le gras au texte.
        italic (bool): Si True, applique l'italique au texte.
        underline (bool): Si True, applique le soulignement au texte.
        strike (bool): Si True, applique le texte barré.
        start_row (int): La première ligne de la plage.
        end_row (int): La dernière ligne de la plage.
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
                cell.font = Font(bold=bold, italic=italic, underline="single" if underline else None, strike=strike)

    wb.save(file_path)
    print(f"Le format du texte des cellules de {start_column}{start_row} à {end_column}{end_row} dans les feuilles {sheet_names} a été modifié (gras={bold}, italique={italic}, souligné={underline}, barré={strike}).")


# Fonction utilitaire pour trouver l'index d'une colonne à partir de l'entête

def get_column_index(sheet, column_header: str) -> int:
    """Retourne l'index de la colonne pour l'entête donné."""
    for cell in sheet[1]:
        if cell.value == column_header:
            return cell.column
    raise ValueError(f"Colonne avec l'entête '{column_header}' non trouvée.")
