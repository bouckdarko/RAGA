"""
Module excel_sheet_manager

Ce module permet de gérer les feuilles d'un fichier Excel en utilisant openpyxl. 
Il fournit des fonctions pour créer, supprimer, copier, réorganiser, verrouiller 
et gérer les propriétés des feuilles, comme la couleur d'onglet et la feuille active.

Fonctionnalités :
----------------
1. Créer, supprimer et copier des feuilles.
2. Réorganiser l'ordre des feuilles dans un fichier Excel.
3. Verrouiller et déverrouiller des feuilles.
4. Changer la couleur de l'onglet d'une feuille.
5. Définir la feuille active dans un fichier Excel.
6. Obtenir la liste et un dictionnaire des feuilles du fichier.

Auteur : Bouck DARKO
Version : 1.0
"""

import openpyxl
from openpyxl.utils import sheet as sheet_utils


def create_sheet(file_path: str, sheet_name: str, position: int = None) -> None:
    """
    Crée une nouvelle feuille dans un fichier Excel.

    Args:
        file_path (str): Le chemin vers le fichier Excel.
        sheet_name (str): Le nom de la nouvelle feuille.
        position (int): La position de la nouvelle feuille (par défaut, dernière position).
    """
    wb = openpyxl.load_workbook(file_path)
    wb.create_sheet(sheet_name, position=position)
    wb.save(file_path)
    print(f"La feuille '{sheet_name}' a été créée dans le fichier {file_path}.")


def delete_sheet(file_path: str, sheet_name: str) -> None:
    """
    Supprime une feuille d'un fichier Excel.

    Args:
        file_path (str): Le chemin vers le fichier Excel.
        sheet_name (str): Le nom de la feuille à supprimer.
    """
    wb = openpyxl.load_workbook(file_path)
    if sheet_name in wb.sheetnames:
        del wb[sheet_name]
        wb.save(file_path)
        print(f"La feuille '{sheet_name}' a été supprimée du fichier {file_path}.")
    else:
        print(f"Feuille '{sheet_name}' non trouvée dans le fichier.")


def set_active_sheet(file_path: str, sheet_name: str) -> None:
    """
    Définit la feuille active dans un fichier Excel.

    Args:
        file_path (str): Le chemin vers le fichier Excel.
        sheet_name (str): Le nom de la feuille à définir comme active.
    """
    wb = openpyxl.load_workbook(file_path)
    if sheet_name in wb.sheetnames:
        wb.active = wb.sheetnames.index(sheet_name)
        wb.save(file_path)
        print(f"La feuille '{sheet_name}' est maintenant active dans le fichier {file_path}.")
    else:
        print(f"Feuille '{sheet_name}' non trouvée dans le fichier.")


def change_sheet_color(file_path: str, sheet_name: str, color: str) -> None:
    """
    Change la couleur de l'onglet d'une feuille.

    Args:
        file_path (str): Le chemin vers le fichier Excel.
        sheet_name (str): Le nom de la feuille dont l'onglet doit changer de couleur.
        color (str): La couleur de l'onglet (en hexadécimal).
    """
    wb = openpyxl.load_workbook(file_path)
    if sheet_name in wb.sheetnames:
        sheet = wb[sheet_name]
        sheet.sheet_properties.tabColor = color
        wb.save(file_path)
        print(f"La couleur de l'onglet de la feuille '{sheet_name}' a été changée en {color}.")
    else:
        print(f"Feuille '{sheet_name}' non trouvée dans le fichier.")


def lock_sheet(file_path: str, sheet_name: str, password: str = None) -> None:
    """
    Verrouille une feuille avec un mot de passe optionnel.

    Args:
        file_path (str): Le chemin vers le fichier Excel.
        sheet_name (str): Le nom de la feuille à verrouiller.
        password (str): Le mot de passe pour verrouiller la feuille (optionnel).
    """
    wb = openpyxl.load_workbook(file_path)
    if sheet_name in wb.sheetnames:
        sheet = wb[sheet_name]
        sheet.protection.sheet = True
        if password:
            sheet.protection.set_password(password)
        wb.save(file_path)
        print(f"La feuille '{sheet_name}' a été verrouillée.")
    else:
        print(f"Feuille '{sheet_name}' non trouvée dans le fichier.")


def unlock_sheet(file_path: str, sheet_name: str, password: str = None) -> None:
    """
    Déverrouille une feuille avec un mot de passe si nécessaire.

    Args:
        file_path (str): Le chemin vers le fichier Excel.
        sheet_name (str): Le nom de la feuille à déverrouiller.
        password (str): Le mot de passe pour déverrouiller la feuille (optionnel).
    """
    wb = openpyxl.load_workbook(file_path)
    if sheet_name in wb.sheetnames:
        sheet = wb[sheet_name]
        sheet.protection.sheet = False
        wb.save(file_path)
        print(f"La feuille '{sheet_name}' a été déverrouillée.")
    else:
        print(f"Feuille '{sheet_name}' non trouvée dans le fichier.")


def copy_sheet(file_path: str, sheet_name: str, new_sheet_name: str) -> None:
    """
    Copie une feuille dans le même fichier Excel.

    Args:
        file_path (str): Le chemin vers le fichier Excel.
        sheet_name (str): Le nom de la feuille à copier.
        new_sheet_name (str): Le nom de la nouvelle feuille créée.
    """
    wb = openpyxl.load_workbook(file_path)
    if sheet_name in wb.sheetnames:
        source = wb[sheet_name]
        wb.copy_worksheet(source).title = new_sheet_name
        wb.save(file_path)
        print(f"La feuille '{sheet_name}' a été copiée sous le nom '{new_sheet_name}'.")
    else:
        print(f"Feuille '{sheet_name}' non trouvée dans le fichier.")


def reorder_sheets(file_path: str, order: str = "custom", sheet_order: list = None, 
                   sheet_name: str = None, direction: str = None) -> None:
    """
    Réorganise les feuilles d'un fichier Excel selon un ordre spécifié :
    - Par ordre alphabétique ascendant/descendant.
    - Par un ordre personnalisé.
    - En déplaçant une feuille d'une position à gauche ou à droite.

    Args:
        file_path (str): Le chemin vers le fichier Excel.
        order (str): Le type de réorganisation ("asc", "desc", "custom", "move").
        sheet_order (list): Une liste de noms de feuilles pour l'ordre personnalisé (requis si order="custom").
        sheet_name (str): Le nom de la feuille à déplacer (requis si order="move").
        direction (str): La direction pour déplacer la feuille ("left", "right") (requis si order="move").
    """
    wb = openpyxl.load_workbook(file_path)
    if order == "asc":
        wb._sheets.sort(key=lambda ws: ws.title)
    elif order == "desc":
        wb._sheets.sort(key=lambda ws: ws.title, reverse=True)
    elif order == "custom" and sheet_order:
        wb._sheets.sort(key=lambda ws: sheet_order.index(ws.title) if ws.title in sheet_order else len(sheet_order))
    elif order == "move" and sheet_name and direction:
        sheets = wb.sheetnames
        current_index = sheets.index(sheet_name)
        if direction == "left" and current_index > 0:
            wb._sheets.insert(current_index - 1, wb._sheets.pop(current_index))
        elif direction == "right" and current_index < len(sheets) - 1:
            wb._sheets.insert(current_index + 1, wb._sheets.pop(current_index))
        else:
            print(f"Impossible de déplacer la feuille '{sheet_name}' vers la direction {direction}.")
    
    wb.save(file_path)
    print(f"Les feuilles du fichier {file_path} ont été réorganisées selon l'ordre '{order}'.")


def get_sheet_list(file_path: str) -> list:
    """
    Retourne une liste des feuilles présentes dans un fichier Excel.

    Args:
        file_path (str): Le chemin vers le fichier Excel.

    Returns:
        list: Une liste contenant les noms des feuilles du fichier.
    """
    wb = openpyxl.load_workbook(file_path)
    return wb.sheetnames


def get_sheet_dict(file_path: str) -> dict:
    """
    Retourne un dictionnaire avec les feuilles du fichier Excel, où les clés sont
    des nombres (1, 2, 3, etc.) correspondant à la position des feuilles.

    Args:
        file_path (str): Le chemin vers le fichier Excel.

    Returns:
        dict: Un dictionnaire avec les positions (1, 2, etc.) comme clés et les noms des feuilles comme valeurs.
    """
    wb = openpyxl.load_workbook(file_path)
    return {i+1: sheet for i, sheet in enumerate(wb.sheetnames)}
