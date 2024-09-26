import pytest
import openpyxl
from excel_text_manager import (
    add_text_to_cell, overwrite_cell, remove_text_from_cell,
    add_text_to_column, overwrite_column, remove_text_from_column,
    add_text_to_row, overwrite_row, remove_text_from_row
)


# Fixture pour créer un fichier Excel temporaire
@pytest.fixture
def excel_file(tmp_path):
    """Crée un fichier Excel temporaire avec différentes feuilles pour les tests."""
    file_path = tmp_path / "test_workbook.xlsx"
    wb = openpyxl.Workbook()

    # Feuille 1
    sheet1 = wb.active
    sheet1.title = "Feuille1"
    sheet1.append(["TRAD", "pouet", "Date"])
    sheet1.append(["TR1234567890123", "Valeur initiale", "2023-09-23"])
    sheet1.append(["TR0987654321098", "Texte existant", "2024-01-01"])

    # Feuille 2
    sheet2 = wb.create_sheet(title="Feuille2")
    sheet2.append(["TRAD", "pouet", "Date"])
    sheet2.append(["TR9999999999999", "Texte feuille 2", "2025-05-05"])

    wb.save(file_path)
    return file_path


# Helper function pour charger la feuille et vérifier la valeur d'une cellule
def get_cell_value(file_path, sheet_name, cell_ref):
    wb = openpyxl.load_workbook(file_path)
    sheet = wb[sheet_name]
    return sheet[cell_ref].value


# Tests pour les cellules

def test_add_text_to_cell_with_delimiter(excel_file):
    """Test pour ajouter du texte à une cellule existante avec un délimiteur."""
    add_text_to_cell(excel_file, "B2", " ajouté", delimiter=" / ", sheets=["Feuille1"])
    assert get_cell_value(excel_file, "Feuille1", "B2") == "Valeur initiale / ajouté"


def test_add_text_to_cell_no_delimiter(excel_file):
    """Test pour ajouter du texte à une cellule existante sans délimiteur."""
    add_text_to_cell(excel_file, "B2", " ajouté", delimiter="", sheets=["Feuille1"])
    assert get_cell_value(excel_file, "Feuille1", "B2") == "Valeur initiale ajouté"


def test_add_text_to_empty_cell(excel_file):
    """Test pour ajouter du texte à une cellule vide (pas de délimiteur nécessaire)."""
    add_text_to_cell(excel_file, "B3", "Texte ajouté", sheets=["Feuille1"])
    assert get_cell_value(excel_file, "Feuille1", "B3") == "Texte ajouté"


def test_overwrite_cell(excel_file):
    """Test pour écraser le texte d'une cellule spécifique."""
    overwrite_cell(excel_file, "B2", "Nouveau texte", sheets=["Feuille1"])
    assert get_cell_value(excel_file, "Feuille1", "B2") == "Nouveau texte"


def test_remove_text_from_cell(excel_file):
    """Test pour supprimer une portion de texte d'une cellule."""
    remove_text_from_cell(excel_file, "B2", "initiale", sheets=["Feuille1"])
    assert get_cell_value(excel_file, "Feuille1", "B2") == "Valeur "


# Tests pour les colonnes

def test_add_text_to_column_with_delimiter(excel_file):
    """Test pour ajouter du texte à une colonne entière avec un délimiteur."""
    add_text_to_column(excel_file, "pouet", " - suffixe", delimiter=" / ", sheets=["Feuille1"])
    assert get_cell_value(excel_file, "Feuille1", "B2") == "Valeur initiale / - suffixe"
    assert get_cell_value(excel_file, "Feuille1", "B3") == "Texte existant / - suffixe"


def test_overwrite_column(excel_file):
    """Test pour écraser le texte dans une colonne entière."""
    overwrite_column(excel_file, "pouet", "Colonne écrasée", sheets=["Feuille1"])
    assert get_cell_value(excel_file, "Feuille1", "B2") == "Colonne écrasée"
    assert get_cell_value(excel_file, "Feuille1", "B3") == "Colonne écrasée"


def test_remove_text_from_column(excel_file):
    """Test pour supprimer une portion de texte dans une colonne entière."""
    remove_text_from_column(excel_file, "pouet", "Texte", sheets=["Feuille1"])
    assert get_cell_value(excel_file, "Feuille1", "B3") == " existant"


# Tests pour les lignes

def test_add_text_to_row_with_delimiter(excel_file):
    """Test pour ajouter du texte à une ligne entière avec un délimiteur."""
    add_text_to_row(excel_file, 2, " - ligne", delimiter=" / ", sheets=["Feuille1"])
    assert get_cell_value(excel_file, "Feuille1", "A2") == "TR1234567890123 / - ligne"
    assert get_cell_value(excel_file, "Feuille1", "B2") == "Valeur initiale / - ligne"
    assert get_cell_value(excel_file, "Feuille1", "C2") == "2023-09-23 / - ligne"


def test_overwrite_row(excel_file):
    """Test pour écraser le texte dans une ligne entière."""
    overwrite_row(excel_file, 2, "Ligne écrasée", sheets=["Feuille1"])
    assert get_cell_value(excel_file, "Feuille1", "A2") == "Ligne écrasée"
    assert get_cell_value(excel_file, "Feuille1", "B2") == "Ligne écrasée"
    assert get_cell_value(excel_file, "Feuille1", "C2") == "Ligne écrasée"


def test_remove_text_from_row(excel_file):
    """Test pour supprimer une portion de texte dans une ligne entière."""
    remove_text_from_row(excel_file, 2, "Valeur", sheets=["Feuille1"])
    assert get_cell_value(excel_file, "Feuille1", "B2") == " initiale"
