import pytest
import openpyxl
from datetime import datetime
from your_module_name import (
    check_column_format, check_numeric_column, check_date_column,
    check_row_format, check_numeric_row, check_date_row,
    check_cell_format, check_numeric_cell, check_date_cell
)

# Fixtures pour créer des fichiers Excel temporaires pour les tests
@pytest.fixture
def excel_file(tmp_path):
    """Crée un fichier Excel temporaire avec différentes feuilles pour les tests."""
    file_path = tmp_path / "test_workbook.xlsx"
    wb = openpyxl.Workbook()

    # Feuille 1
    sheet1 = wb.active
    sheet1.title = "Feuille1"
    sheet1.append(["TRAD", "pouet", "Date"])
    sheet1.append(["TR1234567890123", 123, datetime(2023, 9, 23)])
    sheet1.append(["TR0987654321098", 456, datetime(2024, 1, 1)])
    sheet1.append(["invalid_format", "not_a_number", "invalid_date"])

    # Feuille 2
    sheet2 = wb.create_sheet(title="Feuille2")
    sheet2.append(["TRAD", "pouet", "Date"])
    sheet2.append(["TR9999999999999", 789, datetime(2025, 5, 5)])

    wb.save(file_path)
    return file_path


# Tests pour les colonnes

def test_check_column_format_valid(excel_file):
    """Test vérifiant que les cellules TRAD sont valides."""
    check_column_format(excel_file, "TRAD", sheets=["Feuille1"])


def test_check_column_format_invalid(excel_file):
    """Test vérifiant que certaines cellules TRAD ont un format invalide."""
    with pytest.raises(ValueError):
        check_column_format(excel_file, "TRAD", sheets=["Feuille1"])


def test_check_numeric_column_valid(excel_file):
    """Test vérifiant que les cellules pouet sont toutes numériques."""
    check_numeric_column(excel_file, "pouet", sheets=["Feuille1"])


def test_check_numeric_column_invalid(excel_file):
    """Test vérifiant que certaines cellules pouet ne sont pas numériques."""
    with pytest.raises(ValueError):
        check_numeric_column(excel_file, "pouet", sheets=["Feuille1"])


def test_check_date_column_valid(excel_file):
    """Test vérifiant que les cellules Date sont toutes valides."""
    check_date_column(excel_file, "Date", sheets=["Feuille1"])


def test_check_date_column_invalid(excel_file):
    """Test vérifiant que certaines cellules Date ne sont pas valides."""
    with pytest.raises(ValueError):
        check_date_column(excel_file, "Date", sheets=["Feuille1"])


# Tests pour les lignes

def test_check_row_format_valid(excel_file):
    """Test vérifiant qu'une ligne spécifique respecte le format TRAD."""
    check_row_format(excel_file, 2, "TRAD", sheets=["Feuille1"])


def test_check_row_format_invalid(excel_file):
    """Test vérifiant qu'une ligne ne respecte pas le format TRAD."""
    with pytest.raises(ValueError):
        check_row_format(excel_file, 4, "TRAD", sheets=["Feuille1"])


def test_check_numeric_row_valid(excel_file):
    """Test vérifiant qu'une ligne spécifique contient un nombre dans pouet."""
    check_numeric_row(excel_file, 2, "pouet", sheets=["Feuille1"])


def test_check_numeric_row_invalid(excel_file):
    """Test vérifiant qu'une ligne ne contient pas de valeur numérique."""
    with pytest.raises(ValueError):
        check_numeric_row(excel_file, 4, "pouet", sheets=["Feuille1"])


def test_check_date_row_valid(excel_file):
    """Test vérifiant qu'une ligne spécifique contient une date valide."""
    check_date_row(excel_file, 2, "Date", sheets=["Feuille1"])


def test_check_date_row_invalid(excel_file):
    """Test vérifiant qu'une ligne ne contient pas de date valide."""
    with pytest.raises(ValueError):
        check_date_row(excel_file, 4, "Date", sheets=["Feuille1"])


# Tests pour les cellules

def test_check_cell_format_valid(excel_file):
    """Test vérifiant qu'une cellule spécifique respecte le format TRAD."""
    check_cell_format(excel_file, "A2", sheets=["Feuille1"])


def test_check_cell_format_invalid(excel_file):
    """Test vérifiant qu'une cellule spécifique ne respecte pas le format TRAD."""
    with pytest.raises(ValueError):
        check_cell_format(excel_file, "A4", sheets=["Feuille1"])


def test_check_numeric_cell_valid(excel_file):
    """Test vérifiant qu'une cellule spécifique contient une valeur numérique."""
    check_numeric_cell(excel_file, "B2", sheets=["Feuille1"])


def test_check_numeric_cell_invalid(excel_file):
    """Test vérifiant qu'une cellule spécifique ne contient pas de valeur numérique."""
    with pytest.raises(ValueError):
        check_numeric_cell(excel_file, "B4", sheets=["Feuille1"])


def test_check_date_cell_valid(excel_file):
    """Test vérifiant qu'une cellule spécifique contient une date valide."""
    check_date_cell(excel_file, "C2", sheets=["Feuille1"])


def test_check_date_cell_invalid(excel_file):
    """Test vérifiant qu'une cellule spécifique ne contient pas de date valide."""
    with pytest.raises(ValueError):
        check_date_cell(excel_file, "C4", sheets=["Feuille1"])
