import pytest
import openpyxl
from openpyxl.styles import Border, Side
from excel_border_manager import (
    create_border, apply_border_to_cell, apply_border_to_column, 
    apply_border_to_row, apply_border_to_range
)

# Helper function to get the border style of a cell
def get_cell_border_style(file_path, sheet_name, cell_ref):
    """Récupère le style de la bordure d'une cellule."""
    wb = openpyxl.load_workbook(file_path)
    sheet = wb[sheet_name]
    cell = sheet[cell_ref]
    return cell.border

@pytest.fixture
def excel_file(tmp_path):
    """Crée un fichier Excel temporaire pour les tests."""
    file_path = tmp_path / "test_workbook.xlsx"
    wb = openpyxl.Workbook()

    # Feuille 1 avec des données de test
    sheet1 = wb.active
    sheet1.title = "Feuille1"
    sheet1.append(["Col1", "Col2", "Col3", "Col4", "Col5", "Col6", "Col7"])
    sheet1.append([1, 2, 3, 4, 5, 6, 7])
    sheet1.append([8, 9, 10, 11, 12, 13, 14])
    sheet1.append([15, 16, 17, 18, 19, 20, 21])

    wb.save(file_path)
    return file_path


def test_create_border():
    """Test pour vérifier la création d'une bordure."""
    border = create_border(style="solid", color="FF0000")  # Bordure rouge solide
    assert isinstance(border, Border)
    assert border.left.style == "solid"
    assert border.left.color.rgb == "00FF0000"  # Couleur rouge en hexadécimal


def test_apply_border_to_cell(excel_file):
    """Test pour appliquer une bordure à une cellule spécifique."""
    border = create_border(style="solid", color="000000")  # Bordure noire solide
    apply_border_to_cell(excel_file, "A2", border, sheets=["Feuille1"])
    cell_border = get_cell_border_style(excel_file, "Feuille1", "A2")
    assert cell_border.left.style == "solid"
    assert cell_border.left.color.rgb == "00000000"


def test_apply_border_to_column(excel_file):
    """Test pour appliquer une bordure à une colonne entière."""
    border = create_border(style="dashed", color="0000FF")  # Bordure bleue pointillée
    apply_border_to_column(excel_file, "Col3", border, sheets=["Feuille1"])
    cell_border_2 = get_cell_border_style(excel_file, "Feuille1", "C2")
    cell_border_3 = get_cell_border_style(excel_file, "Feuille1", "C3")
    assert cell_border_2.left.style == "dashed"
    assert cell_border_2.left.color.rgb == "000000FF"
    assert cell_border_3.left.style == "dashed"
    assert cell_border_3.left.color.rgb == "000000FF"


def test_apply_border_to_row(excel_file):
    """Test pour appliquer une bordure à une ligne entière."""
    border = create_border(style="double", color="FF00FF")  # Bordure double rose
    apply_border_to_row(excel_file, 2, border, sheets=["Feuille1"])
    cell_border_1 = get_cell_border_style(excel_file, "Feuille1", "A2")
    cell_border_7 = get_cell_border_style(excel_file, "Feuille1", "G2")
    assert cell_border_1.left.style == "double"
    assert cell_border_1.left.color.rgb == "00FF00FF"
    assert cell_border_7.left.style == "double"
    assert cell_border_7.left.color.rgb == "00FF00FF"


def test_apply_border_to_range(excel_file):
    """Test pour appliquer une bordure à une plage de colonnes et de lignes."""
    border = create_border(style="thick", color="00FF00")  # Bordure verte épaisse
    apply_border_to_range(excel_file, "A", "C", 2, 3, border, sheets=["Feuille1"])
    cell_border_1 = get_cell_border_style(excel_file, "Feuille1", "A2")
    cell_border_2 = get_cell_border_style(excel_file, "Feuille1", "B3")
    assert cell_border_1.left.style == "thick"
    assert cell_border_1.left.color.rgb == "0000FF00"
    assert cell_border_2.left.style == "thick"
    assert cell_border_2.left.color.rgb == "0000FF00"
