import pytest
import openpyxl
from excel_data_validation import (
    apply_list_validation,
    apply_numeric_validation,
    apply_date_validation,
    apply_text_length_validation
)

# Helper function to get the validation rules for a cell range
def get_validation_rules(file_path, sheet_name, cell_range):
    """Récupère les règles de validation pour une plage de cellules."""
    wb = openpyxl.load_workbook(file_path)
    sheet = wb[sheet_name]
    dv_list = sheet.data_validations.dataValidation
    rules = []
    for dv in dv_list:
        if dv.ranges.contains(cell_range):
            rules.append(dv)
    return rules


@pytest.fixture
def excel_file(tmp_path):
    """Crée un fichier Excel temporaire pour les tests."""
    file_path = tmp_path / "test_workbook.xlsx"
    wb = openpyxl.Workbook()

    # Feuille 1 avec des données de test
    sheet1 = wb.active
    sheet1.title = "Feuille1"
    sheet1.append(["Col1", "Col2", "Col3", "Col4"])
    sheet1.append(["", "", "", ""])
    sheet1.append(["", "", "", ""])

    wb.save(file_path)
    return file_path


def test_apply_list_validation(excel_file):
    """Test pour valider une liste de valeurs autorisées dans une plage de cellules."""
    apply_list_validation(excel_file, "A2:A4", ["Option1", "Option2", "Option3"], sheets=["Feuille1"])
    rules = get_validation_rules(excel_file, "Feuille1", "A2:A4")
    assert len(rules) == 1
    assert rules[0].formula1 == '"Option1,Option2,Option3"'


def test_apply_numeric_validation(excel_file):
    """Test pour valider des valeurs numériques entre un minimum et un maximum dans une plage de cellules."""
    apply_numeric_validation(excel_file, "B2:B4", 0, 100, sheets=["Feuille1"])
    rules = get_validation_rules(excel_file, "Feuille1", "B2:B4")
    assert len(rules) == 1
    assert rules[0].formula1 == '0'
    assert rules[0].formula2 == '100'
    assert rules[0].type == 'decimal'


def test_apply_date_validation(excel_file):
    """Test pour valider une plage de dates autorisées dans une plage de cellules."""
    apply_date_validation(excel_file, "C2:C4", "2023-01-01", "2023-12-31", sheets=["Feuille1"])
    rules = get_validation_rules(excel_file, "Feuille1", "C2:C4")
    assert len(rules) == 1
    assert rules[0].formula1 == '"2023-01-01"'
    assert rules[0].formula2 == '"2023-12-31"'
    assert rules[0].type == 'date'


def test_apply_text_length_validation(excel_file):
    """Test pour valider la longueur maximale de texte dans une plage de cellules."""
    apply_text_length_validation(excel_file, "D2:D4", 20, sheets=["Feuille1"])
    rules = get_validation_rules(excel_file, "Feuille1", "D2:D4")
    assert len(rules) == 1
    assert rules[0].formula1 == '20'
    assert rules[0].type == 'textLength'
