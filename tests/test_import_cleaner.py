import os
import pytest
from import_cleaner import get_module_members, replace_import_star, clean_directory

# Création d'un répertoire temporaire pour les tests
@pytest.fixture
def temp_dir(tmp_path):
    """Crée un répertoire temporaire avec des fichiers Python pour les tests."""
    temp_directory = tmp_path / "test_project"
    temp_directory.mkdir()

    # Création d'un module factice avec des fonctions et classes
    module_file = temp_directory / "fake_module.py"
    module_file.write_text("""
def foo():
    pass

def bar():
    pass

class MyClass:
    pass
""")

    # Création d'un fichier avec un `from fake_module import *`
    test_file = temp_directory / "test_file.py"
    test_file.write_text("""
from fake_module import *

def use_functions():
    foo()
    bar()
    obj = MyClass()
""")

    return temp_directory

# Test de la fonction `get_module_members`
def test_get_module_members(temp_dir):
    """Test pour s'assurer que la fonction retourne bien les membres publics du module."""
    module_members = get_module_members("fake_module")
    assert "foo" in module_members
    assert "bar" in module_members
    assert "MyClass" in module_members

# Test pour remplacer `from module import *` dans un fichier
def test_replace_import_star(temp_dir):
    """Test pour vérifier que `replace_import_star` remplace correctement l'import * par une importation explicite."""
    test_file_path = temp_dir / "test_file.py"

    # Remplacer `from fake_module import *` par une liste explicite d'importations
    replace_import_star(test_file_path)

    # Vérifier que le fichier a été modifié
    with open(test_file_path, 'r') as test_file:
        content = test_file.read()

    # Les membres de fake_module devraient être importés explicitement
    assert "from fake_module import foo, bar, MyClass" in content

# Test pour vérifier qu'aucune modification n'est apportée si `import *` n'est pas présent
def test_no_import_star(temp_dir):
    """Test pour vérifier que le fichier n'est pas modifié si aucun `import *` n'est présent."""
    non_star_import_file = temp_dir / "no_star_import.py"
    non_star_import_file.write_text("""
from fake_module import foo, bar

def use_foo():
    foo()
""")

    # Remplacer les imports globaux dans ce fichier (il n'y en a pas)
    replace_import_star(non_star_import_file)

    # Vérifier que le contenu du fichier n'a pas été modifié
    with open(non_star_import_file, 'r') as f:
        content = f.read()

    assert "from fake_module import foo, bar" in content  # L'importation doit rester intacte

# Test de la fonction `clean_directory`
def test_clean_directory(temp_dir):
    """Test pour vérifier que `clean_directory` nettoie correctement tous les fichiers d'un répertoire."""
    clean_directory(temp_dir)

    # Vérifier que le fichier test_file.py a été mis à jour avec une importation explicite
    test_file_path = temp_dir / "test_file.py"
    with open(test_file_path, 'r') as test_file:
        content = test_file.read()

    assert "from fake_module import foo, bar, MyClass" in content

# Test pour vérifier la gestion des erreurs d'importation
def test_get_module_members_error_handling():
    """Test pour s'assurer que la fonction gère correctement les erreurs d'importation."""
    module_members = get_module_members("non_existent_module")
    assert module_members == []  # Si le module n'existe pas, il devrait retourner une liste vide
