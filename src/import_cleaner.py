"""
Module import_cleaner

Ce module parcourt un répertoire de code source et remplace toutes les occurrences de
`from module import *` par des importations explicites. 

Il extrait les noms des fonctions, classes, ou variables disponibles dans les modules concernés
et les insère explicitement dans l'importation.

Note : Ce script suppose que les modules à nettoyer peuvent être importés sans effets secondaires.

Fonctionnalités :
----------------
1. Remplacer `from module import *` par des importations explicites.
2. Parcourir tous les fichiers Python dans un répertoire donné.

Auteur : Bouck DARKO
Version : 1.0
"""

import os
import ast
import importlib

def get_module_members(module_name):
    """
    Retourne une liste de tous les attributs publics (fonctions, classes, variables) 
    d'un module donné.
    
    Args:
        module_name (str): Le nom du module à analyser.
    
    Returns:
        list: Une liste des membres publics du module.
    """
    try:
        module = importlib.import_module(module_name)
        return [name for name in dir(module) if not name.startswith('_')]
    except ImportError:
        print(f"Impossible d'importer le module {module_name}.")
        return []

def replace_import_star(file_path):
    """
    Remplace toutes les occurrences de `from module import *` dans un fichier donné 
    par une liste explicite des importations.
    
    Args:
        file_path (str): Le chemin du fichier Python à analyser et modifier.
    """
    with open(file_path, 'r') as file:
        tree = ast.parse(file.read())

    updated_code = []
    updated = False

    for node in tree.body:
        if isinstance(node, ast.ImportFrom) and node.names[0].name == '*' and node.module:
            module_name = node.module
            members = get_module_members(module_name)
            if members:
                import_statement = f"from {module_name} import {', '.join(members)}\n"
                updated_code.append(import_statement)
                updated = True
            else:
                print(f"Impossible de trouver les membres de {module_name}. L'importation a été ignorée.")
        else:
            updated_code.append(ast.get_source_segment(open(file_path).read(), node))

    if updated:
        with open(file_path, 'w') as file:
            file.writelines(updated_code)
        print(f"Le fichier {file_path} a été mis à jour.")

def clean_directory(directory_path):
    """
    Parcourt tous les fichiers Python d'un répertoire et remplace les `from module import *`.
    
    Args:
        directory_path (str): Le chemin du répertoire à analyser.
    """
    for root, dirs, files in os.walk(directory_path):
        for file in files:
            if file.endswith('.py'):
                file_path = os.path.join(root, file)
                print(f"Analyse de {file_path}...")
                replace_import_star(file_path)

# Exemple d'utilisation
if __name__ == "__main__":
    directory = "chemin/vers/ton/projet"
    clean_directory(directory)
