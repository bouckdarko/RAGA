"""Cette fonction commence par obtenir le chemin d’exécution du script principal à l’aide de os.path.abspath(__file__). Ensuite, elle vérifie si le fichier spécifié existe dans ce répertoire. Si ce n’est pas le cas, elle parcourt récursivement les sous-dossiers pour trouver le fichier. Si le fichier est trouvé, elle retourne le chemin absolu ; sinon, elle renvoie None"""

import os

def chemin_absolu(nom_fichier: str) -> str:
    """
    Crée un chemin absolu en utilisant le chemin d'exécution du script principal et le nom du fichier spécifié.

    Args:
        nom_fichier (str): Le nom du fichier recherché.

    Returns:
        str: Le chemin absolu vers le fichier s'il est trouvé, sinon None.
    """
    # Chemin d'exécution du script principal
    chemin_execution = os.path.dirname(os.path.abspath(__file__))

    # Vérifier si le fichier existe dans le répertoire d'exécution
    chemin_fichier = os.path.join(chemin_execution, nom_fichier)
    if os.path.exists(chemin_fichier):
        return chemin_fichier

    # Chercher le fichier dans les sous-dossiers du dossier d'exécution
    for dossier, sous_dossiers, fichiers in os.walk(chemin_execution):
        if nom_fichier in fichiers:
            return os.path.join(dossier, nom_fichier)

    # Si le fichier n'est pas trouvé, retourner None
    return None

# Exemple d'utilisation
nom_fichier_specifie = "mon_fichier.txt"
chemin = chemin_absolu(nom_fichier_specifie)
if chemin:
    print(f"Chemin absolu vers '{nom_fichier_specifie}': {chemin}")
else:
    print(f"Le fichier '{nom_fichier_specifie}' n'a pas été trouvé.")
