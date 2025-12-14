import os


def extraire_noms_fichiers(chemin_dossier):
    """
    Extrait les noms des fichiers d'un dossier donné

    Args:
        chemin_dossier (str): Le chemin vers le dossier à analyser

    Returns:
        list: Liste des noms de fichiers
    """
    # Vérifie si le dossier existe
    if not os.path.exists(chemin_dossier):
        raise Exception("Le dossier spécifié n'existe pas")

    # Obtient la liste des fichiers
    fichiers = []

    # Parcourt le dossier
    for nom_fichier in os.listdir(chemin_dossier):
        # Vérifie si c'est un fichier (et non un sous-dossier)
        if os.path.isfile(os.path.join(chemin_dossier, nom_fichier)):
            fichiers.append(nom_fichier)

    return fichiers
chemin = "invoice"
liste_fichiers = extraire_noms_fichiers(chemin)

# Exemple d'utilisation
if __name__ == "__main__":
    # Remplacez ceci par le chemin de votre dossier
    liste_invoice = []
    chemin = "invoice"
    try:
        liste_fichiers = extraire_noms_fichiers(chemin)
        print("Liste des fichiers trouvés :")
        for fichier in liste_fichiers:
            liste_invoice.append(fichier)
        print(liste_invoice)
            # print(f"- {fichier}")
    except Exception as e:
        print(f"Erreur : {e}")