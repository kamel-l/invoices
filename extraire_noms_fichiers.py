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
        for fichier in liste_fichiers:
            liste_invoice.append(fichier)
        print(liste_invoice)
            # print(f"- {fichier}")
    except Exception as e:
        print(f"Erreur : {e}")
        
        
import os
from pathlib import Path
def remplacer_symbole_direct(fichier_entree, fichier_sortie, ancien_symbole, nouveau_symbole):
    """
    Version alternative qui traite le fichier directement comme un fichier texte
    Utile pour des fichiers très grands ou des cas spéciaux
    """
    try:
        with open(fichier_entree, 'r', encoding='utf-8') as f_entree:
            contenu = f_entree.read()

        # Remplacer le symbole
        nouveau_contenu = contenu.replace(ancien_symbole, nouveau_symbole)

        with open(fichier_sortie, 'w', encoding='utf-8') as f_sortie:
            f_sortie.write(nouveau_contenu)

        print(f"Le remplacement a été effectué avec succès. Fichier sauvegardé : {fichier_sortie}")

    except Exception as e:
        print(f"Une erreur s'est produite : {str(e)}")
        

        
list_invoices = [nom_fichier for nom_fichier in os.listdir("invoice24") if
                     os.path.isfile(os.path.join("invoice24", nom_fichier))]

chemain24 = "invoice24"
for invoice in chemain24:
    print(invoice)
    remplacer_symbole_direct(invoice, 'invoices10.dat', '%20', '_')


 # Créer le dossier de destination s'il n'existe pas
download_path = Path("extrair_fichier")
download_path.mkdir(parents=True, exist_ok=True)
#remplacer_symbole_direct('invoice/10000.dat', 'invoices10.dat', '%20', '_')        