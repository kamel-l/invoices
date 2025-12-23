import os
from datfilereader import DatFileReader
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


# Créer le dossier de destination s'il n'existe pas
download_path = Path("invoices_downloaded")
download_path.mkdir(parents=True, exist_ok=True)

list_invoices = [nom_fichier for nom_fichier in os.listdir("invoices") 
                 if os.path.isfile(os.path.join("invoices", nom_fichier))]

lecteur = DatFileReader()

for invoice in list_invoices:
    # Lire le contenu du fichier (peut retourner une liste)
    contenu = lecteur.lire_texte_simple(f'invoices/{invoice}')
    
    # Vérifier si le contenu est une liste et la convertir en chaîne
    if isinstance(contenu, list):
        # Joindre les éléments de la liste avec des sauts de ligne
        contenu_str = '\n'.join(str(item) for item in contenu)
    else:
        # Sinon, convertir en chaîne
        contenu_str = str(contenu)
    
    # Remplacer les symboles dans le contenu
    nouveau_contenu = contenu_str.replace('%20', '_')
    
    # Construire le chemin du fichier de sortie
    nom_fichier = f"{invoice.replace('%20', '_')}"  # Remplacer aussi dans le nom de fichier
    chemin_complet = download_path / nom_fichier
    
    # Écrire le contenu modifié dans le nouveau fichier
    with open(chemin_complet, 'w', encoding='utf-8') as f_sortie:
        f_sortie.write(nouveau_contenu)
    
    print(f"Fichier traité : {invoice} -> {chemin_complet}")