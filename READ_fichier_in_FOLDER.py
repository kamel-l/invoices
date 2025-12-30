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

        print(f"✔️ Remplacement effectué : {fichier_sortie}")

    except Exception as e:
        print(f"❌ Une erreur s'est produite : {str(e)}")


# --- Partie principale ---
lecteur = DatFileReader()
source = os.listdir("invoices25")

for invoice in source:
    # Chemin complet du fichier d'entrée
    fichier_entree = os.path.join("invoices25", invoice)

    # Chemin complet du fichier de sortie (نفس الاسم بعد التعديل)
    fichier_sortie = os.path.join("invoices25", invoice)

    # Exécuter le remplacement
    remplacer_symbole_direct(fichier_entree, fichier_sortie, " ", "_")
