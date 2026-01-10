import os
from pathlib import Path
from datfilereader import DatFileReader


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


list_invoices = [nom_fichier for nom_fichier in os.listdir("items") if
                     os.path.isfile(os.path.join("items", nom_fichier))]

# Create output directory if it doesn't exist
output_dir = Path("items_processed")
output_dir.mkdir(parents=True, exist_ok=True)

# Process each invoice
for invoice in list_invoices:
    input_path = f'items/{invoice}'
    output_path = output_dir / invoice
    
    # Replace %20 with _ in each file
    remplacer_symbole_direct(input_path, str(output_path), '%20', '_')