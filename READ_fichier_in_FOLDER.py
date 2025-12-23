import os
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


list_invoices = [nom_fichier for nom_fichier in os.listdir("invoices") if
                     os.path.isfile(os.path.join("invoices", nom_fichier))]

print(list_invoices)
lecteur = DatFileReader()
for invoice in list_invoices:
      
      remplacer_symbole_direct(invoice, "100002.dat", '%20', '_')      