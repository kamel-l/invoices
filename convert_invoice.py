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

remplacer_symbole_direct("invoices2024.csv", 'invoices10.csv', '%20', ' ')