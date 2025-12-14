import os


fichiers = []
chemin = "invoice"
# Parcourt le dossier
for nom_fichier in os.listdir(chemin):
    # Vérifie si c'est un fichier (et non un sous-dossier)
    if os.path.isfile(os.path.join(chemin, nom_fichier)):
            fichiers.append(nom_fichier)

print(fichiers)