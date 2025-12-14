import pandas as pd
import struct
import numpy as np

class DatFileReader:
    def __init__(self):
        pass

    def lire_texte_simple(self, chemin_fichier):
        """
        Lecture simple d'un fichier .dat comme un fichier texte
        """
        try:
            with open(chemin_fichier, 'r', encoding='utf-8') as fichier:
                contenu = fichier.readlines()
                return [ligne.strip() for ligne in contenu]
        except UnicodeDecodeError:
            print("Le fichier n'est pas au format texte. Essayez la méthode binaire.")
            return None

    def lire_binaire(self, chemin_fichier):
        """
        Lecture d'un fichier .dat en mode binaire
        """
        try:
            with open(chemin_fichier, 'rb') as fichier:
                contenu = fichier.read()
                return contenu
        except Exception as e:
            print(f"Erreur lors de la lecture binaire: {str(e)}")
            return None

    def lire_structure(self, chemin_fichier, format_structure):
        """
        Lecture d'un fichier .dat avec une structure spécifique
        format_structure: format de struct (ex: 'iif' pour int,int,float)
        """
        try:
            with open(chemin_fichier, 'rb') as fichier:
                taille = struct.calcsize(format_structure)
                donnees = []
                while True:
                    buffer = fichier.read(taille)
                    if not buffer:
                        break
                    donnees.append(struct.unpack(format_structure, buffer))
                return donnees
        except Exception as e:
            print(f"Erreur lors de la lecture structurée: {str(e)}")
            return None

    def lire_csv_like(self, chemin_fichier, separateur=','):
        """
        Lecture d'un fichier .dat comme un CSV
        """
        try:
            df = pd.read_csv(chemin_fichier, sep=separateur)
            return df
        except Exception as e:
            print(f"Erreur lors de la lecture CSV: {str(e)}")
            return None

    def lire_fixed_width(self, chemin_fichier, largeurs_colonnes, noms_colonnes=None):
        """
        Lecture d'un fichier .dat avec des colonnes de largeur fixe
        largeurs_colonnes: liste des largeurs de chaque colonne
        noms_colonnes: liste optionnelle des noms de colonnes
        """
        try:
            df = pd.read_fwf(chemin_fichier, widths=largeurs_colonnes, names=noms_colonnes)
            return df
        except Exception as e:
            print(f"Erreur lors de la lecture fixed-width: {str(e)}")
            return None


lecteur = DatFileReader()
contenu = lecteur.lire_texte_simple('invoice/10000.dat')
print(contenu)