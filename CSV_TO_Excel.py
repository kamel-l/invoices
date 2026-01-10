import pandas as pd
from pathlib import Path
import os

def convertir_csv_vers_excel(dossier_source, dossier_destination=None):
    """
    Convertit tous les fichiers CSV d'un dossier en fichiers Excel (.xlsx)
    
    Args:
        dossier_source (str): Chemin du dossier contenant les fichiers CSV
        dossier_destination (str, optional): Chemin du dossier de sortie. 
                                            Si None, utilise le dossier source
    """
    # Créer les objets Path
    chemin_source = Path(dossier_source)
    
    # Définir le dossier de destination
    if dossier_destination is None:
        chemin_destination = chemin_source
    else:
        chemin_destination = Path(dossier_destination)
        # Créer le dossier de destination s'il n'existe pas
        chemin_destination.mkdir(parents=True, exist_ok=True)
    
    # Vérifier si le dossier source existe
    if not chemin_source.exists():
        print(f"❌ Erreur : Le dossier '{dossier_source}' n'existe pas.")
        return
    
    # Trouver tous les fichiers CSV
    fichiers_csv = list(chemin_source.glob('*.csv'))
    
    if not fichiers_csv:
        print(f"⚠️  Aucun fichier CSV trouvé dans '{dossier_source}'")
        return
    
    print(f"📁 {len(fichiers_csv)} fichier(s) CSV trouvé(s)\n")
    
    # Convertir chaque fichier
    fichiers_convertis = 0
    fichiers_erreur = 0
    
    for fichier_csv in fichiers_csv:
        try:
            # Lire le fichier CSV
            df = pd.read_csv(fichier_csv, encoding='utf-8')
            
            # Créer le nom du fichier Excel
            sheet_name = fichier_csv.stem
            nom_excel = fichier_csv.stem + '.xlsx'
            chemin_excel = chemin_destination / nom_excel
            
            # Écrire dans Excel
            with pd.ExcelWriter(chemin_excel, engine='xlsxwriter') as writer:
                df.to_excel(writer, sheet_name=sheet_name, index=False)
                
                # Optionnel : Ajuster la largeur des colonnes
                worksheet = writer.sheets[sheet_name]
                for idx, col in enumerate(df.columns):
                    max_length = max(
                        df[col].astype(str).apply(len).max(),
                        len(str(col))
                    )
                    worksheet.set_column(idx, idx, min(max_length + 2, 50))
            
            print(f"✅ {fichier_csv.name} → {nom_excel}")
            fichiers_convertis += 1
            
        except Exception as e:
            print(f"❌ Erreur lors de la conversion de {fichier_csv.name}: {str(e)}")
            fichiers_erreur += 1
    
    # Résumé
    print(f"\n{'='*50}")
    print(f"✅ Fichiers convertis avec succès : {fichiers_convertis}")
    if fichiers_erreur > 0:
        print(f"❌ Fichiers en erreur : {fichiers_erreur}")
    print(f"{'='*50}")


# ==================== UTILISATION ====================

if __name__ == "__main__":
    # Option 1 : Convertir les CSV du même dossier
    #convertir_csv_vers_excel('analyse 2025')
    
    # Option 2 : Convertir vers un autre dossier
    convertir_csv_vers_excel('analyse 2025', 'resultats_excel')
    
    # Option 3 : Convertir depuis le dossier courant
    # convertir_csv_vers_excel('.')