import os
from pathlib import Path


def renommer_fichiers(dossier, ancien_pattern, nouveau_pattern):
    """
    Rename files in a folder by replacing a pattern in the filename
    
    Args:
        dossier: Path to the folder containing the files
        ancien_pattern: Pattern to replace (e.g., '_20')
        nouveau_pattern: New pattern (e.g., '_')
    """
    dossier_path = Path(dossier)
    
    if not dossier_path.exists():
        print(f"Error: Folder '{dossier}' does not exist")
        return
    
    fichiers_renommes = 0
    
    # Iterate through all files in the folder
    for fichier in dossier_path.iterdir():
        if fichier.is_file() and ancien_pattern in fichier.name:
            # Create new filename
            nouveau_nom = fichier.name.replace(ancien_pattern, nouveau_pattern)
            nouveau_chemin = fichier.parent / nouveau_nom
            
            # Rename the file
            try:
                fichier.rename(nouveau_chemin)
                print(f"Renamed: {fichier.name} -> {nouveau_nom}")
                fichiers_renommes += 1
            except Exception as e:
                print(f"Error renaming {fichier.name}: {str(e)}")
    
    print(f"\nTotal files renamed: {fichiers_renommes}")


# Example usage:
# Replace 'invoice' with your actual folder name
renommer_fichiers('items', '_2F4_2F4', '')