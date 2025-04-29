from excel_manager import ExcelManager

def main():
    # Remplacer par le chemin réel vers ton fichier tokens.xlsx
    filepath = "tokens.xlsx"
    
    # Initialiser
    manager = ExcelManager(filepath)
    
    # Charger
    print("Chargement du fichier (test excel)...")
    manager.load_excel()
    

    
    # Modifier un token (exemple : première ligne, colonne "QTT owned")
    print("Mise à jour du champ 'qtt_owned' à 5 pour la première ligne.")
    manager.update_token_field(0, "qtt_owned", 5)
    
    # Rafraîchir last_scraped sur les 3 premiers tokens
    print("Mise à jour du champ 'last_scraped' pour les 3 premiers tokens.")
    manager.update_last_scraped([0, 1, 2])
    
    # Vérifier état dirty
    if manager.is_dirty():
        print("Modifications détectées, sauvegarde en cours...")
        manager.save_excel()
    else:
        print("Aucune modification détectée.")
    
    print("✅ Test terminé.")

if __name__ == "__main__":
    main()
