# src/config.py

# Chemins importants
EXCEL_FILE_PATH = "data/tokens.xlsx"  # Chemin vers ton fichier Excel principal
LOG_FOLDER_PATH = "logs/"             # Dossier où enregistrer les logs

# Paramètres de l'application
DIRTY_STATE_CHECK_INTERVAL = 60  # en secondes (vérification toutes les 60s si besoin)
DEFAULT_LATENCY_AFTER_ACTION = 5  # secondes d'attente après envoi de transactions

# Comportement
SAVE_WARNING_ON_EXIT = True  # Avertir si fichier non sauvegardé
SELECTION_PERSISTENCE = True # Ne pas effacer la sélection après action

# Options d'affichage
MODIFIED_ROW_COLOR = "#FFF0F0" # Couleur légère pour lignes modifiées
INACTIVE_ROW_COLOR = "#F0F0F0" # Couleur grisée pour actif=0

# Options de log
ENABLE_LOCAL_LOG = True
LOCAL_LOG_FILENAME = "logs/project_log.txt"
