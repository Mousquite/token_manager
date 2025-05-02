# src/main.py

import sys
from PyQt5.QtWidgets import QApplication
from excel_manager import ExcelManager
from main_window import MainWindow

def main():
    # Initialisation de l'application Qt
    app = QApplication(sys.argv)

    # Chargement des données Excel
    excel_manager = ExcelManager()
    excel_manager.load_excel()

    # Création et affichage de la fenêtre principale
    window = MainWindow(None, excel_manager)
    window.show()

    # Boucle principale Qt
    sys.exit(app.exec_())

if __name__ == "__main__":
    main()
