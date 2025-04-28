# src/main.py

from excel_manager import ExcelManager
from gui_manager import GuiManager

def main():
    # Initialiser Excel Manager
    excel_manager = ExcelManager()

    # Initialiser GUI Manager
    gui = GuiManager(excel_manager)
    
    # Lancer l'application
    gui.run()

if __name__ == "__main__":
    main()
