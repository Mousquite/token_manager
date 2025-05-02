import sys
from PyQt5.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QPushButton,
    QLabel, QTableWidget, QTableWidgetItem, QMessageBox
)
from excel_manager import ExcelManager

class MainWindow(QMainWindow):
    def __init__(self, root, excel_manager):
        super().__init__()

        self.setWindowTitle("Token Manager")
        self.setGeometry(100, 100, 1200, 600)

        self.manager = ExcelManager("tokens.xlsx")

        # Interface
        self.central_widget = QWidget()
        self.setCentralWidget(self.central_widget)
        self.layout = QVBoxLayout()
        self.central_widget.setLayout(self.layout)




                # Création du bouton "Sauvegarder"
        self.save_button = QPushButton("Sauvegarder")
        self.save_button.clicked.connect(self.save_data)

        self.label = QLabel("Bienvenue dans le gestionnaire de tokens.")
        self.button = QPushButton("Charger les données")
        self.button.clicked.connect(self.load_table)

        self.table = QTableWidget()
                #gestion en-têtes
        self.table.setColumnCount(len(self.manager.headers))
        self.table.setHorizontalHeaderLabels([h.capitalize() for h in self.manager.headers])


        # Ajout des elements au layout
        self.layout.addWidget(self.label)
        self.layout.addWidget(self.button)
        self.layout.addWidget(self.table)
        self.layout.addWidget(self.save_button)


        

    def load_table(self):
        try:
            self.manager.load_excel()
            data = self.manager.get_all_data()
            if not data:
                QMessageBox.information(self, "Info", "Aucune donnée chargée.")
                return

            headers = list(data[0].keys())
            self.table.setRowCount(len(data))
            self.table.setColumnCount(len(headers))
            self.table.setHorizontalHeaderLabels(headers)

            for row_idx, row in enumerate(data):
                for col_idx, header in enumerate(headers):
                    value = str(row[header]) if row[header] is not None else ""
                    self.table.setItem(row_idx, col_idx, QTableWidgetItem(value))

            self.label.setText("Données chargées depuis tokens.xlsx")
        except Exception as e:
            QMessageBox.critical(self, "Erreur", str(e))


    def save_data(self):
        # Appel de la méthode save_excel() de excel_manager
        self.manager.update_from_table(self.table)
        self.manager.save_excel()
        self.label.setText("Données sauvegardées dans tokens.xlsx")


if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    sys.exit(app.exec_())
