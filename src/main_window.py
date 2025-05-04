import sys
from PyQt5.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QPushButton, QShortcut,
    QLabel, QTableWidget, QTableWidgetItem, QMessageBox,
    QLineEdit, QMenu, QAction, QInputDialog, QAbstractItemView
)
from PyQt5.QtCore import Qt, QPoint, QObject, QEvent
from excel_manager import ExcelManager
from PyQt5.QtGui import QKeySequence

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

        self.table = TokenTableWidget()

        self.table.setFocusPolicy(Qt.StrongFocus)
        self.table.setFocus()

        delete_shortcut = QShortcut(QKeySequence(Qt.Key_Delete), self.table)
        delete_shortcut.activated.connect(self.clear_selected_cells)


        self.table.setSortingEnabled(True) #activer le tri

        # ajout le glisser-deposer des colonnes
        self.table.setDragEnabled(True)
        self.table.setAcceptDrops(True)
        self.table.setDragDropOverwriteMode(False)
        self.table.setDropIndicatorShown(True)
        self.table.setDragDropMode(QAbstractItemView.InternalMove)
        self.table.horizontalHeader().setSectionsMovable(True)
        self.table.horizontalHeader().setDragEnabled(True)
        self.table.horizontalHeader().setDragDropMode(QAbstractItemView.InternalMove)

                #gestion en-têtes
        self.table.setColumnCount(len(self.manager.headers))
        self.table.setHorizontalHeaderLabels([h.capitalize() for h in self.manager.headers])

            # ajout des menus contextuels
        self.table.setContextMenuPolicy(Qt.CustomContextMenu)
        self.table.customContextMenuRequested.connect(self.show_table_context_menu)

        self.table.horizontalHeader().setContextMenuPolicy(Qt.CustomContextMenu)
        self.table.horizontalHeader().customContextMenuRequested.connect(self.show_header_context_menu)



        # bouton recherche
        self.search_input = QLineEdit()
        self.search_input.setPlaceholderText("Rechercher...")
        self.search_input.textChanged.connect(self.filter_table)

        # Ajout des elements au layout
        self.layout.addWidget(self.label)
        self.layout.addWidget(self.button)
        self.layout.addWidget(self.table)
        self.layout.addWidget(self.save_button)
        self.layout.addWidget(self.search_input)

        

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


    def filter_table(self, text):
        text = text.strip().lower()
        for row in range(self.table.rowCount()):
            match = False
            for column in range(self.table.columnCount()):
                item = self.table.item(row, column)
                if item and text in item.text().lower():
                    match = True
                    break
            self.table.setRowHidden(row, not match)


    #Méthode pour le menu contextuel du tableau
    def show_table_context_menu(self, pos):
        menu = QMenu(self)
        
        add_row_action = QAction("Ajouter une ligne", self)
        add_row_action.triggered.connect(self.add_row)
        menu.addAction(add_row_action)

        delete_row_action = QAction("Supprimer la ligne", self)
        delete_row_action.triggered.connect(self.delete_selected_row)
        menu.addAction(delete_row_action)

        duplicate_row_action = QAction("Dupliquer la ligne", self)
        duplicate_row_action.triggered.connect(self.duplicate_selected_row)
        menu.addAction(duplicate_row_action)

        # supprimer cell par clic droit
        clear_cells_action = QAction("Effacer les cellules sélectionnées", self)
        clear_cells_action.triggered.connect(self.clear_selected_cells)
        menu.addAction(clear_cells_action)


        # afficher les menus définis
        menu.exec_(self.table.viewport().mapToGlobal(pos))
        

    # Méthode pour le menu contextuel de l'en-tête
    def show_header_context_menu(self, pos):
        menu = QMenu(self)
        index = self.table.horizontalHeader().logicalIndexAt(pos)


        # masquer la colonne
        hide_column_action = QAction(f"Masquer la colonne '{self.manager.headers[index]}'", self)
        hide_column_action.triggered.connect(lambda: self.table.setColumnHidden(index, True))
        menu.addAction(hide_column_action)

        # afficher toutes les colonnes
        show_all_columns_action = QAction("Afficher toutes les colonnes", self)
        show_all_columns_action.triggered.connect(self.show_all_columns)
        menu.addAction(show_all_columns_action)

        # ajouter colonne
        add_column_action = QAction("Ajouter une colonne", self)
        add_column_action.triggered.connect(self.add_column)
        menu.addAction(add_column_action)

        
        # renommer colonne
        rename_column_action = QAction("Renommer la colonne", self)
        rename_column_action.triggered.connect(lambda: self.rename_column(index))
        menu.addAction(rename_column_action)

        # delete colonne
        delete_column_action = QAction("Supprimer la colonne", self)
        delete_column_action.triggered.connect(lambda: self.delete_column(index))
        menu.addAction(delete_column_action)


        # afficher les menu definis
        menu.exec_(self.table.horizontalHeader().viewport().mapToGlobal(pos))

        
    def rename_column(self, index):
        new_name, ok = QInputDialog.getText(self, "Renommer la colonne", "Nouveau nom :")
        if ok and new_name:
            self.manager.headers[index] = new_name
            self.table.setHorizontalHeaderLabels(self.manager.headers)

    



    def show_all_columns(self):
        for col in range(self.table.columnCount()):
            self.table.setColumnHidden(col, False)

    def add_row(self):
        row_position = self.table.rowCount()
        self.table.insertRow(row_position)
        for col in range(self.table.columnCount()):
            self.table.setItem(row_position, col, QTableWidgetItem(""))

    def add_column(self):
        column_name, ok = QInputDialog.getText(self, "Ajouter une colonne", "Nom de la nouvelle colonne :")
        if ok and column_name:
            self.manager.headers.append(column_name)
            current_column_count = self.table.columnCount()
            self.table.setColumnCount(current_column_count + 1)
            self.table.setHorizontalHeaderLabels(self.manager.headers)

            # Ajoute des cellules vides pour chaque ligne existante
            for row in range(self.table.rowCount()):

                  self.table.setItem(row, current_column_count, QTableWidgetItem(""))

    def delete_selected_row(self):
        selected_rows = sorted(set(index.row() for index in self.table.selectedIndexes()), reverse=True)
        for row in selected_rows:
            self.table.removeRow(row)

    def delete_column(self, index):
        self.table.removeColumn(index)
        if index < len(self.manager.headers):
            del self.manager.headers[index]


    def clear_selected_cells(self):
        for item in self.table.selectedItems():
            if item is not None:
                item.setText("")

    def duplicate_selected_row(self):
        selected_rows = list(set(index.row() for index in self.table.selectedIndexes()))
        if not selected_rows:
            return

        times, ok = QInputDialog.getInt(self, "Dupliquer la ligne", "Combien de fois ?", 1, 1)
        if ok:
            for row in selected_rows:
                row_data = [
                    self.table.item(row, col).text() if self.table.item(row, col) else ""
                    for col in range(self.table.columnCount())
                ]
                for _ in range(times):
                    row_position = self.table.rowCount()
                    self.table.insertRow(row_position)
                    for col_idx, value in enumerate(row_data):
                        self.table.setItem(row_position, col_idx, QTableWidgetItem(value))




class TokenTableWidget(QTableWidget):
    def __init__(self, parent=None):
        super().__init__(parent)

    def keyPressEvent(self, event):
        if event.key() == Qt.Key_Delete:
            for item in self.selectedItems():
                if item:
                    item.setText("")
        else:
            super().keyPressEvent(event)


if __name__ == "__main__":
    app = QApplication(sys.argv)
    manager = ExcelManager("tokens.xlsx")
    window = MainWindow(None, manager)
    window.show()
    sys.exit(app.exec_())
