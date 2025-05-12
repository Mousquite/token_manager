import sys
from PyQt5.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QPushButton, QShortcut,
    QLabel, QTableWidget, QTableWidgetItem, QMessageBox,
    QLineEdit, QMenu, QAction, QInputDialog, QAbstractItemView
)
from PyQt5.QtCore import Qt, QPoint, QObject, QEvent, QTimer
from excel_manager import ExcelManager
from PyQt5.QtGui import QKeySequence, QKeyEvent

class MainWindow(QMainWindow):
    def __init__(self, root, excel_manager):
        super().__init__()

        self.loading = False
        self.last_saved_state = None

        # Gestion Timer
        self.undo_timer = QTimer()
        self.undo_timer.setSingleShot(True)
        self.undo_timer.timeout.connect(self.save_state_for_undo)


        # Fenêtre principale
        self.setWindowTitle("Token Manager")
        self.setGeometry(100, 100, 1200, 600)

        # Gestionnaire Excel
        self.manager = ExcelManager("tokens.xlsx")

        # Interface centrale
        self.central_widget = QWidget()
        self.setCentralWidget(self.central_widget)
        self.layout = QVBoxLayout()
        self.central_widget.setLayout(self.layout)

        # Widgets de base
        self.label = QLabel("Bienvenue dans le gestionnaire de tokens.")
        self.button = QPushButton("Charger les données")
        self.button.clicked.connect(self.load_table)
        self.save_button = QPushButton("Sauvegarder")
        self.save_button.clicked.connect(self.save_data)
        self.undo_button = QPushButton("Annuler")
        self.undo_button.clicked.connect(self.undo_last_change)
        self.redo_button = QPushButton("Refaire")
        self.redo_button.clicked.connect(self.redo_last_change)
        #self.update_button = QPushButton("Mettre à jour la base de données")
        #self.update_button.clicked.connect(self.update_database)

        


        # Table principale
        self.table = TokenTableWidget()
        self.table.setFocusPolicy(Qt.StrongFocus)
        self.table.setFocus()
        self.table.setSortingEnabled(True)
        self.table.itemChanged.connect(self.save_state_for_undo)


        # En-têtes de colonnes
        self.table.setColumnCount(len(self.manager.headers))
        self.table.setHorizontalHeaderLabels([h.capitalize() for h in self.manager.headers])

        # Menu contextuel (table et en-têtes)
        self.table.setContextMenuPolicy(Qt.CustomContextMenu)
        self.table.customContextMenuRequested.connect(self.show_table_context_menu)
        self.table.horizontalHeader().setContextMenuPolicy(Qt.CustomContextMenu)
        self.table.horizontalHeader().customContextMenuRequested.connect(self.show_header_context_menu)

        # Glisser-déposer des colonnes
        self.table.setDragEnabled(False)
        self.table.setAcceptDrops(False)
        self.table.setDragDropMode(QAbstractItemView.NoDragDrop)

        #self.table.setDragDropOverwriteMode(False)
        #self.table.setDropIndicatorShown(True)
        
        header = self.table.horizontalHeader()
        header.setSectionsMovable(True)
        header.setDragEnabled(True)
        header.setDragDropMode(QAbstractItemView.InternalMove)


        # Raccourcis clavier
        delete_shortcut = QShortcut(QKeySequence(Qt.Key_Delete), self.table)
        delete_shortcut.activated.connect(self.clear_selected_cells)
        undo_shortcut = QShortcut(QKeySequence.Undo, self)  
        undo_shortcut.activated.connect(self.undo_last_change)
        copy_shortcut = QShortcut(QKeySequence.Copy, self.table)
        copy_shortcut.activated.connect(self.copy_cells)
        paste_shortcut = QShortcut(QKeySequence.Paste, self.table)
        paste_shortcut.activated.connect(self.paste_cells)
        cut_shortcut = QShortcut(QKeySequence.Cut, self.table)
        cut_shortcut.activated.connect(self.cut_cells)
        
        # Pile d'annulation
        self.undo_stack = []
        self.redo_stack = []
        self.loading = False
        self.table.itemChanged.connect(self.handle_item_changed)


        # Champ de recherche
        self.search_input = QLineEdit()
        self.search_input.setPlaceholderText("Rechercher...")
        self.search_input.textChanged.connect(self.filter_table)

        # Ajout des widgets au layout
        self.layout.addWidget(self.label)
        self.layout.addWidget(self.button)
        self.layout.addWidget(self.search_input)
        self.layout.addWidget(self.table)
        self.layout.addWidget(self.save_button)
        self.layout.addWidget(self.undo_button)
        self.layout.addWidget(self.redo_button)
        #self.layout.addWidget(self.update_button)
        


        

    def load_table(self):
        self.loading = True
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
        self.loading = False
        self.save_state_for_undo()


    def save_data(self):
        # Appel de la méthode save_excel() de excel_manager
        self.manager.update_from_table(self.table)
        self.manager.save_excel()
        self.label.setText("Données sauvegardées dans tokens.xlsx")


    def filter_table(self, text):
        self.loading = True
        text = text.strip().lower()
        for row in range(self.table.rowCount()):
            match = False
            for column in range(self.table.columnCount()):
                item = self.table.item(row, column)
                if item and text in item.text().lower():
                    match = True
                    break
            self.table.setRowHidden(row, not match)
        self.loading = False


    #Méthode pour le menu contextuel du tableau
    def show_table_context_menu(self, pos):
        self.loading = True
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

        # copy cut paste
        copy_action = QAction("Copier", self)
        copy_action.triggered.connect(self.copy_cells)
        menu.addAction(copy_action)

        cut_action = QAction("Couper", self)
        cut_action.triggered.connect(self.cut_cells)
        menu.addAction(cut_action)

        paste_action = QAction("Coller", self)
        paste_action.triggered.connect(self.paste_cells)
        menu.addAction(paste_action)


        # afficher les menus définis
        menu.exec_(self.table.viewport().mapToGlobal(pos))
        self.loading = False

    # Méthode pour le menu contextuel de l'en-tête
    def show_header_context_menu(self, pos):
        self.loading = True
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
        self.loading = False
        
    def rename_column(self, index):
        self.loading = True
        new_name, ok = QInputDialog.getText(self, "Renommer la colonne", "Nouveau nom :")
        if ok and new_name:
            self.manager.headers[index] = new_name
            self.table.setHorizontalHeaderLabels(self.manager.headers)
        self.loading = False
        self.save_state_for_undo()

    



    def show_all_columns(self):
        self.loading = True
        for col in range(self.table.columnCount()):
            self.table.setColumnHidden(col, False)
        self.loading = False
        self.save_state_for_undo()

    def add_row(self):
        row_position = self.table.rowCount()
        self.table.insertRow(row_position)
        for col in range(self.table.columnCount()):
            self.table.setItem(row_position, col, QTableWidgetItem(""))
        self.save_state_for_undo()

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
        self.save_state_for_undo()

    def delete_selected_row(self):
        selected_rows = sorted(set(index.row() for index in self.table.selectedIndexes()), reverse=True)
        for row in selected_rows:
            self.table.removeRow(row)
        self.save_state_for_undo()

    def delete_column(self, index):
        self.table.removeColumn(index)
        if index < len(self.manager.headers):
            del self.manager.headers[index]
        self.save_state_for_undo()


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

    def copy_cells(self):
        selected = self.table.selectedRanges()
        if selected:
            copied_text = ""
            for r in range(selected[0].topRow(), selected[0].bottomRow() + 1):
                row_data = []
                for c in range(selected[0].leftColumn(), selected[0].rightColumn() + 1):
                    item = self.table.item(r, c)
                    row_data.append(item.text() if item else "")
                copied_text += "\t".join(row_data) + "\n"
            QApplication.clipboard().setText(copied_text)
        

    def cut_cells(self):
        self.copy_cells()
        self.clear_selected_cells()
        

    def paste_cells(self):
        clipboard = QApplication.clipboard()
        text = clipboard.text()
        rows = text.splitlines()
        selected = self.table.selectedRanges()
        if not selected:
            return
        start_row = selected[0].topRow()
        start_col = selected[0].leftColumn()


        for r, row_text in enumerate(rows):
            columns = row_text.split("\t")
            for c, value in enumerate(columns):
                row_idx = start_row + r
                col_idx = start_col + c
                if row_idx < self.table.rowCount() and col_idx < self.table.columnCount():
                    self.table.setItem(row_idx, col_idx, QTableWidgetItem(value))



    ### LA PARTIE SUIVANTE EST A PEAUFINER
    ### probleme dans le stack, il faut regrouper les actions dans un seul stack
    ### pour undo en une action

                    
    # stacks pour undo
    def save_state_for_undo(self):


        #verif table
        rows = self.table.rowCount()
        cols = self.table.columnCount()
        if rows == 0 or cols == 0:
            print("Table vide ou non initialisée, état ignoré")
            return

        
        # Ne pas sauvegarder l'état si on est en train de charger/restaurer
        if getattr(self, "loading", False):
            print("État ignoré (chargement ou annulation en cours).")
            return

        # Capture l'état actuel
        state = []
        for row in range(rows):
            row_data = []
            for col in range(cols):
                item = self.table.item(row, col)
                row_data.append(item.text() if item else "")
            state.append(row_data)

        # Ne pas sauvegarder si l'état est identique au dernier
        if self.undo_stack and self.undo_stack[-1] == state:
            print("Aucun changement détecté, état non sauvegardé.")
            return

        # Sauvegarde l'état dans la pile undo
        self.undo_stack.append(state)

        # Vide la pile redo dès qu'un nouveau changement est fait
        self.redo_stack.clear()

        # Limite la taille de la pile undo à 50 entrées
        if len(self.undo_stack) > 50:
            self.undo_stack.pop(0)

        print(f"État sauvegardé pour annulation. Taille de la pile : {len(self.undo_stack)}")




    def handle_item_changed(self, item):
        if self.loading:
            return
        # Lancer le timer ou le redémarrer
        self.undo_timer.start(200)  # 200ms : regroupe les modifs faites rapidement

    def undo_last_change(self):
        if len(self.undo_stack) < 2:
            print("Aucun état précédent à restaurer.")
            return

        print(f"Undo demandé. Taille de la pile avant pop : {len(self.undo_stack)}")


        # Sauvegarder l'état courant dans redo_stack
        current_state = self.undo_stack[-1]
        if current_state:  # Assure que ce n’est pas None
            self.redo_stack.append(current_state.copy())
        else:
            print("État actuel invalide, non ajouté à redo_stack.")


            
        print("Annulation en cours")
        # Supprimer l'état courant
        self.undo_stack.pop()
        self.redo_stack.append(current_state)

        
        # Restaurer l'état précédent
        last_state = self.undo_stack[-1]

        
        self.loading = True  # désactiver temporairement save_state_for_undo pendant le remplissage
        self.table.blockSignals(True) #timer
        self.table.setRowCount(len(last_state))
        self.table.setColumnCount(len(last_state[0]))

        for row_idx, row_data in enumerate(last_state):
            for col_idx, value in enumerate(row_data):
                self.table.setItem(row_idx, col_idx, QTableWidgetItem(value))
        self.table.blockSignals(False) 
        # self.loading = False
        QTimer.singleShot(0, lambda: setattr(self, "loading", False))
        print("Annulation effectuée.")

    def redo_last_change(self):
        if not self.redo_stack:
            print("Aucun état à refaire.")
            return

        print(f"Redo demandé. Taille de la pile redo : {len(self.redo_stack)}")

        state_to_restore = self.redo_stack.pop()
        self.undo_stack.append(state_to_restore)

        self.restore_table_state(state_to_restore)
        print("Refaire effectué.")

    def restore_table_state(self, state):
        self.loading = True
        self.table.blockSignals(True)
        
        self.table.setRowCount(len(state))
        self.table.setColumnCount(len(state[0]))

        for row_idx, row_data in enumerate(state):
            for col_idx, value in enumerate(row_data):
                self.table.setItem(row_idx, col_idx, QTableWidgetItem(value))

        self.table.blockSignals(False)
        QTimer.singleShot(0, lambda: setattr(self, "loading", False))


### LA PARTIE PRECEDENTE EST A PEAUFINER
### probleme dans le stack, il faut regrouper les actions dans un seul stack
### pour undo en une action
    



class TokenTableWidget(QTableWidget):
    def __init__(self, parent=None):
        super().__init__(parent)

    def keyPressEvent(self, event):
        self.loading = True
        print("Key pressed:", event.key())  # <-- Debug

            
        if isinstance(event, QKeyEvent):
            # Détection de Meta+Z (Undo sur macOS)
            if event.key() == QKeySequence.Undo:
                print("Undo detected via Meta+Z")
                self.parent().undo_last_change()
                return

                # Supprimer contenu des cellules
            elif event.key() in (Qt.Key_Delete, Qt.Key_Backspace):
                print("Delete key pressed")
                for item in self.selectedItems():
                        if item is not None:
                            item.setText("")
                return

            else: super().keyPressEvent(event)

        self.loading = False







if __name__ == "__main__":
    app = QApplication(sys.argv)
    manager = ExcelManager("tokens.xlsx")
    window = MainWindow(None, manager)
    window.show()
    sys.exit(app.exec_())


