import sys
from PyQt5.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QPushButton, QShortcut,
    QLabel, QTableWidget, QTableWidgetItem, QMessageBox,
    QLineEdit, QMenu, QAction, QInputDialog, QAbstractItemView
)
from PyQt5.QtCore import Qt, QPoint, QObject, QEvent, QTimer
from excel_manager import ExcelManager
from PyQt5.QtGui import QKeySequence, QKeyEvent, QFont, QColor

import pandas as pd
import json
import os
import hashlib

def hash_df(df):
        return hashlib.md5(pd.util.hash_pandas_object(df, index=True).values).hexdigest()

def compare_dfs(df1, df2):

    diffs = []
    rows = max(len(df1), len(df2))
    cols = set(df1.columns).union(set(df2.columns))

    for i in range(rows):
        for col in cols:
            val1 = df1[col][i] if i < len(df1) and col in df1.columns else "<missing>"
            val2 = df2[col][i] if i < len(df2) and col in df2.columns else "<missing>"
            if str(val1) != str(val2):
                diffs.append(f"🟥 Diff ligne {i}, colonne '{col}': '{val1}' -> '{val2}'")
    return diffs

def log_df_differences(df_before: pd.DataFrame, df_after: pd.DataFrame, locked_cells: set):
    for row in range(min(len(df_before), len(df_after))):
        for col in range(min(len(df_before.columns), len(df_after.columns))):
            val_before = df_before.iat[row, col]
            val_after = df_after.iat[row, col]
            if pd.isna(val_before): val_before = ""
            if pd.isna(val_after): val_after = ""
            if str(val_before) != str(val_after):
                col_name = df_before.columns[col]
                is_locked = (row, col) in locked_cells
                status = "🔒" if is_locked else "⚠️"
                print(f"{status} Diff (row={row}, col={col}, '{col_name}') : '{val_before}' → '{val_after}'")


class MainWindow(QMainWindow):
    def __init__(self, root, excel_manager):
        super().__init__()

        self.loading = False
        self.last_saved_state = None
        self.locked_cells = set()


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
        self.button.clicked.connect(lambda: self.load_table(from_file=True))
        self.save_button = QPushButton("Sauvegarder")
        self.save_button.clicked.connect(self.save_data)
        self.undo_button = QPushButton("Annuler")
        self.undo_button.clicked.connect(self.undo_last_change)
        self.redo_button = QPushButton("Refaire")
        self.redo_button.clicked.connect(self.redo_last_change)
        self.import_button = QPushButton("Importer")
        self.import_button.clicked.connect(self.import_new_tokens)

        # Table principale
        self.table = TokenTableWidget()
        self.table.setFocusPolicy(Qt.StrongFocus)
        self.table.setFocus()
        self.table.setSortingEnabled(True)
        self.table.itemChanged.connect(self.save_state_for_undo)
        self.table.itemChanged.connect(self.handle_cell_change)

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
        self.layout.addWidget(self.import_button)

        self.load_table_settings()

    def load_table(self, from_file: bool = True):
        self.loading = True
        try:
            if from_file:
                print(">>> Chargement des données depuis tokens.xlsx via load_excel()")
                self.manager.load_excel()
                print("🟢 Données chargées depuis le fichier :")
                print(self.manager.df.head(10).to_string()) 
            else:
                print(">>> Chargement des données depuis la mémoire (self.manager.df)")

            data = self.manager.get_all_data()
            if not data:
                QMessageBox.information(self, "Info", "Aucune donnée chargée.")
                return

            df = self.manager.df

            # Supprime toute colonne 'checked' déjà présente (on la gère via la table uniquement)
            if "checked" in df.columns:
                df.drop(columns=["checked"], inplace=True)

            # Chargement ou création de la colonne temporaire "checked" pour l'affichage
            checked_values = [False] * len(df)
            if hasattr(self, 'locked_cells'):  # pour s'assurer que locked_cells existe
                self.locked_cells = set()

            if "checked" not in df.columns:
                df["checked"] = False

            headers = ["✔"] + [col for col in df.columns if col != "checked"]

            # Initialisation de la table
            self.table.clear()
            self.table.setRowCount(0)
            self.table.setColumnCount(len(headers))
            self.table.setHorizontalHeaderLabels(headers)
            self.table.setRowCount(len(df))

            for row_idx, (_, row) in enumerate(df.iterrows()):
                # Colonne 0 : checkbox
                checkbox_item = QTableWidgetItem()
                checkbox_item.setFlags(Qt.ItemIsUserCheckable | Qt.ItemIsEnabled)
                checked = row.get("checked", False) if "checked" in row else False
                checkbox_item.setCheckState(Qt.Checked if checked else Qt.Unchecked)
                self.table.setItem(row_idx, 0, checkbox_item)

                # Autres colonnes
                for col_idx, (col_name, value) in enumerate(row.items()):
                    if col_name == "checked":
                        continue  # déjà traité

                    item = QTableWidgetItem(str(value) if pd.notna(value) else "")
                    col_pos = col_idx + 1  # +1 car colonne 0 = checkbox

                    if (row_idx, col_pos) in self.locked_cells:
                        if not value or str(value).strip() == "":
                            self.locked_cells.discard((row_idx, col_pos))
                        else:
                            font = item.font()
                            font.setBold(True)
                            item.setFont(font)

                    self.table.setItem(row_idx, col_pos, item)

            # Chargement des cellules verrouillées
            locked_path = os.path.join(os.path.dirname(self.manager.filepath), "locked_cells.json")
            self.locked_cells = set()
            if os.path.exists(locked_path):
                with open(locked_path, "r") as f:
                    loaded = json.load(f)
                    self.locked_cells = set(tuple(cell) for cell in loaded)

            self.label.setText("Données chargées depuis tokens.xlsx")

        except Exception as e:
            QMessageBox.critical(self, "Erreur", str(e))

        self.load_table_settings()
        self.load_locked_cells()
        self.apply_checked_column()
        self.loading = False
        self.save_state_for_undo()

        # Appliquer les styles aux cellules verrouillées
        for (row, col) in self.locked_cells:
            item = self.table.item(row, col)
            if item:
                font = item.font()
                font.setBold(True)
                item.setFont(font)
                item.setBackground(QColor(80, 80, 80))

        print("🟡 Données extraites de la table vers df (sans les cases cochées) :")
        print(self.manager.df.head(10).to_string())

    def save_data(self):
        print("🔽 Sauvegarde en cours...")

        # Capture avant
        df_before = self.manager.df.copy(deep=True)

        # Sauvegarde JSON verrou
        locked_path = os.path.join(os.path.dirname(self.manager.filepath), "locked_cells.json")
        with open(locked_path, "w") as f:
            json.dump(list(self.locked_cells), f)

        # Mise à jour du DataFrame depuis la table (hors colonne des cases)
        self.update_df_from_table(skip_columns=[0])
        self.sync_checked_column()

        # Sauvegarde Excel
        self.manager.save_excel()

        # Rechargement du fichier
        self.manager.load_excel()
        df_after = self.manager.df.copy(deep=True)

        # Rechargement dans la table
        self.load_table()

        # Logs de différence
        hash_before = hash_df(df_before)
        hash_after = hash_df(df_after)
        print(f"✅ Hash DF avant save: {hash_before}")
        print(f"✅ Hash DF après save: {hash_after}")
        if hash_before != hash_after:
            print("❌ Le DataFrame a changé pendant la sauvegarde")
            log_df_differences(df_before, df_after, self.locked_cells)
        else:
            print("✅ Aucune différence détectée dans le DataFrame.")

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

          # --- Nouvelles actions : verrouiller / déverrouiller ---
        lock_action = QAction("Verrouiller la sélection", self)
        lock_action.triggered.connect(self.lock_selected_cells)

        unlock_action = QAction("Déverrouiller la sélection", self)
        unlock_action.triggered.connect(self.unlock_selected_cells)
        menu.addAction(lock_action)
        menu.addAction(unlock_action)

        # afficher les menus définis
        menu.exec_(self.table.viewport().mapToGlobal(pos))
        self.loading = False

    # Méthode pour le menu contextuel de l'en-tête
    def show_header_context_menu(self, pos):
        index = self.table.horizontalHeader().logicalIndexAt(pos)
        if index < 0 or index >= self.table.columnCount():
            return  # clic hors des colonnes
        self.loading = True
        menu = QMenu(self)
        index = self.table.horizontalHeader().logicalIndexAt(pos)


        # masquer la colonne
        header_item = self.table.horizontalHeaderItem(index)
        header_label = header_item.text() if header_item else f"Colonne {index}"
        hide_column_action = QAction(f"Masquer la colonne '{header_label}'", self)
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

    def sync_checked_column(self):
        checked_values = []
        for row in range(self.table.rowCount()):
            item = self.table.item(row, 0)
            checked = item.checkState() == Qt.Checked if item else False
            checked_values.append(checked)

        if "checked" in self.manager.df.columns:
            self.manager.df["checked"] = checked_values
        else:
            self.manager.df.insert(0, "checked", checked_values)

    def apply_checked_column(self):
        if "checked" in self.manager.df.columns:
            for row in range(self.table.rowCount()):
                item = self.table.item(row, 0)
                if item:
                    state = Qt.Checked if self.manager.df.at[row, "checked"] else Qt.Unchecked
                    item.setCheckState(state)

    def add_row(self):
        row_position = self.table.rowCount()
        self.table.insertRow(row_position)
        for col in range(self.table.columnCount()):
            self.table.setItem(row_position, col, QTableWidgetItem(""))
        self.update_df_from_table(skip_columns=[0])
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

    def save_table_settings(self, path="table_settings.json"):
        header = self.table.horizontalHeader()
        settings = {
            "column_order": [header.visualIndex(i) for i in range(header.count())],
            "hidden_columns": [i for i in range(self.table.columnCount()) if self.table.isColumnHidden(i)],
            "column_widths": {str(i): self.table.columnWidth(i) for i in range(self.table.columnCount())}
        }
        with open(path, "w") as f:
            json.dump(settings, f)    
    
    def load_table_settings(self, path="table_settings.json"):
        try:
            with open(path, "r") as f:
                settings = json.load(f)

            header = self.table.horizontalHeader()

            # Ordre des colonnes
            if "column_order" in settings:
                for logical, visual in enumerate(settings["column_order"]):
                    header.moveSection(header.visualIndex(logical), visual)

            # Colonnes masquées
            if "hidden_columns" in settings:
                for i in range(self.table.columnCount()):
                    self.table.setColumnHidden(i, i in settings["hidden_columns"])

            # Largeurs de colonnes
            if "column_widths" in settings:
                for i_str, width in settings["column_widths"].items():
                    i = int(i_str)
                    self.table.setColumnWidth(i, width)

        except Exception as e:
            print(f"Erreur lors du chargement des préférences d'affichage : {e}")

    def handle_cell_change(self, item: QTableWidgetItem):
        if self.loading or not item:
            return

        row = item.row()
        col = item.column()

        if (row, col) in self.locked_cells:
            print(f"[🔒 VERROUILLÉ] Cellule ({row}, {col}) → modification annulée.")
            old_value = self.manager.df.iat[row, col]
            self.table.blockSignals(True)
            item.setText(str(old_value) if pd.notna(old_value) else "")
            self.table.blockSignals(False)
            return

        new_value = item.text()
        old_value = self.manager.df.iat[row, col]

        if pd.isna(old_value):
            old_value = ""

        if new_value != str(old_value):
            self.manager.df.iat[row, col] = new_value
            print(f"📝 Cellule modifiée : ({row}, {col}) « {old_value} » → « {new_value} »")
            self.save_state_for_undo()

    
    def clear_selected_cells(self):
         for item in self.table.selectedItems():
            if item is None:
                continue
            row = item.row()
            col = item.column()
            if (row, col) in self.table.locked_cells:
                continue 
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
        if not text:
            return
    
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


                if row_idx >= self.table.rowCount() and col_idx < self.table.columnCount():
                    continue

                if (row_idx, col_idx) in self.table.locked_cells:
                        continue 

                item = QTableWidgetItem(value)
                self.table.setItem(row_idx, col_idx, QTableWidgetItem(value))
                #self.mark_cell_modified(row_idx, col_idx) 

    def import_new_tokens(self):
        try:
            
            self.update_df_from_table()
            tnew = pd.read_excel("newtokens.xlsx")

            # On nettoie préventivement "checked" si elle est là
            if "checked" in tnew.columns:
                tnew = tnew.drop(columns=["checked"])

            self.manager.import_table(tnew)
            self.load_table(from_file=False)
            QMessageBox.information(self, "Import réussi", "Les données ont été importées et fusionnées avec succès.")
        except Exception as e:
            QMessageBox.critical(self, "Erreur d'import", str(e))

    def lock_selected_cells(self):
        for item in self.table.selectedIndexes():
            row, col = item.row(), item.column()
            self.locked_cells.add((row, col))  # corrigé ici
            item_widget = self.table.item(row, col)
            if item_widget:
                font = item_widget.font()
                font.setBold(True)
                item_widget.setFont(font)
                item_widget.setBackground(QColor(80, 80, 80))  

    def unlock_selected_cells(self):
        for item in self.table.selectedIndexes():
            row, col = item.row(), item.column()
            self.locked_cells.discard((row, col))  # corrigé ici
            item_widget = self.table.item(row, col)
            if item_widget:
                font = item_widget.font()
                font.setBold(False)
                item_widget.setFont(font)
                item_widget.setBackground(QColor(0, 0, 0))  

    def load_locked_cells(self):
        locked_path = os.path.join(os.path.dirname(self.manager.filepath), "locked_cells.json")
        if os.path.exists(locked_path):
            with open(locked_path, "r") as f:
                loaded = json.load(f)
                self.locked_cells = set()
                for row, col in loaded:
                    if row < self.table.rowCount() and col < self.table.columnCount():
                        self.locked_cells.add((row, col))
            print(f"🔐 {len(self.locked_cells)} cellules verrouillées chargées.")
        else:
            self.locked_cells = set()

    def toggle_check_selection(self, check=True):
        state = Qt.Checked if check else Qt.Unchecked
        selected = self.table.selectedRanges()
        if not selected:
            return

        for rng in selected:
            for row in range(rng.topRow(), rng.bottomRow() + 1):
                item = self.table.item(row, 0)
                if item is not None:
                    item.setCheckState(state)
        
   
    """def mark_cell_modified(self, row, col):
        item = self.table.item(row, col)
        if item:
            item.setBackground(QColor(255, 255, 200))"""
    ### LA PARTIE SUIVANTE EST A PEAUFINER
    ### probleme dans le stack, il faut regrouper les actions dans un seul stack
    ### pour undo en une action
    ### il y a aussi des stacks vide wen undo puis redo
                    
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

    def update_df_from_table(self, skip_columns=None):
        if self.df is None:
            return

        if skip_columns is None:
            skip_columns = []

        for row in range(self.rowCount()):
            for col in range(1, self.columnCount()):  # col=0 = checkbox
                model_col = col - 1  # Décalage : DataFrame n’a pas la checkbox

                if col in skip_columns:
                    continue
                if (row, model_col) in self.locked_cells:
                    print(f"🔒 [SKIP] Cellule verrouillée ignorée ({row}, {model_col})")
                    continue

                item = self.item(row, col)
                value = item.text() if item else None
                value = value if value != "" else None
                self.df.iat[row, model_col] = value

        print("🟡 Données extraites de la table vers df (sans les cases cochées) :")
        print(self.df.head(10).to_string())







class TokenTableWidget(QTableWidget):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.locked_cells = set()  # (row, column) tuples

    def keyPressEvent(self, event):
        self.loading = True
            
        if isinstance(event, QKeyEvent):
            # Détection de Meta+Z (Undo sur macOS)
            if event.key() == QKeySequence.Undo:
                self.parent().undo_last_change()
                return

                # Supprimer contenu des cellules
            elif event.key() in (Qt.Key_Delete, Qt.Key_Backspace):
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


