# src/excel_manager.py

import openpyxl
from datetime import datetime

class ExcelManager:
    def __init__(self, filepath="tokens.xlsx"):
        self.filepath = filepath
        self.workbook = None
        self.sheet = None
        self.headers = []
        self.dirty = False

    def load_excel(self):
        try:
            print("Chargement du fichier...")
            self.workbook = openpyxl.load_workbook(self.filepath)
            self.sheet = self.workbook.active
            self.headers = [cell.value for cell in self.sheet[1]]
            print(f"Chargement réussi ({self.sheet.max_row - 1} lignes).")
        except FileNotFoundError:
            raise FileNotFoundError(f"Fichier non trouvé : {self.filepath}")

    def get_all_tokens(self):
        tokens = []
        for row in self.sheet.iter_rows(min_row=2, values_only=True):
            token = {self.headers[i]: row[i] for i in range(len(self.headers))}
            tokens.append(token)
        return tokens

    def update_token_field(self, row_idx, field, value):
        if field not in self.headers:
            raise ValueError(f"Champ {field} introuvable.")
        col_idx = self.headers.index(field) + 1
        self.sheet.cell(row=row_idx + 2, column=col_idx, value=value)
        self.dirty = True

    def update_last_scraped(self, row_idx):
        if "last_scraped" not in self.headers:
            raise ValueError("Champ 'last_scraped' non trouvé.")
        col_idx = self.headers.index("last_scraped") + 1
        today = datetime.now().strftime("%Y-%m-%d")
        self.sheet.cell(row=row_idx + 2, column=col_idx, value=today)
        self.dirty = True

    def save(self):
        if not self.dirty:
            print("Aucune modification à sauvegarder.")
            return
        self.workbook.save(self.filepath)
        print(f"Modifications sauvegardées dans {self.filepath}.")
        self.dirty = False

    def is_dirty(self):
        return self.dirty
