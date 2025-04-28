import pandas as pd
import os
from datetime import datetime

class ExcelManager:
    def __init__(self, filepath):
        self.filepath = filepath
        self.df = None
        self.dirty = False  # Pour savoir si des modifications non sauvegardées existent

    def load_excel(self):
        """Charge le fichier Excel dans un DataFrame."""
        if not os.path.exists(self.filepath):
            raise FileNotFoundError(f"Fichier non trouvé : {self.filepath}")
        self.df = pd.read_excel(self.filepath, engine="openpyxl")
        self.check_expired_listings()
        print(f"Chargement réussi ({len(self.df)} lignes).")

    def save_excel(self):
        """Sauvegarde le DataFrame actuel dans le fichier Excel."""
        if self.df is None:
            raise ValueError("Aucun fichier chargé.")
        self.df.to_excel(self.filepath, index=False, engine="openpyxl")
        self.dirty = False
        print("Sauvegarde réussie.")

    def mark_dirty(self):
        """Indique qu'il y a eu des modifications non sauvegardées."""
        self.dirty = True

    def update_token_field(self, index, field, value):
        """Met à jour un champ pour une ligne précise."""
        if self.df is None:
            raise ValueError("Aucun fichier chargé.")
        if field not in self.df.columns:
            raise ValueError(f"Champ {field} introuvable.")
        self.df.at[index, field] = value
        self.mark_dirty()

    def batch_update_tokens(self, indices, field, value):
        """Met à jour un champ pour plusieurs lignes."""
        if self.df is None:
            raise ValueError("Aucun fichier chargé.")
        if field not in self.df.columns:
            raise ValueError(f"Champ {field} introuvable.")
        for idx in indices:
            self.df.at[idx, field] = value
        self.mark_dirty()

    def refresh_last_scraped(self, indices):
        """Met à jour la date last_scraped des tokens mis à jour."""
        now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        for idx in indices:
            self.df.at[idx, "last_scraped"] = now
        self.mark_dirty()

    def update_last_listing(self, index):
        """Met à jour last_listing et last_duration lors d'un listing."""
        now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        self.df.at[index, "last_listing"] = now
        duration = self.df.at[index, "duration"]
        self.df.at[index, "last_duration"] = duration
        self.mark_dirty()

    def check_expired_listings(self):
        """Réinitialise last_listing si le listing est expiré."""
        now = datetime.now()
        for idx, row in self.df.iterrows():
            last_listing = row.get("last_listing")
            last_duration = row.get("last_duration")
            if pd.notna(last_listing) and pd.notna(last_duration):
                try:
                    listing_date = pd.to_datetime(last_listing)
                    days_passed = (now - listing_date).days
                    if days_passed > int(last_duration):
                        # Expiré
                        self.df.at[idx, "last_listing"] = ""
                        self.df.at[idx, "last_duration"] = ""
                        print(f"Listing expiré pour index {idx} (réinitialisé).")
                except Exception as e:
                    print(f"Erreur parsing date index {idx}: {e}")
        self.mark_dirty()

    def get_filtered_tokens(self, filters=None):
        """Retourne un sous-ensemble du tableau selon des filtres."""
        if self.df is None:
            raise ValueError("Aucun fichier chargé.")
        if not filters:
            return self.df
        df_filtered = self.df.copy()
        for field, condition in filters.items():
            if field not in self.df.columns:
                continue
            df_filtered = df_filtered.query(condition)
        return df_filtered

    def is_dirty(self):
        """Retourne l'état de modification."""
        return self.dirty
