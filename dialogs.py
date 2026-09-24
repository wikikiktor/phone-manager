import tkinter as tk
from tkinter import messagebox, ttk

import db


class AddEventDialog(tk.Toplevel):
    def __init__(self, parent, phone_id, on_save_callback):
        super().__init__(parent)
        self.phone_id = phone_id
        self.on_save_callback = on_save_callback

        self.title("Dodaj zdarzenie / uwagę")
        self.geometry("450x300")
        self.transient(parent)
        self.grab_set()

        self.build_ui()

    def build_ui(self):
        ttk.Label(self, text="Kategoria:").pack(anchor=tk.W, padx=15, pady=(15, 2))
        
        self.cat_combo = ttk.Combobox(
            self,
            values=[
                "Notatka / Uwaga",
                "Zmiana użytkownika",
                "Awaria / Serwis",
                "Wydanie",
                "Zwrot",
                "Wymiana karty SIM",
                "Inne",
            ],
            state="readonly",
        )
        self.cat_combo.set("Notatka / Uwaga")
        self.cat_combo.pack(fill=tk.X, padx=15)

        ttk.Label(self, text="Opis zdarzenia / uwagi:").pack(anchor=tk.W, padx=15, pady=(10, 2))
        
        self.txt = tk.Text(self, height=5, wrap=tk.WORD)
        self.txt.pack(fill=tk.BOTH, expand=True, padx=15, pady=5)
        self.txt.focus()

        ttk.Button(self, text="Zapisz wpis", command=self.save_event).pack(
            pady=10, padx=15, anchor=tk.E
        )

    def save_event(self):
        opis = self.txt.get("1.0", tk.END).strip()

        if not opis:
            messagebox.showwarning("Puste pole", "Wpisz treść zdarzenia lub uwagi.")
            return

        db.add_event(self.phone_id, self.cat_combo.get(), opis)
        self.destroy()
        if self.on_save_callback:
            self.on_save_callback()