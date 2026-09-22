import sqlite3
import tkinter as tk
from tkinter import messagebox, ttk
from datetime import datetime

DB_NAME = "telefony.db"

def init_db():
    conn = sqlite3.connect(DB_NAME)
    cursor = conn.cursor()

    cursor.execute('''
        CREATE TABLE IF NOT EXISTS telefony (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            model TEXT NOT NULL,
            nr_tel TEXT NOT NULL,
            nr_sim TEXT,
            imei TEXT,
            nr_seryjny TEXT,
            rodzaj TEXT,
            osoba_uzytkujaca TEXT,
            osoba_odpowiedzialna TEXT
        )
    ''')

    cursor.execute('''
        CREATE TABLE IF NOT EXISTS historia (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            telefon_id INTEGER NOT NULL,
            data TEXT NOT NULL,
            kategoria TEXT NOT NULL,
            opis TEXT,
            FOREIGN KEY (telefon_id) REFERENCES telefony (id)
        )''')

    cursor.execute("SELECT COUNT(*) FROM telefony")
    if cursor.fetchone()[0] == 0:
        cursor.execute('''
            INSERT INTO telefony (model, nr_tel, nr_sim, imei, nr_seryjny, rodzaj, osoba_uzytkujaca, osoba_odpowiedzialna, data_dodania)
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)''', 
            (
                "Przykładowy Model",
                "123456789", 
                "SIM123456", 
                "IMEI123456789", 
                "SERIAL123456", 
                "Rodzaj1", 
                "Użytkownik1", 
                "Odpowiedzialny1",
                ),)
        tel_id = cursor.lastrowid
        cursor.execute('''
            INSERT INTO historia (telefon_id, data, kategoria, opis)
            VALUES (?, ?, ?, ?)
        ''', (tel_id, 
              datetime.now().strftime("%Y-%m-%d %H:%M:%S"), 
              "Dodanie", 
              "Telefon dodany do bazy danych",
              ),)

    conn.commit()
    conn.close()

class APP(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Baza Telefonów")
        self.geometry("1100x680")
        self.minsize(950, 600)

        self.selected_phone_id = None
        init_db()
        self.build_ui()
        self.load_phone_list()

    def create_widgets(self):
        self.tree = ttk.Treeview(self, columns=("ID", "Model", "Nr Tel", "Nr SIM", "IMEI", "Nr Seryjny", "Rodzaj", "Osoba Użytk.", "Osoba Odp."), show="headings")
        for col in self.tree["columns"]:
            self.tree.heading(col, text=col)
            self.tree.column(col, width=100)
        self.tree.pack(fill=tk.BOTH, expand=True)

        self.add_button = tk.Button(self, text="Dodaj Telefon", command=self.add_telefon)
        self.add_button.pack(pady=10)

    def populate_treeview(self):
        for row in self.tree.get_children():
            self.tree.delete(row)

        conn = sqlite3.connect(DB_NAME)
        cursor = conn.cursor()
        cursor.execute("SELECT * FROM telefony")
        for row in cursor.fetchall():
            self.tree.insert("", tk.END, values=row)
        conn.close()

    def add_telefon(self):
        AddTelefonWindow(self)