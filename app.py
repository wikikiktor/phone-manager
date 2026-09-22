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

    def build_ui(self):
        print("test1")
    def load_phone_list(self):
        print("test2")
    def on_phone_select(self, event):
        print("test3")
    def prepare_new_phone(self):
        print("test4")
    def save_phone(self):
        print("test5")
    def delete_phone(self):
        print("test6")
    def open_add_event_popup(self):
        print("test7")

if __name__ == "__main__":
    app = APP()
    app.mainloop()