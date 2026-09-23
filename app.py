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
        cursor.execute("""
            INSERT INTO telefony (model, nr_tel, nr_sim, imei, nr_seryjny, rodzaj, osoba_uzytkujaca, osoba_odpowiedzialna)
            VALUES (?, ?, ?, ?, ?, ?, ?, ?)
        """,
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
        self.phone_list_frame = ttk.Frame(self)
        self.phone_list_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        self.phone_list = ttk.Treeview(self.phone_list_frame, columns=("model", "nr_tel", "nr_sim", "imei", "nr_seryjny", "rodzaj", "osoba_uzytkujaca", "osoba_odpowiedzialna"), show="headings")
        self.phone_list.heading("model", text="Model")
        self.phone_list.heading("nr_tel", text="Nr Tel")
        self.phone_list.heading("nr_sim", text="Nr SIM")
        self.phone_list.heading("imei", text="IMEI")
        self.phone_list.heading("nr_seryjny", text="Nr Seryjny")
        self.phone_list.heading("rodzaj", text="Rodzaj")
        self.phone_list.heading("osoba_uzytkujaca", text="Osoba Użytk. ")
        self.phone_list.heading("osoba_odpowiedzialna", text="Osoba Odpow.")
        self.phone_list.pack(fill=tk.BOTH, expand=True)

        self.phone_list.bind("<Double-1>", self.on_phone_select)

        self.details_frame = ttk.Frame(self)
        self.details_frame.pack(side=tk.RIGHT, fill=tk.BOTH, expand=True)

        # Add labels and entry fields for phone details
        labels = ["Model:", "Nr Tel:", "Nr SIM:", "IMEI:", "Nr Seryjny:", "Rodzaj:", "Osoba Użytk.:", "Osoba Odpow.:"]
        self.entries = {}
        
        for i, label in enumerate(labels):
            ttk.Label(self.details_frame, text=label).grid(row=i, column=0, sticky=tk.W, padx=5, pady=5)
            entry = ttk.Entry(self.details_frame)
            entry.grid(row=i, column=1, sticky=tk.EW, padx=5, pady=5)
            self.entries[label[:-1].lower().replace(" ", "_")] = entry

        
    def load_phone_list(self):
        search = f"%{self.search_var.get().strip()}"
        conn = sqlite3.connect(DB_NAME)
        cursor = conn.cursor()
        cursor.execute(
            """
            SELECT id, nr_tel, osoba_uzytkujaca, model FROM telefony
            WHERE nr_tel LIKE ? OR osoba_uzytkujaca LIKE ? OR model LIKE ? OR imei LIKE ?
            ORDER BY id DESC
            """,
            (search, search, search, search)
        )
        rows = cursor.fetchall()
        conn.close()

        self.tree.delete(*self.tree.get_children())
        for row in rows:
            self.tree.insert(
                            "", 
                            tk.END,
                            iid=str(row[0]), 
                            values=(row[1], row[2] or "[BRAK]", row[3])
                            )

    def on_phone_select(self, event):
        selected = self.phone_list.selection()
        if not selected:
            return
        phone_id = int(selected[0])
        self.selected_phone_id = phone_id
        conn = sqlite3.connect(DB_NAME)
        cursor = conn.cursor()
        cursor.execute(
            """
            SELECT model, nr_tel, nr_sim, imei, nr_seryjny, rodzaj, osoba_uzytkujaca, osoba_odpowiedzialna
            FROM telefony
            WHERE id = ?
            """,
            (phone_id,),
        )
        row = cursor.fetchone()
        if row:
            keys = [
                "model", 
                "nr_tel", 
                "nr_sim", 
                "imei", 
                "nr_seryjny", 
                "rodzaj", 
                "osoba_uzytkujaca", 
                "osoba_odpowiedzialna"
            ]
            for i, key in enumerate(keys):
                self.entries[key].delete(0, tk.END)
                if row[i]:
                    self.entries[key].insert(0, row[i])
        cursor.execute(
            """
            SELECT data, kategoria, opis
            FROM historia
            WHERE telefon_id = ?
            ORDER BY id DESC
            """,
            (phone_id,),
        )
        history_rows = cursor.fetchall()
        conn.close()

        self.history_tree.delete(*self.history_tree.get_children())
        for history_row in history_rows:
            self.history_tree.insert("", tk.END, values=history_row)

    def prepare_new_phone(self):
        print("cos2")
    def save_phone(self):
        print("cos3")
    def delete_phone(self):
        print("cos4")
    def open_add_event_popup(self):
        print("cos5")
    def save_event(self):
        print("cos6")

if __name__ == "__main__":
    app = APP()
    app.mainloop()