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
        self.selected_phone_id = None
        self.tree.selection_remove(*self.tree.selection())
        for entry in self.entries.values():
            entry.delete(0, tk.END)
        self.history_tree.delete(*self.history_tree.get_children())
        self.entries["model"].focus()

    def save_phone(self):
        data = {k: ent.get().strip() for k, ent in self.entries.items()}
        if not data["model"] or not data["nr_tel"]:
            messagebox.showwarning(
                "Błąd", "Model i Nr Tel są wymagane."
            )
            return

        conn = sqlite3.connect(DB_NAME)
        cursor = conn.cursor()

        if self.selected_phone_id is None:
            cursor.execute(
                """
                INSERT INTO telefony (model, nr_tel, nr_sim, imei, nr_seryjny, rodzaj, osoba_uzytkujaca, osoba_odpowiedzialna)
                VALUES (?, ?, ?, ?, ?, ?, ?, ?)
                """,
                (
                    data["model"],
                    data["nr_tel"],
                    data["nr_sim"],
                    data["imei"],
                    data["nr_seryjny"],
                    data["rodzaj"],
                    data["osoba_uzytkujaca"],
                    data["osoba_odpowiedzialna"],
                ),
            )
            self.selected_phone_id = cursor.lastrowid

            cursor.execute(
                """
                INSERT INTO historia (telefon_id, data, kategoria, opis)
                VALUES (?, ?, ?, ?)
                """,
                (
                    self.selected_phone_id,
                    datetime.now().strftime("%Y-%m-%d %H:%M"),
                    "Dodanie do bazy",
                    "Zarejestrowano urządzenie w systemie.",
                ),
            )
            messagebox.showinfo("Sukces", "Telefon został dodany do bazy.")
        else:
            cursor.execute(
                """
                UPDATE telefony
                SET model = ?, nr_tel = ?, nr_sim = ?, imei = ?, nr_seryjny = ?, rodzaj = ?, osoba_uzytkujaca = ?, osoba_odpowiedzialna = ?
                WHERE id = ?
                """,
                (
                    data["model"],
                    data["nr_tel"],
                    data["nr_sim"],
                    data["imei"],
                    data["nr_seryjny"],
                    data["rodzaj"],
                    data["osoba_uzytkujaca"],
                    data["osoba_odpowiedzialna"],
                    self.selected_phone_id,
                ),
            )
            cursor.execute(
                """
                INSERT INTO historia (telefon_id, data, kategoria, opis)
                VALUES (?, ?, ?, ?)
                """,
                (
                    self.selected_phone_id,
                    datetime.now().strftime("%Y-%m-%d %H:%M"),
                    "Edycja danych",
                    "Zaktualizowano dane urządzenia.",
                ),
            )
            messagebox.showinfo("Sukces", "Dane telefonu zostały zaktualizowane.")

        conn.commit()
        conn.close()

        self.load_phone_list
        self.tree.selection_set(str(self.selected_phone_id))

    def delete_phone(self):
        if not self.selected_phone_id:
            messagebox.showwarning(
                "Wybierz telefon", "Nie wybrano telefonu do usunięcia."
            )
            return

        if messagebox.askyesno(
            "Potwierdzenie", "Czy na pewno chcesz usunąć ten telefon i całą jego historię?"
        ):
            conn = sqlite3.connect(DB_NAME)
            cursor = conn.cursor()
            cursor.execute(
                "DELETE FROM historia WHERE telefon_id = ?", 
                (self.selected_phone_id,)
            )
            cursor.execute(
                "DELETE FROM telefony WHERE id = ?", 
                (self.selected_phone_id,)
            )
            conn.commit()
            conn.close()

            messagebox.showinfo("Sukces", "Telefon został usunięty z bazy.")
            self.prepare_new_phone()
            self.load_phone_list()

    def open_add_event_popup(self):
        if not self.selected_phone_id:
            messagebox.showwarning(
            "Wybierz telefon", 
            "Wybierz lub zapisz telefon przed dodaniem zdarzenia."
            )
            return

        popup = tk.Toplevel(self)
        popup.title("Dodaj zdarzenie / uwagę")
        popup.geometry("450x300")
        popup.transient(self)
        popup.grab_set()

        ttk.Label(popup, text="Kategoria:").pack(anchor=tk.W, padx=15, pady=(15,2))
        cat_combo = ttk.Combobox(
            popup,
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
        cat_combo.set("Notatka / Uwaga")
        cat_combo.pack(fill=tk.X, padx=15)

        ttk.Label(popup, text="Opis zdarzenia / uwagi:").pack(
            anchor=tk.W, padx=15, pady=(10,2)
        )
        txt = tk.Text(popup, height=5, wrap=tk.WORD)
        txt.pack(fill=tk.both,expand=True, padx=15, pady=5)
        txt.focus()
        
        def save_event():
            opis = txt.get("1.0", tk.END).strip()
            if not opis:
                messagebox.showwarning(
                    "Puste pole", 
                    "Wpisz treść zdarzenia lub uwagi."
                )
                return

            now_str = datetime.now().strftime("%Y-%m-%d %H:%M")
            conn = sqlite3.connect(DB_NAME)
            cursor = conn.cursor()
            cursor.execute(
                """
                INSERT INTO historia (telefon_id, data, kategoria, opis)
                VALUES (?, ?, ?, ?)
                """,
                (
                    self.selected_phone_id,
                    now_str,
                    cat_combo.get(),
                    opis,
                ),
            )
            conn.commit()
            conn.close()

            popup.destroy()
            self.on_phone_select(None)

        ttk.Button(popup, text="Zapisz wpis", command=save_event).pack(
            pady=10, padx=15, anchor=tk.E
        )
        

if __name__ == "__main__":
    app = APP()
    app.mainloop()