import sqlite3
from datetime import datetime

DB_NAME = "telefony.db"

def get_connection():
    return sqlite3.connect(DB_NAME)

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
              datetime.now().strftime("%Y-%m-%d %H:%M"), 
              "Dodanie", 
              "Telefon dodany do bazy danych",
              ),)

    conn.commit()
    conn.close()

def search_phones(query_str):
    search = f"%{query_str}%"
    with get_connection() as conn:
        cursor = conn.cursor()
        cursor.execute(
            """
            SELECT id, nr_tel, osoba_uzytkujaca, model FROM telefony
            WHERE nr_tel LIKE ? OR osoba_uzytkujaca LIKE ? OR model LIKE ? OR imei LIKE ?
            ORDER BY id DESC
            """,
            (search, search, search, search)
        )
        return cursor.fetchall()

def get_phone_by_id(phone_id):
    with get_connection() as conn:
        cursor = conn.cursor()
        cursor.execute(
            """
            SELECT model, nr_tel, nr_sim, imei, nr_seryjny, rodzaj, osoba_uzytkujaca, osoba_odpowiedzialna
            FROM telefony
            WHERE id = ?
            """,
            (phone_id,),
        )
        return cursor.fetchone()

def get_phone_history(phone_id):
    with get_connection() as conn:
        cursor = conn.cursor()
        cursor.execute(
            """
            SELECT data, kategoria, opis
            FROM historia
            WHERE telefon_id = ?
            ORDER BY id DESC
            """,
            (phone_id,),
        )
        return cursor.fetchall()

def insert_phone(data):
    with get_connection() as conn:
        cursor = conn.cursor()
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
        phone_id = cursor.lastrowid
        cursor.execute(
            """
            INSERT INTO historia (telefon_id, data, kategoria, opis)
            VALUES (?, ?, ?, ?)
            """,
            (
                phone_id,
                datetime.now().strftime("%Y-%m-%d %H:%M"),
                "Dodanie do bazy",
                "Zarejestrowano urządzenie w systemie.",
            ),
        )
        return phone_id

def update_phone(phone_id, data):
    labels = {
        "model": "Model",
        "nr_tel": "Nr telefonu",
        "nr_sim": "Nr SIM",
        "imei": "IMEI",
        "nr_seryjny": "Nr seryjny",
        "rodzaj": "Rodzaj",
        "osoba_uzytkujaca": "Osoba użytkująca",
        "osoba_odpowiedzialna": "Osoba odpowiedzialna"
    }

    with get_connection() as conn:
        cursor = conn.cursor()

        cursor.execute(
            """
            Select model, nr_tel, nr_sim, imei, nr_seryjny, rodzaj, osoba_uzytkujaca, osoba_odpowiedzialna
            FROM telefony
            WHERE id = ?
            """,
            (phone_id,)
        )
        old_row = cursor.fetchone()

        changes = []
        if old_row:
            keys = [
                "model", "nr_tel", "nr_sim", "imei",
                "nr_seryjny", "rodzaj", "osoba_uzytkujaca", "osoba_odpowiedzialna"
            ]
            for i, key in enumerate(keys):
                old_val = (old_row[i] or "").strip()
                new_val = (data[key] or "").strip()
                if old_val != new_val:
                    old = old_val if old_val else "[puste]"
                    new = new_val if new_val else "[puste]"
                    changes.append(f"{labels[key]}: '{old}' ➔ '{new}'")

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
                phone_id
            )
        )

        if changes:
            opis = "Zmieniono parametry: " + "; ".join(changes)
            kategoria = "Zmiana użytkownika" if any("Osoba użytkująca" in c for c in changes) else "Edycja danych"

            cursor.execute(
                """
                INSERT INTO historia (telefon_id, data, kategoria, opis)
                VALUES (?, ?, ?, ?)
                """,
                (
                    phone_id,
                    datetime.now().strftime("%Y-%m-%d %H:%M"),
                    kategoria,
                    opis,
                ),
            )

def delete_phone_by_id(phone_id):
    with get_connection() as conn:
        cursor = conn.cursor()
        cursor.execute(
            "DELETE FROM historia WHERE telefon_id = ?", 
            (phone_id,)
        )
        cursor.execute(
            "DELETE FROM telefony WHERE id = ?", 
            (phone_id,)
        )

def add_history_entry(phone_id, category, description):
    with get_connection() as conn:
        cursor = conn.cursor()
        cursor.execute(
            """
            INSERT INTO historia (telefon_id, data, kategoria, opis)
            VALUES (?, ?, ?, ?)
            """,
            (
                phone_id,
                datetime.now().strftime("%Y-%m-%d %H:%M"),
                category,
                description,
            ),
        )