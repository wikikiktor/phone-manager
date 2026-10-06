import tkinter as tk
from datetime import datetime
from tkinter import filedialog, messagebox, ttk

import openpyxl

import db
import protocol_generator


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
        self.cat_combo.set("Notatka")
        self.cat_combo.pack(fill=tk.X, padx=15)

        ttk.Label(self, text="Opis zdarzenia:").pack(anchor=tk.W, padx=15, pady=(10, 2))
        
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

        db.add_history_entry(self.phone_id, self.cat_combo.get(), opis)
        self.destroy()
        if self.on_save_callback:
            self.on_save_callback()

class AddNoteDialog(tk.Toplevel):
    def __init__(self, parent, history_id, current_note, on_save_callback):
        super().__init__(parent)
        self.history_id = history_id
        self.on_save_callback = on_save_callback

        self.title("Dodaj / Edytuj notatkę")
        self.geometry("400x250")
        self.transient(parent)
        self.grab_set()

        ttk.Label(self, text="Uwaga do zdarzenia:").pack(anchor=tk.W, padx=15, pady=(15, 2))
        self.note_text = tk.Text(self, height=5, wrap=tk.WORD)
        self.note_text.pack(fill=tk.BOTH, expand=True, padx=15, pady=5)
        if current_note:
            self.note_text.insert("1.0", current_note)
        self.note_text.focus()

        btn_box = ttk.Frame(self)
        btn_box.pack(fill=tk.X, padx=15, pady=10)
        ttk.Button(btn_box, text="Anuluj", command=self.destroy).pack(side=tk.RIGHT, padx=5)
        ttk.Button(btn_box, text="Zapisz uwagę", command=self.save_note).pack(side=tk.RIGHT)

    def save_note(self):
        note = self.note_text.get("1.0", tk.END).strip()
        db.update_history_note(self.history_id, note)
        self.destroy()
        if self.on_save_callback:
            self.on_save_callback()

class ExcelImportDialog(tk.Toplevel):
    TARGET_FIELDS = [  # noqa: RUF012
        ("model", "Model telefonu *"),
        ("nr_tel", "Numer telefonu *"),
        ("nr_sim", "Numer SIM"),
        ("imei", "IMEI"),
        ("nr_seryjny", "Numer seryjny"),
        ("rodzaj", "Rodzaj"),
        ("osoba_uzytkujaca", "Osoba użytkująca"),
        ("osoba_odpowiedzialna", "Osoba odpowiedzialna"),
        ("czy_protokol", "Protokół zdawczo-odbiorczy"),
        ("wyposazenie", "Wyposażenie dodatkowe"),
        ("stan_baterii", "Stan baterii"),
    ]

    def __init__(self, parent, file_path, on_success_callback):
        super().__init__(parent)
        self.file_path = file_path
        self.on_success_callback = on_success_callback

        self.title("Import danych z Excela")
        self.geometry("600x400")
        self.transient(parent)
        self.grab_set()

        self.wb = None
        self.sheet = None
        self.header = []
        self.combos = {}
        self.has_history_sheet = False
        self.import_history_var = tk.BooleanVar(value=False)

        if self.load_excel_headers():
            self.build_ui()

    def load_excel_headers(self):
        try:
            self.wb = openpyxl.load_workbook(self.file_path, data_only=True)
            if "Telefony" in self.wb.sheetnames:
                self.sheet = self.wb["Telefony"]
            else:
                self.sheet = self.wb.active

            self.has_history_sheet = "Historia zdarzeń" in self.wb.sheetnames
            if self.has_history_sheet:
                self.import_history_var.set(True)

            first_row = next(self.sheet.iter_rows(values_only=True), None)
            if not first_row:
                messagebox.showerror("Błąd", "Plik Excel jest pusty.")
                self.destroy()
                return False

            self.header = [str(val).strip() if val is not None else f"Kolumna {idx+1}" for idx, val in enumerate(first_row)]
            
            return True
        except Exception as e:  # noqa: BLE001
            messagebox.showerror("Błąd", f"Nie można wczytać pliku Excel: {e}")
            self.destroy()
            return False

    def build_ui(self):
        ttk.Label(
            self,
            text="Przypisz kolumny z arkusza Excel do odpowiednich pól.\nPola oznaczone gwiazdką (*) są wymagane.",
            padding=10,
            justify=tk.LEFT
        ).pack(anchor=tk.W)

        form_frame = ttk.Frame(self, padding=(15, 5))
        form_frame.pack(fill=tk.BOTH, expand=True)

        options = ["[Ignoruj / brak]"] + self.header

        for idx, (field_key, field_label) in enumerate(self.TARGET_FIELDS):
            ttk.Label(form_frame, text=field_label).grid(row=idx, column=0, sticky=tk.W, pady=4, padx=5)
            
            combo = ttk.Combobox(form_frame, values=options, state="readonly", width=28)
            
            # Próba automatycznego dopasowania po nazwie
            matched = False
            for header in self.header:
                clean_field = field_label.lower().replace("*", "").strip()
                if header.lower() in clean_field or clean_field in header.lower():
                    combo.set(header)
                    matched = True
                    break
            if not matched:
                combo.set("[Ignoruj / brak]")

            combo.grid(row=idx, column=1, sticky=tk.EW, pady=4, padx=5)
            self.combos[field_key] = combo

        form_frame.columnconfigure(1, weight=1)

        if self.has_history_sheet:
            chk_frame = ttk.Frame(self, padding=(15, 5))
            chk_frame.pack(fill=tk.X)
            self.chk_history = ttk.Checkbutton(
                chk_frame,
                text="Importuj również historię (znaleziono arkusz 'Historia zdarzeń')",
                variable=self.import_history_var
            )
            self.chk_history.pack(anchor=tk.W)

        btn_box = ttk.Frame(self, padding=10)
        btn_box.pack(fill=tk.X)
        ttk.Button(btn_box, text="Anuluj", command=self.destroy).pack(side=tk.RIGHT, padx=5)
        ttk.Button(btn_box, text="Importuj dane", command=self.import_data).pack(side=tk.RIGHT, padx=5)

    def import_data(self):

        mapping = {}
        for key, combo in self.combos.items():
            selected = combo.get()
            if selected != "[Ignoruj / brak]":
                mapping[key] = self.header.index(selected)

        if "model" not in mapping or "nr_tel" not in mapping:
            messagebox.showwarning(
                "Brakujące mapowanie",
                "Wymagane jest przypisanie kolumn: 'Model telefonu' oraz 'Numer telefonu'.",
                parent=self
            )
            return

        rows = list(self.sheet.iter_rows(values_only=True))
        data_rows = rows[1:]

        phone_records = []
        for row in data_rows:
            if not any(row):
                continue

            record = {}
            for field_key, col_idx in mapping.items():
                val = row[col_idx] if col_idx < len(row) else ""
                record[field_key] = str(val) if val is not None else ""
            phone_records.append(record)

        if not phone_records:
            messagebox.showinfo("Brak danych", "Nie znaleziono żadnych danych do zaimportowania.", parent=self)
            return

        history_records = self._extract_history_records()

        phones_cnt, hist_cnt = db.bulk_insert_phones(phone_records, history_records)

        msg = f"Pomyślnie zaimportowano {phones_cnt} telefon(ów)."
        if history_records is not None:
            msg += f"\nZaimportowano również {hist_cnt} wpis(ów) historii."

        messagebox.showinfo("Sukces", msg, parent=self)
        self.destroy()
        if self.on_success_callback:
            self.on_success_callback()

    def _extract_history_records(self):
        if not self.has_history_sheet or not self.import_history_var.get():
            return None

        ws_hist = self.wb["Historia zdarzeń"]
        rows = list(ws_hist.iter_rows(values_only=True))
        if len(rows) < 2:
            return None

        hist_headers = [str(col).strip().lower() if col is not None else "" for col in rows[0]]

        def find_col(keywords, exclude=None):
            for i, h in enumerate(hist_headers):
                if exclude and any(ex in h for ex in exclude):
                    continue
                if any(k in h for k in keywords):
                    return i
            return None

        idx_data = find_col(["data"])
        # Szukamy kolumny z numerem telefonu, wykluczając kolumnę z modelem:
        idx_nr = find_col(["nr tel", "numer tel", "nr_tel", "telefon"], exclude=["model"])
        idx_kat = find_col(["kategoria"])
        idx_opis = find_col(["opis zdarzenia", "opis"])
        idx_uwagi = find_col(["uwagi", "uwaga"])


        if idx_nr is None:
            return None

        history_records = []
        for r in rows[1:]:
            if not any(r):
                continue

            date_val = r[idx_data] if idx_data is not None and idx_data < len(r) else ""
            if isinstance(date_val, datetime):
                date_str = date_val.strftime("%Y-%m-%d %H:%M")
            else:
                date_str = str(date_val).strip() if date_val is not None else ""

            history_records.append({
                "data": date_str,
                "nr_tel": str(r[idx_nr]) if idx_nr < len(r) and r[idx_nr] is not None else "",
                "kategoria": str(r[idx_kat]).strip() if idx_kat is not None and idx_kat < len(r) and r[idx_kat] is not None else "Import",
                "opis": str(r[idx_opis]).strip() if idx_opis is not None and idx_opis < len(r) and r[idx_opis] is not None else "",
                "uwagi": str(r[idx_uwagi]).strip() if idx_uwagi is not None and idx_uwagi < len(r) and r[idx_uwagi] is not None else "",
            })

        return history_records

class EmployerDialog(tk.Toplevel):
    def __init__(self, parent):
        super().__init__(parent)

        self.title("Dane pracodawcy")
        self.geometry("400x300")
        self.transient(parent)
        self.grab_set()

        self.build_ui()
        self.load_data()

    def build_ui(self):
        frame = ttk.Frame(self, padding=15)
        frame.pack(fill=tk.BOTH, expand=True)

        ttk.Label(frame, text="Nazwa firmy:").grid(row=0, column=0, sticky=tk.W, pady=5)
        self.ent_nazwa = ttk.Entry(frame, width=32)
        self.ent_nazwa.grid(row=0, column=1, sticky=tk.EW, pady=5, padx=5)

        ttk.Label(frame, text="Adres:").grid(row=1, column=0, sticky=tk.W, pady=5)
        self.ent_adres = ttk.Entry(frame, width=32)
        self.ent_adres.grid(row=1, column=1, sticky=tk.EW, pady=5, padx=5)

        ttk.Label(frame, text="NIP:").grid(row=2, column=0, sticky=tk.W, pady=5)
        self.ent_nip = ttk.Entry(frame, width=32)
        self.ent_nip.grid(row=2, column=1, sticky=tk.EW, pady=5, padx=5)

        frame.columnconfigure(1, weight=1)

        btn_box = ttk.Frame(frame)
        btn_box.grid(row=3, column=0, columnspan=2, pady=(15, 0), sticky=tk.E)

        ttk.Button(btn_box, text="Anuluj", command=self.destroy).pack(side=tk.RIGHT, padx=5)
        ttk.Button(btn_box, text="Zapisz", command=self.save_data).pack(side=tk.RIGHT, padx=5)

    def load_data(self):
        data = db.get_employer()
        if data:
            nazwa, adres, nip = data
            if nazwa:
                self.ent_nazwa.insert(0, nazwa)
            if adres:
                self.ent_adres.insert(0, adres)
            if nip:
                self.ent_nip.insert(0, nip)

    def save_data(self):
        nazwa = self.ent_nazwa.get().strip()
        adres = self.ent_adres.get().strip()
        nip = self.ent_nip.get().strip()

        db.save_employer(nazwa, adres, nip)
        messagebox.showinfo("Zapisano", "Dane pracodawcy zostały zapisane.", parent=self)
        self.destroy()

class ProtocolDialog(tk.Toplevel):
    def __init__(self, parent, initial_phone_id=None):
        super().__init__(parent)
        self.initial_phone_id = initial_phone_id

        self.title("Generuj protokół telefonu")
        self.geometry("520x330")
        self.transient(parent)
        self.grab_set()

        self.phone_map = {}
        self.build_ui()

    def build_ui(self):
        container = ttk.Frame(self, padding=15)
        container.pack(fill=tk.BOTH, expand=True)

        ttk.Label(container, text="Typ protokołu:").grid(row=0, column=0, sticky=tk.W, pady=6)
        self.type_combo = ttk.Combobox(
            container,
            values=["Protokół przekazania", "Protokół zwrotu"],
            state="readonly",
            width=36,
        )
        self.type_combo.set("Protokół przekazania")
        self.type_combo.grid(row=0, column=1, sticky=tk.EW, pady=6, padx=5)

        ttk.Label(container, text="Wybierz telefon:").grid(row=1, column=0, sticky=tk.W, pady=6)

        phones = db.search_phones("", show_deleted=False)
        phone_labels = []
        selected_label = ""

        for row in phones:
            p_id, p_nr, p_user, p_model = row
            label = f"{p_nr} | {p_user or '[Brak użytkownika]'}"
            self.phone_map[label] = p_id
            phone_labels.append(label)
            if self.initial_phone_id and p_id == self.initial_phone_id:
                selected_label = label

        self.phone_combo = ttk.Combobox(
            container,
            values=phone_labels,
            state="readonly",
            width=36,
        )
        if selected_label:
            self.phone_combo.set(selected_label)
        elif phone_labels:
            self.phone_combo.set(phone_labels[0])

        self.phone_combo.grid(row=1, column=1, sticky=tk.EW, pady=6, padx=5)
        self.phone_combo.bind("<<ComboboxSelected>>", self.update_preview)

        self.preview_box = ttk.LabelFrame(container, text="Dane wybranego telefonu", padding=8)
        self.preview_box.grid(row=2, column=0, columnspan=2, sticky=tk.EW, pady=10)

        self.lbl_user = ttk.Label(self.preview_box, text="Użytkownik: -")
        self.lbl_user.pack(anchor=tk.W)
        self.lbl_model = ttk.Label(self.preview_box, text="Model: - | IMEI: -")
        self.lbl_model.pack(anchor=tk.W)
        self.lbl_equip = ttk.Label(self.preview_box, text="Wyposażenie: -")
        self.lbl_equip.pack(anchor=tk.W)
        self.lbl_battery = ttk.Label(self.preview_box, text="Stan baterii: -")
        self.lbl_battery.pack(anchor=tk.W)

        self.update_preview()

        btn_bar = ttk.Frame(container)
        btn_bar.grid(row=3, column=0, columnspan=2, pady=(15, 0), sticky=tk.E)

        ttk.Button(btn_bar, text="Anuluj", command=self.destroy).pack(side=tk.RIGHT, padx=4)
        ttk.Button(btn_bar, text="Eksportuj do PDF", command=lambda: self.export_protocol("pdf")).pack(side=tk.RIGHT, padx=4)
        ttk.Button(btn_bar, text="Eksportuj do DOCX", command=lambda: self.export_protocol("docx")).pack(side=tk.RIGHT, padx=4)

    def get_selected_phone_id(self):
        label = self.phone_combo.get()
        return self.phone_map.get(label)

    def update_preview(self, _event=None):
        phone_id = self.get_selected_phone_id()
        if not phone_id:
            return
        p_row = db.get_phone_by_id(phone_id)
        if p_row:
            self.lbl_user.config(text=f"Użytkownik: {p_row[6] or '[BRAK]'}")
            self.lbl_model.config(text=f"Model: {p_row[0]} | IMEI: {p_row[3] or '[BRAK]'}")
            self.lbl_equip.config(text=f"Wyposażenie: {p_row[9] or 'Brak'}")
            self.lbl_battery.config(text=f"Stan baterii: {p_row[10] or 'Brak informacji'}")

    def export_protocol(self, ext):
        phone_id = self.get_selected_phone_id()
        if not phone_id:
            messagebox.showwarning("Brak telefonu", "Wybierz telefon z listy.", parent=self)
            return

        phone_row = db.get_phone_by_id(phone_id)
        employer_row = db.get_employer()

        is_przekazanie = "przekazania" in self.type_combo.get().lower()
        prot_type = "przekazanie" if is_przekazanie else "zwrot"
        prefix = "Protokol_przekazania" if is_przekazanie else "Protokol_zwrotu"

        clean_model = "".join(c for c in (phone_row[0] or "") if c.isalnum() or c in ("-", "_")).strip()
        date_str = datetime.now().strftime("%Y%m%d")
        default_name = f"{prefix}_{clean_model}_{date_str}.{ext}"

        file_types = [("Dokument Word (*.docx)", "*.docx")] if ext == "docx" else [("Dokument PDF (*.pdf)", "*.pdf")]

        file_path = filedialog.asksaveasfilename(
            parent=self,
            title=f"Zapisz protokół ({ext.upper()})",
            defaultextension=f".{ext}",
            initialfile=default_name,
            filetypes=file_types,
        )
        if not file_path:
            return

        try:
            if ext == "docx":
                protocol_generator.generate_docx(file_path, phone_row, employer_row, prot_type)
            else:
                protocol_generator.generate_pdf(file_path, phone_row, employer_row, prot_type)

            messagebox.showinfo(
                "Sukces",
                f"Protokół został pomyślnie wygenerowany i zapisany:\n{file_path}",
                parent=self,
            )
            self.destroy()
        except PermissionError:
            messagebox.showerror(
                "Błąd zapisu",
                "Plik jest otwarty w innym programie (np. Word / Adobe Reader). Zamknij go i spróbuj ponownie.",
                parent=self,
            )
        except Exception as e:
            messagebox.showerror("Błąd", f"Wystąpił błąd podczas generowania protokołu:\n{e}", parent=self)