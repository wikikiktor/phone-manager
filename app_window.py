import tkinter as tk
from datetime import datetime
from tkinter import filedialog, messagebox, ttk

import db
from dialogs import AddEventDialog, ExcelImportDialog, AddNoteDialog


class APP(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Baza Telefonów")
        self.geometry("1100x680")
        self.minsize(950, 600)

        self.selected_phone_id = None
        self.show_deleted_var = tk.BooleanVar(value=False)

        self.sort_column = None
        self.sort_reverse = False

        self.history_sort_column = None
        self.history_sort_reverse = False

        db.init_db()
        self.build_ui()
        self.load_phone_list()
        
    def build_ui(self):
        # Górny pasek: wyszukiwarka + przycisk nowego telefonu
        top_bar = ttk.Frame(self, padding=10)
        top_bar.pack(fill=tk.X)

        ttk.Label(top_bar, text="Szukaj (nr, model, użytkownik, IMEI):").pack(side=tk.LEFT, padx=5)
        self.search_var = tk.StringVar()
        self.search_var.trace_add("write", lambda *args: self.load_phone_list())
        search_entry = ttk.Entry(top_bar, textvariable=self.search_var, width=35)
        search_entry.pack(side=tk.LEFT, padx=5)

        self.chk_trash = ttk.Checkbutton(
            top_bar,
            text="Pokaż kosz (usunięte)",
            variable=self.show_deleted_var,
            command=self.on_toggle_trash_view,
        )
        self.chk_trash.pack(side=tk.LEFT, padx=10)

        ttk.Button(top_bar, text="+ Nowy telefon", command=self.prepare_new_phone).pack(side=tk.RIGHT, padx=5)
        ttk.Button(top_bar, text="Eksportuj do Excela", command=self.export_to_excel).pack(side=tk.RIGHT, padx=5)
        ttk.Button(top_bar, text="Importuj z Excela", command=self.open_excel_import).pack(side=tk.RIGHT, padx=5)

        # Główny podział
        main_paned = ttk.PanedWindow(self, orient=tk.HORIZONTAL)
        main_paned.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)

        # --- LEWY PANEL (Lista urządzeń) ---
        left_frame = ttk.Frame(main_paned, width=380)
        main_paned.add(left_frame, weight=1)

        self.lbl_count = ttk.Label(
            left_frame, 
            text="Łącznie telefonów: 0", 
            anchor=tk.W, 
            font=("Calibri", 9, "italic"),
            padding=(6, 6)
        )
        self.lbl_count.pack(side=tk.BOTTOM, fill=tk.X)

        cols = ("nr_tel", "uzytkownik", "model")
        self.tree = ttk.Treeview(left_frame, columns=cols, show="headings", selectmode="browse")

        self.col_titles = {
            "nr_tel": "Nr telefonu",
            "uzytkownik": "Użytkownik",
            "model": "Model",
        }
        for col, title in self.col_titles.items():
            self.tree.heading(col, text=title, command=lambda c=col: self.on_phone_column_click(c))

        self.tree.column("nr_tel", width=105)
        self.tree.column("uzytkownik", width=120)
        self.tree.column("model", width=120)

        scrollbar = ttk.Scrollbar(left_frame, orient=tk.VERTICAL, command=self.tree.yview)
        self.tree.configure(yscrollcommand=scrollbar.set)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.tree.pack(fill=tk.BOTH, expand=True)
        self.tree.bind("<<TreeviewSelect>>", self.on_phone_select)

        # --- PRAWY PANEL (Szczegóły + Historia) ---
        right_frame = ttk.Frame(main_paned)
        main_paned.add(right_frame, weight=2)

        # Sekcja 1: Dane bieżące
        details_box = ttk.LabelFrame(right_frame, text="Dane bieżące telefonu", padding=10)
        details_box.pack(fill=tk.X, padx=5, pady=5)

        self.entries = {}
        fields = [
            ("Model tel:", "model", 0, 0),
            ("Nr telefonu:", "nr_tel", 0, 2),
            ("Nr SIM:", "nr_sim", 1, 0),
            ("Nr IMEI:", "imei", 1, 2),
            ("Nr seryjny:", "nr_seryjny", 2, 0),
            ("Rodzaj użytkowania:", "rodzaj", 2, 2),
            ("Osoba użytkująca:", "osoba_uzytkujaca", 3, 0),
            ("Osoba odpowiedzialna:", "osoba_odpowiedzialna", 3, 2),
        ]

        for label_text, key, r, c in fields:
            ttk.Label(details_box, text=label_text).grid(row=r, column=c, sticky=tk.W, padx=5, pady=3)
            if key == "rodzaj":
                ent = ttk.Combobox(
                    details_box,
                    values=["Montage/Service"],
                )
            else:
                ent = ttk.Entry(details_box, width=28)

                if key == "nr_tel":
                    ent.bind("<FocusOut>", self._format_phone_entry)
            ent.grid(row=r, column=c + 1, sticky=tk.EW, padx=5, pady=3)
            self.entries[key] = ent

        details_box.columnconfigure(1, weight=1)
        details_box.columnconfigure(3, weight=1)

        btn_bar = ttk.Frame(details_box)
        btn_bar.grid(row=4, column=0, columnspan=4, pady=10, sticky=tk.E)

        self.btn_delete = ttk.Button(btn_bar, text="Przenieś do kosza", command=self.soft_delete_phone)
        self.btn_delete.pack(side=tk.LEFT, padx=5)

        self.btn_save = ttk.Button(btn_bar, text="Zapisz zmiany", command=self.save_phone)
        self.btn_save.pack(side=tk.LEFT, padx=5)

        # Przyciski trybu kosza (początkowo ukryte)
        self.btn_restore = ttk.Button(btn_bar, text="Przywróć telefon", command=self.restore_phone)
        self.btn_hard_delete = ttk.Button(btn_bar, text="Usuń trwale z bazy", command=self.hard_delete_phone)

        # Sekcja 2: Historia i uwagi
        history_box = ttk.LabelFrame(right_frame, text="Dziennik zdarzeń i uwagi", padding=10)
        history_box.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)

        self.h_top = ttk.Frame(history_box)
        self.h_top.pack(fill=tk.X, pady=(0, 5))

        # Przyciski akcji historii
        self.btn_add_event = ttk.Button(self.h_top, text="+ Dodaj wpis", command=self.open_add_event_popup)
        self.btn_add_event.pack(side=tk.RIGHT)

        # Przycisk uwagi: początkowo ukryty, pojawi się obok po kliknięciu wpisu
        self.btn_add_note = ttk.Button(self.h_top, text="+ Dodaj / edytuj uwagę", command=self.open_add_note_popup)

        # Tabela historii: 'uwagi' po prawej stronie od 'opis'
        self.history_tree = ttk.Treeview(
            history_box,
            columns=("data", "kategoria", "opis", "uwagi"),
            show="headings",
            selectmode="browse",
        )

        self.history_col_titles = {
            "data": "Data",
            "kategoria": "Kategoria",
            "opis": "Opis zdarzenia",
            "uwagi": "Uwagi",
        }
        for col, title in self.history_col_titles.items():
            self.history_tree.heading(col, text=title, command=lambda c=col: self.on_history_column_click(c))
        
        self.history_tree.column("data", width=120, stretch=False)
        self.history_tree.column("kategoria", width=110, stretch=False)
        self.history_tree.column("opis", width=250)
        self.history_tree.column("uwagi", width=220)

        # Zdarzenie kliknięcia wiersza w historii
        self.history_tree.bind("<<TreeviewSelect>>", self.on_history_select)

        h_scroll = ttk.Scrollbar(history_box, orient=tk.VERTICAL, command=self.history_tree.yview)
        self.history_tree.configure(yscrollcommand=h_scroll.set)
        h_scroll.pack(side=tk.RIGHT, fill=tk.Y)
        self.history_tree.pack(fill=tk.BOTH, expand=True)

    def on_history_select(self, _event=None):
        selected = self.history_tree.selection()
        if selected:
            self.btn_add_note.pack(side=tk.RIGHT, padx=(0, 6))
        else:
            self.btn_add_note.pack_forget()

    def open_add_note_popup(self):
        selected = self.history_tree.selection()
        if not selected:
            messagebox.showwarning("Wybierz wpis", "Wybierz wpis z historii, aby dodać lub edytować uwagę.")
            return

        history_id = int(selected[0])
        current_values = self.history_tree.item(selected[0], "values")
        current_note = current_values[3] if len(current_values) > 3 else ""

        AddNoteDialog(
            parent=self,
            history_id=history_id,
            current_note=current_note,
            on_save_callback=lambda: self._after_note_saved(history_id)
        )

    def _after_note_saved(self, history_id):
        self.refresh_selected_details()
        if self.history_tree.exists(str(history_id)):
            self.history_tree.selection_set(str(history_id))
            self.btn_add_note.pack(side=tk.RIGHT, padx=(0, 6))

    def on_phone_column_click(self, col):
        if self.sort_column == col:
            self.sort_reverse = not self.sort_reverse
        else:
            self.sort_column = col
            self.sort_reverse = False

        self.apply_phone_sorting()

    def apply_phone_sorting(self):
        for col, title in self.col_titles.items():
            if col == self.sort_column:
                arrow = "  ▼" if self.sort_reverse else "  ▲"
                self.tree.heading(col, text=f"{title}{arrow}")
            else:
                self.tree.heading(col, text=title)

        if not self.sort_column:
            return

        items = [(self.tree.set(k, self.sort_column), k) for k in self.tree.get_children("")]
        items.sort(key=lambda t: t[0].lower(), reverse=self.sort_reverse)

        for index, (_, k) in enumerate(items):
            self.tree.move(k, "", index)

    def on_history_column_click(self, col):
        if self.history_sort_column == col:
            self.history_sort_reverse = not self.history_sort_reverse
        else:
            self.history_sort_column = col
            self.history_sort_reverse = False

        self.apply_history_sorting()

    def apply_history_sorting(self):
        for col, title in self.history_col_titles.items():
            if col == self.history_sort_column:
                arrow = "  ▼" if self.history_sort_reverse else "  ▲"
                self.history_tree.heading(col, text=f"{title}{arrow}")
            else:
                self.history_tree.heading(col, text=title)

        if not self.history_sort_column:
            return

        items = [(self.history_tree.set(k, self.history_sort_column), k) for k in self.history_tree.get_children("")]
        items.sort(key=lambda t: t[0].lower(), reverse=self.history_sort_reverse)

        for index, (_, k) in enumerate(items):
            self.history_tree.move(k, "", index)

    def load_phone_list(self):
        query = self.search_var.get().strip()
        is_trash = self.show_deleted_var.get()
        rows = db.search_phones(query, show_deleted=is_trash)

        self.tree.delete(*self.tree.get_children())
        for row in rows:
            self.tree.insert(
                "",
                tk.END,
                iid=str(row[0]),
                values=(row[1], row[2] or "[BRAK]", row[3])
            )

        total = db.get_phones_count(show_deleted=is_trash)

        prefix = "W koszu:" if is_trash else "Łącznie aktywnych:"
        if query:
            self.lbl_count.config(text=f"{prefix} znaleziono {len(rows)} z {total}")
        else:
            self.lbl_count.config(text=f"{prefix} {total} telefonów")

        if hasattr(self, "sort_column") and self.sort_column:
            self.apply_phone_sorting()

    def refresh_selected_details(self):
        if not self.selected_phone_id:
            return

        row = db.get_phone_by_id(self.selected_phone_id)
        if row:
            keys = [
                "model", "nr_tel", "nr_sim", "imei",
                "nr_seryjny", "rodzaj", "osoba_uzytkujaca", "osoba_odpowiedzialna"
            ]
            for i, key in enumerate(keys):
                self.entries[key].delete(0, tk.END)
                if row[i]:
                    self.entries[key].insert(0, row[i])

        history_rows = db.get_phone_history(self.selected_phone_id)
        self.history_tree.delete(*self.history_tree.get_children())
        self.btn_add_note.pack_forget()

        for h_id, h_data, h_kat, h_opis, h_uwagi in history_rows:
            self.history_tree.insert(
                "",
                tk.END,
                iid=str(h_id),
                values=(h_data, h_kat, h_opis, h_uwagi)
            )

    def on_phone_select(self, _event):
        selected = self.tree.selection()
        if not selected:
            return
        self.selected_phone_id = int(selected[0])
        self.refresh_selected_details()

    def prepare_new_phone(self):
        self.selected_phone_id = None
        self.tree.selection_remove(*self.tree.selection())
        for entry in self.entries.values():
            entry.delete(0, tk.END)
        self.history_tree.delete(*self.history_tree.get_children())
        self.btn_add_note.pack_forget()
        self.entries["model"].focus()

    def save_phone(self):
        self._format_phone_entry() 
        data = {k: ent.get().strip() for k, ent in self.entries.items()}

        if not data["model"] or not data["nr_tel"]:
            messagebox.showwarning("Błąd", "Model i Nr Tel są wymagane.")
            return

        if self.selected_phone_id is None:
            self.selected_phone_id = db.insert_phone(data)
            messagebox.showinfo("Sukces", "Telefon został dodany do bazy.")
        else:
            db.update_phone(self.selected_phone_id, data)
            messagebox.showinfo("Sukces", "Dane telefonu zostały zaktualizowane.")

        self.load_phone_list()
        self.tree.selection_set(str(self.selected_phone_id))
        self.refresh_selected_details()

    def soft_delete_phone(self):
        if not self.selected_phone_id:
            messagebox.showwarning("Wybierz telefon", "Nie wybrano telefonu do usunięcia.")
            return

        if messagebox.askyesno("Potwierdzenie", "Czy na pewno chcesz przenieść ten telefon do kosza?"):
            db.soft_delete_phone(self.selected_phone_id)
            messagebox.showinfo("Kosz", "Telefon został przeniesiony do kosza.")
            self.prepare_new_phone()
            self.load_phone_list()

    def restore_phone(self):
        if not self.selected_phone_id:
            messagebox.showwarning("Wybierz telefon", "Wybierz telefon z kosza, który chcesz przywrócić.")
            return

        if messagebox.askyesno("Potwierdzenie", "Czy chcesz przywrócić ten telefon do aktywnych urządzeń?"):
            db.restore_phone(self.selected_phone_id)
            messagebox.showinfo("Sukces", "Telefon został przywrócony.")
            self.prepare_new_phone()
            self.load_phone_list()

    def hard_delete_phone(self):
        if not self.selected_phone_id:
            messagebox.showwarning("Wybierz telefon", "Wybierz telefon do trwałego usunięcia.")
            return

        msg = (
            "UWAGA: Ta operacja jest NIEODWRACALNA!\n"
            "Telefon oraz cała jego historia zdarzeń zostaną trwale wykasowane z bazy.\n\n"
            "Czy na pewno chcesz kontynuować?"
        )
        if messagebox.askyesno("Ostrzeżenie", msg, icon=messagebox.WARNING):
            db.hard_delete_phone(self.selected_phone_id)
            messagebox.showinfo("Usunięto", "Telefon został trwale wykasowany z bazy.")
            self.prepare_new_phone()
            self.load_phone_list()

    def open_add_event_popup(self):
        if not self.selected_phone_id:
            messagebox.showwarning("Wybierz telefon", "Wybierz lub zapisz telefon przed dodaniem zdarzenia.")
            return

        AddEventDialog(self, self.selected_phone_id, on_save_callback=self.refresh_selected_details)

    def open_excel_import(self):
        file_path = filedialog.askopenfilename(
            title="Wybierz plik Excel do importu",
            filetypes=[("Excel files", "*.xlsx *.xls *.xlsm"), ("All files", "*.*")],
        )
        if not file_path:
            return

        ExcelImportDialog(self, file_path, on_success_callback=self.load_phone_list)

    def _format_phone_entry(self, _event=None):
        current = self.entries["nr_tel"].get()
        formatted = db.format_phone_number(current)
        if current != formatted:
            self.entries["nr_tel"].delete(0, tk.END)
            self.entries["nr_tel"].insert(0, formatted)

    def export_to_excel(self):
        visible_ids = [int(item_id) for item_id in self.tree.get_children()]

        if not visible_ids:
            messagebox.showwarning(
                "Brak danych",
                "Lista jest pusta. Brak wyników do wyeksportowania.",
            )
            return

        is_trash = self.show_deleted_var.get()
        prefix = "kosz" if is_trash else "telefony"
        default_filename = f"{prefix}_eksport_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx"

        file_path = filedialog.asksaveasfilename(
            title="Wybierz miejsce zapisu pliku Excel",
            defaultextension=".xlsx",
            initialfile=default_filename,
            filetypes=[("Pliki Excel", "*.xlsx"), ("Wszystkie pliki", "*.*")],
        )
        if not file_path:
            return

        try:
            exported_count = db.export_to_excel(file_path, phone_ids=visible_ids)
            messagebox.showinfo(
                "Eksport zakończony",
                f"Pomyślnie wyeksportowano {exported_count} urządzeń wraz z historią do pliku:\n{file_path}",
            )
        except PermissionError:
            messagebox.showerror(
            "Błąd zapisu",
            "Nie można zapisać pliku. Upewnij się, że plik nie jest obecnie otwarty w programie Excel.",
            )
        except Exception as e:
            messagebox.showerror("Błąd eksportu", f"Wystąpił nieoczekiwany błąd:\n{e}")

    def on_toggle_trash_view(self):
        self.prepare_new_phone()
        is_trash = self.show_deleted_var.get()

        if is_trash:
            self.btn_save.pack_forget()
            self.btn_delete.pack_forget()
            self.btn_restore.pack(side=tk.LEFT, padx=5)
            self.btn_hard_delete.pack(side=tk.LEFT, padx=5)
        else:
            self.btn_restore.pack_forget()
            self.btn_hard_delete.pack_forget()
            self.btn_delete.pack(side=tk.LEFT, padx=5)
            self.btn_save.pack(side=tk.LEFT, padx=5)
        
        self.load_phone_list()