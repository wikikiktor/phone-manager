import sqlite3
import tkinter as tk
from datetime import datetime
from tkinter import filedialog, messagebox, ttk

import db
from dialogs import (
    AddEventDialog,
    AddNoteDialog,
    EmployerDialog,
    ExcelImportDialog,
    ProtocolDialog,
)


class APP(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Baza Telefonów Służbowych")
        self.geometry("1180x720")
        self.minsize(1020, 640)

        self.selected_phone_id = None
        self.show_deleted_var = tk.BooleanVar(value=False)

        self.sort_column = None
        self.sort_reverse = False

        self.history_sort_column = None
        self.history_sort_reverse = False

        self.var_protocol = tk.BooleanVar(value=False)
        self.wyposazenie_options = ["Ładowarka", "Etui", "Kabel USB", "Szkło ochronne"]
        self.wyposazenie_vars = {
            opt: tk.BooleanVar(value=False) for opt in self.wyposazenie_options
        }

        db.init_db()
        self.setup_styles()
        self.build_ui()
        self.load_phone_list()

    def setup_styles(self):
        self.style = ttk.Style(self)
        try:
            self.style.theme_use("clam")
        except Exception:
            pass

        # Paleta barw
        self.BG_APP = "#F1F5F9"        # Jasne tło okna
        self.BG_CARD = "#FFFFFF"       # Białe tło kart
        self.BORDER_COLOR = "#CBD5E1"  # Ramki
        self.TEXT_MAIN = "#0F172A"     # Ciemny tekst główny
        self.TEXT_MUTED = "#64748B"    # Wygaszony tekst

        self.configure(bg=self.BG_APP)

        # Bazowe konfiguracje
        self.style.configure(".", background=self.BG_APP, foreground=self.TEXT_MAIN, font=("Segoe UI", 9))
        self.style.configure("TFrame", background=self.BG_APP)
        self.style.configure("Card.TFrame", background=self.BG_CARD)
        self.style.configure("TLabel", background=self.BG_APP, foreground=self.TEXT_MAIN, font=("Segoe UI", 9))
        self.style.configure("Card.TLabel", background=self.BG_CARD, foreground=self.TEXT_MAIN, font=("Segoe UI", 9))

        # Nagłówki
        self.style.configure("BrandTitle.TLabel", background=self.BG_APP, foreground="#1E3A8A", font=("Segoe UI", 13, "bold"))
        self.style.configure("BrandSub.TLabel", background=self.BG_APP, foreground=self.TEXT_MUTED, font=("Segoe UI", 8))
        self.style.configure("CardTitle.TLabel", background=self.BG_CARD, foreground="#1E293B", font=("Segoe UI", 10, "bold"))
        self.style.configure("StatusInfo.TLabel", background=self.BG_CARD, foreground=self.TEXT_MUTED, font=("Segoe UI", 8, "italic"))

        # Plakietka licznika (Badge)
        self.style.configure(
            "Badge.TLabel",
            background="#E0E7FF",
            foreground="#3730A3",
            font=("Segoe UI", 9, "bold"),
            padding=(10, 4),
        )

        # Ramki LabelFrame
        self.style.configure(
            "Card.TLabelframe",
            background=self.BG_CARD,
            bordercolor=self.BORDER_COLOR,
            relief="solid",
            borderwidth=1,
        )
        self.style.configure(
            "Card.TLabelframe.Label",
            background=self.BG_CARD,
            foreground="#1E3A8A",
            font=("Segoe UI", 10, "bold"),
        )

        # Pola tekstowe i listy rozwijane
        self.style.configure("TEntry", fieldbackground="#FFFFFF", padding=5)
        self.style.configure("TCombobox", fieldbackground="#FFFFFF", padding=4)
        self.style.map("TCombobox", fieldbackground=[("readonly", "#FFFFFF")])

        # Checkbuttony
        self.style.configure("TCheckbutton", background=self.BG_APP, font=("Segoe UI", 9))
        self.style.configure("Card.TCheckbutton", background=self.BG_CARD, font=("Segoe UI", 9))

        # --- PRZYCISKI ---
        # 1. Primary Button (Akcja główna, np. Nowy telefon)
        self.style.configure(
            "Primary.TButton",
            background="#1E40AF",
            foreground="#FFFFFF",
            font=("Segoe UI", 9, "bold"),
            borderwidth=0,
            padding=(10, 6),
        )
        self.style.map(
            "Primary.TButton",
            background=[("active", "#1D4ED8"), ("disabled", "#94A3B8")],
            foreground=[("disabled", "#F1F5F9")],
        )

        # 2. Success Button (Zapisz zmiany)
        self.style.configure(
            "Success.TButton",
            background="#059669",
            foreground="#FFFFFF",
            font=("Segoe UI", 9, "bold"),
            borderwidth=0,
            padding=(10, 6),
        )
        self.style.map(
            "Success.TButton",
            background=[("active", "#047857"), ("disabled", "#94A3B8")],
            foreground=[("disabled", "#F1F5F9")],
        )

        # 3. Danger Button (Kosz / Usuń)
        self.style.configure(
            "Danger.TButton",
            background="#EF4444",
            foreground="#FFFFFF",
            font=("Segoe UI", 9, "bold"),
            borderwidth=0,
            padding=(9, 5),
        )
        self.style.map(
            "Danger.TButton",
            background=[("active", "#DC2626"), ("disabled", "#FCA5A5")],
        )

        # 4. Secondary Button (Standardowy szary/biały z ramką)
        self.style.configure(
            "Secondary.TButton",
            background="#FFFFFF",
            foreground="#1E293B",
            font=("Segoe UI", 9),
            borderwidth=1,
            bordercolor=self.BORDER_COLOR,
            padding=(8, 5),
        )
        self.style.map(
            "Secondary.TButton",
            background=[("active", "#F8FAFC")],
        )

        # --- TABELE (Treeview) ---
        self.style.configure(
            "Treeview",
            background="#FFFFFF",
            fieldbackground="#FFFFFF",
            foreground=self.TEXT_MAIN,
            rowheight=26,
            font=("Segoe UI", 9),
            bordercolor=self.BORDER_COLOR,
            borderwidth=1,
        )
        self.style.configure(
            "Treeview.Heading",
            background="#E2E8F0",
            foreground="#0F172A",
            font=("Segoe UI", 9, "bold"),
            padding=(6, 6),
            relief="flat",
        )
        self.style.map(
            "Treeview.Heading",
            background=[("active", "#CBD5E1")],
        )
        self.style.map(
            "Treeview",
            background=[("selected", "#DBEAFE")],
            foreground=[("selected", "#1E3A8A")],
        )

    def build_ui(self):
        # =========================================================================
        # 1. GÓRNY PASEK NARZĘDZIOWY (2 wiersze: Pasek akcji + Pasek wyszukiwania)
        # =========================================================================
        header_container = ttk.Frame(self, padding=(14, 10, 14, 6))
        header_container.pack(fill=tk.X)

        # --- Wiersz 1: Logo i główne akcje ---
        top_row = ttk.Frame(header_container)
        top_row.pack(fill=tk.X, pady=(0, 8))

        brand_frame = ttk.Frame(top_row)
        brand_frame.pack(side=tk.LEFT)
        ttk.Label(brand_frame, text="📱 Baza Telefonów", style="BrandTitle.TLabel").pack(anchor=tk.W)
        ttk.Label(brand_frame, text="System ewidencji sprzętu i protokołów", style="BrandSub.TLabel").pack(anchor=tk.W)

        # Przyciski akcji z prawej
        btn_actions = ttk.Frame(top_row)
        btn_actions.pack(side=tk.RIGHT)

        ttk.Button(btn_actions, text="+ Nowy telefon", style="Primary.TButton", command=self.prepare_new_phone).pack(side=tk.LEFT, padx=3)
        ttk.Button(btn_actions, text="📋 Protokół", style="Secondary.TButton", command=self.open_protocol_dialog).pack(side=tk.LEFT, padx=3)
        ttk.Button(btn_actions, text="📊 Eksport Excel", style="Secondary.TButton", command=self.export_to_excel).pack(side=tk.LEFT, padx=3)
        ttk.Button(btn_actions, text="📥 Import Excel", style="Secondary.TButton", command=self.open_excel_import).pack(side=tk.LEFT, padx=3)

        # Menu narzędzi bazy (zamiast wielu uciętych przycisków)
        self.db_menu = tk.Menu(self, tearoff=0)
        self.db_menu.add_command(label="🏢 Dane firmy / pracodawcy", command=self.open_employer_dialog)
        self.db_menu.add_separator()
        self.db_menu.add_command(label="💾 Utwórz kopię zapasową bazy (.db)", command=self.create_db_backup)
        self.db_menu.add_command(label="🔄 Przywróć bazę z pliku (.db)", command=self.restore_db_backup)

        self.btn_db_tools = ttk.Button(
            btn_actions,
            text="⚙️ Opcje bazy ▾",
            style="Secondary.TButton",
            command=self.show_db_tools_menu,
        )
        self.btn_db_tools.pack(side=tk.LEFT, padx=3)

        # --- Wiersz 2: Wyszukiwanie, filtr kosza i licznik ---
        search_row = ttk.Frame(header_container)
        search_row.pack(fill=tk.X)

        search_box = ttk.Frame(search_row)
        search_box.pack(side=tk.LEFT, fill=tk.X)

        ttk.Label(search_box, text="🔍", font=("Segoe UI", 10)).pack(side=tk.LEFT, padx=(0, 4))
        self.search_var = tk.StringVar()
        self.search_var.trace_add("write", lambda *args: self.load_phone_list())
        self.search_entry = ttk.Entry(search_box, textvariable=self.search_var, width=36)
        self.search_entry.pack(side=tk.LEFT, padx=(0, 4))

        # Przycisk czyszczenia wyszukiwarki
        ttk.Button(search_box, text="✕", width=2, style="Secondary.TButton", command=self.clear_search).pack(side=tk.LEFT, padx=(0, 10))

        self.chk_trash = ttk.Checkbutton(
            search_box,
            text="Pokaż kosz (usunięte)",
            variable=self.show_deleted_var,
            command=self.on_toggle_trash_view,
        )
        self.chk_trash.pack(side=tk.LEFT)

        # Plakietka licznika
        self.lbl_count = ttk.Label(search_row, text="Łącznie telefonów: 0", style="Badge.TLabel")
        self.lbl_count.pack(side=tk.RIGHT)

        # =========================================================================
        # 2. GŁÓWNY PODZIAŁ (PANED WINDOW: LEWY I PRAWY PANEL)
        # =========================================================================
        main_paned = ttk.PanedWindow(self, orient=tk.HORIZONTAL)
        main_paned.pack(fill=tk.BOTH, expand=True, padx=14, pady=(4, 12))

        # --- LEWY PANEL: LISTA TELEFONÓW ---
        left_frame = ttk.Frame(main_paned, width=390)
        main_paned.add(left_frame, weight=1)

        tree_card = ttk.LabelFrame(left_frame, text="  Lista urządzeń  ", style="Card.TLabelframe", padding=6)
        tree_card.pack(fill=tk.BOTH, expand=True)

        cols = ("nr_tel", "uzytkownik", "model")
        self.tree = ttk.Treeview(tree_card, columns=cols, show="headings", selectmode="browse")

        self.col_titles = {
            "nr_tel": "Nr telefonu",
            "uzytkownik": "Użytkownik",
            "model": "Model",
        }
        for col, title in self.col_titles.items():
            self.tree.heading(col, text=title, command=lambda c=col: self.on_phone_column_click(c))

        self.tree.column("nr_tel", width=115, minwidth=90, stretch=False)
        self.tree.column("uzytkownik", width=135, minwidth=100)
        self.tree.column("model", width=125, minwidth=90)

        # Paski zebry w tabeli
        self.tree.tag_configure("even", background="#FFFFFF")
        self.tree.tag_configure("odd", background="#F8FAFC")

        scrollbar = ttk.Scrollbar(tree_card, orient=tk.VERTICAL, command=self.tree.yview)
        self.tree.configure(yscrollcommand=scrollbar.set)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.tree.pack(fill=tk.BOTH, expand=True)
        self.tree.bind("<<TreeviewSelect>>", self.on_phone_select)

        # --- PRAWY PANEL: SZCZEGÓŁY + HISTORIA ---
        right_frame = ttk.Frame(main_paned)
        main_paned.add(right_frame, weight=2)

        # -------------------------------------------------------------
        # Sekcja A: Formularz edycji danych telefonu
        # -------------------------------------------------------------
        details_box = ttk.LabelFrame(right_frame, text="  Szczegóły telefonu  ", style="Card.TLabelframe", padding=12)
        details_box.pack(fill=tk.X, pady=(0, 8))

        form_grid = ttk.Frame(details_box, style="Card.TFrame")
        form_grid.pack(fill=tk.X)

        self.entries = {}

        # Kolumna lewa: Dane sprzętowe
        left_sub = ttk.LabelFrame(form_grid, text=" Identyfikacja urządzenia ", style="Card.TLabelframe", padding=8)
        left_sub.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=(0, 6))

        dev_fields = [
            ("Model tel:", "model"),
            ("Nr telefonu:", "nr_tel"),
            ("Nr IMEI:", "imei"),
            ("Nr SIM:", "nr_sim"),
            ("Nr seryjny:", "nr_seryjny"),
        ]
        for r, (lbl, key) in enumerate(dev_fields):
            ttk.Label(left_sub, text=lbl, style="Card.TLabel").grid(row=r, column=0, sticky=tk.W, pady=3, padx=(0, 6))
            ent = ttk.Entry(left_sub)
            if key == "nr_tel":
                ent.bind("<FocusOut>", self._format_phone_entry)
            ent.grid(row=r, column=1, sticky=tk.EW, pady=3)
            self.entries[key] = ent
        left_sub.columnconfigure(1, weight=1)

        # Kolumna prawa: Przypisanie i eksploatacja
        right_sub = ttk.LabelFrame(form_grid, text=" Przypisanie i eksploatacja ", style="Card.TLabelframe", padding=8)
        right_sub.pack(side=tk.RIGHT, fill=tk.BOTH, expand=True, padx=(6, 0))

        ttk.Label(right_sub, text="Osoba użytkująca:", style="Card.TLabel").grid(row=0, column=0, sticky=tk.W, pady=3, padx=(0, 6))
        self.entries["osoba_uzytkujaca"] = ttk.Entry(right_sub)
        self.entries["osoba_uzytkujaca"].grid(row=0, column=1, sticky=tk.EW, pady=3)

        ttk.Label(right_sub, text="Osoba odpowiedzialna:", style="Card.TLabel").grid(row=1, column=0, sticky=tk.W, pady=3, padx=(0, 6))
        self.entries["osoba_odpowiedzialna"] = ttk.Entry(right_sub)
        self.entries["osoba_odpowiedzialna"].grid(row=1, column=1, sticky=tk.EW, pady=3)

        ttk.Label(right_sub, text="Rodzaj użytkowania:", style="Card.TLabel").grid(row=2, column=0, sticky=tk.W, pady=3, padx=(0, 6))
        self.entries["rodzaj"] = ttk.Combobox(
            right_sub,
            values=["Montage/Service", "Biurowy", "Zarząd", "Kierownik", "Inny"],
        )
        self.entries["rodzaj"].grid(row=2, column=1, sticky=tk.EW, pady=3)

        ttk.Label(right_sub, text="Stan baterii:", style="Card.TLabel").grid(row=3, column=0, sticky=tk.W, pady=3, padx=(0, 6))
        self.entries["stan_baterii"] = ttk.Combobox(
            right_sub,
            values=["Dobry", "Zadowalający", "Słaby", "Brak informacji"],
            state="readonly",
        )
        self.entries["stan_baterii"].set("Brak informacji")
        self.entries["stan_baterii"].grid(row=3, column=1, sticky=tk.EW, pady=3)

        # Protokół checkbutton
        ttk.Label(right_sub, text="Wysłany protokół:", style="Card.TLabel").grid(row=4, column=0, sticky=tk.W, pady=3, padx=(0, 6))
        self.chk_protocol = ttk.Checkbutton(right_sub, text="Podpisany i dostarczony", variable=self.var_protocol, style="Card.TCheckbutton")
        self.chk_protocol.grid(row=4, column=1, sticky=tk.W, pady=3)

        right_sub.columnconfigure(1, weight=1)

        # Sekcja wyposażenia (akcesoria)
        equip_card = ttk.Frame(details_box, style="Card.TFrame")
        equip_card.pack(fill=tk.X, pady=(8, 0))

        ttk.Label(equip_card, text="Wyposażenie:", font=("Segoe UI", 9, "bold"), style="Card.TLabel").pack(side=tk.LEFT, padx=(4, 12))
        for item in self.wyposazenie_options:
            ttk.Checkbutton(
                equip_card,
                text=item,
                variable=self.wyposazenie_vars[item],
                style="Card.TCheckbutton",
            ).pack(side=tk.LEFT, padx=8)

        # Dolny pasek akcji formularza
        btn_bar = ttk.Frame(details_box, style="Card.TFrame")
        btn_bar.pack(fill=tk.X, pady=(12, 0))

        self.lbl_editing_status = ttk.Label(btn_bar, text="Nowy telefon (niewprowadzony)", style="StatusInfo.TLabel")
        self.lbl_editing_status.pack(side=tk.LEFT, padx=4)

        # Przyciski akcji z prawej
        self.btn_save = ttk.Button(btn_bar, text="💾 Zapisz zmiany", style="Success.TButton", command=self.save_phone)
        self.btn_save.pack(side=tk.RIGHT, padx=4)

        self.btn_delete = ttk.Button(btn_bar, text="🗑️ Do kosza", style="Danger.TButton", command=self.soft_delete_phone)
        self.btn_delete.pack(side=tk.RIGHT, padx=4)

        self.btn_clear = ttk.Button(btn_bar, text="Wyczyść", style="Secondary.TButton", command=self.prepare_new_phone)
        self.btn_clear.pack(side=tk.RIGHT, padx=4)

        # Przyciski trybu kosza (początkowo ukryte)
        self.btn_restore = ttk.Button(btn_bar, text="↩️ Przywróć telefon", style="Success.TButton", command=self.restore_phone)
        self.btn_hard_delete = ttk.Button(btn_bar, text="❌ Usuń trwale z bazy", style="Danger.TButton", command=self.hard_delete_phone)

        # -------------------------------------------------------------
        # Sekcja B: Dziennik zdarzeń i uwagi
        # -------------------------------------------------------------
        history_box = ttk.LabelFrame(right_frame, text="  Dziennik zdarzeń i uwagi  ", style="Card.TLabelframe", padding=10)
        history_box.pack(fill=tk.BOTH, expand=True)

        self.h_top = ttk.Frame(history_box, style="Card.TFrame")
        self.h_top.pack(fill=tk.X, pady=(0, 6))

        ttk.Label(self.h_top, text="Historia operacji i notatki do urządzenia", style="Card.TLabel", font=("Segoe UI", 9, "italic")).pack(side=tk.LEFT)

        self.btn_add_event = ttk.Button(self.h_top, text="+ Dodaj wpis", style="Primary.TButton", command=self.open_add_event_popup)
        self.btn_add_event.pack(side=tk.RIGHT, padx=(4, 0))

        self.btn_add_note = ttk.Button(self.h_top, text="✏️ Edytuj uwagę", style="Secondary.TButton", command=self.open_add_note_popup)

        # Tabela historii
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
        self.history_tree.column("kategoria", width=125, stretch=False)
        self.history_tree.column("opis", width=250)
        self.history_tree.column("uwagi", width=220)

        self.history_tree.tag_configure("even", background="#FFFFFF")
        self.history_tree.tag_configure("odd", background="#F8FAFC")

        self.history_tree.bind("<<TreeviewSelect>>", self.on_history_select)
        self.history_tree.bind("<Double-1>", lambda event: self.open_add_note_popup())

        h_scroll = ttk.Scrollbar(history_box, orient=tk.VERTICAL, command=self.history_tree.yview)
        self.history_tree.configure(yscrollcommand=h_scroll.set)
        h_scroll.pack(side=tk.RIGHT, fill=tk.Y)
        self.history_tree.pack(fill=tk.BOTH, expand=True)

    def show_db_tools_menu(self):
        try:
            x = self.btn_db_tools.winfo_rootx()
            y = self.btn_db_tools.winfo_rooty() + self.btn_db_tools.winfo_height()
            self.db_menu.post(x, y)
        except Exception:
            pass

    def clear_search(self):
        self.search_var.set("")
        self.search_entry.focus_set()

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
            on_save_callback=lambda: self._after_note_saved(history_id),
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

        # Odświeżenie kolorów zebry po sortowaniu
        for index, item_id in enumerate(self.tree.get_children("")):
            tag = "even" if index % 2 == 0 else "odd"
            self.tree.item(item_id, tags=(tag,))

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

        for index, item_id in enumerate(self.history_tree.get_children("")):
            tag = "even" if index % 2 == 0 else "odd"
            self.history_tree.item(item_id, tags=(tag,))

    def load_phone_list(self):
        query = self.search_var.get().strip()
        is_trash = self.show_deleted_var.get()
        rows = db.search_phones(query, show_deleted=is_trash)

        self.tree.delete(*self.tree.get_children())
        for idx, row in enumerate(rows):
            tag = "even" if idx % 2 == 0 else "odd"
            self.tree.insert(
                "",
                tk.END,
                iid=str(row[0]),
                values=(row[1], row[2] or "[BRAK]", row[3]),
                tags=(tag,),
            )

        total = db.get_phones_count(show_deleted=is_trash)
        prefix = "W koszu:" if is_trash else "Aktywnych:"
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
                "nr_seryjny", "rodzaj", "osoba_uzytkujaca", "osoba_odpowiedzialna",
            ]
            for i, key in enumerate(keys):
                self.entries[key].delete(0, tk.END)
                if row[i]:
                    self.entries[key].insert(0, row[i])

            self.var_protocol.set(bool(row[8]))

            raw_equip = row[9] or ""
            current_equip = [x.strip() for x in raw_equip.split(",") if x.strip()]
            for item, var in self.wyposazenie_vars.items():
                var.set(item in current_equip)

            self.entries["stan_baterii"].set(row[10] if row[10] else "Brak informacji")
            self.lbl_editing_status.config(text=f"Edytujesz ID #{self.selected_phone_id}: {row[0]} ({row[1]})")

        history_rows = db.get_phone_history(self.selected_phone_id)
        self.history_tree.delete(*self.history_tree.get_children())
        self.btn_add_note.pack_forget()

        for idx, (h_id, h_data, h_kat, h_opis, h_uwagi) in enumerate(history_rows):
            tag = "even" if idx % 2 == 0 else "odd"
            self.history_tree.insert(
                "",
                tk.END,
                iid=str(h_id),
                values=(h_data, h_kat, h_opis, h_uwagi),
                tags=(tag,),
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
        for key, entry in self.entries.items():
            if key == "stan_baterii":
                entry.set("Brak informacji")
            elif key == "rodzaj":
                entry.set("")
            else:
                entry.delete(0, tk.END)

        self.var_protocol.set(False)
        for var in self.wyposazenie_vars.values():
            var.set(False)

        self.lbl_editing_status.config(text="Nowy telefon (niewprowadzony)")
        self.history_tree.delete(*self.history_tree.get_children())
        self.btn_add_note.pack_forget()
        self.entries["model"].focus()

    def save_phone(self):
        self._format_phone_entry()
        data = {k: ent.get().strip() for k, ent in self.entries.items()}

        data["czy_protokol"] = 1 if self.var_protocol.get() else 0
        selected_equip = [item for item, var in self.wyposazenie_vars.items() if var.get()]
        data["wyposazenie"] = ", ".join(selected_equip)

        if not data["model"] or not data["nr_tel"]:
            messagebox.showwarning("Wymagane pola", "Pola 'Model tel' oraz 'Nr telefonu' są obowiązkowe.")
            return

        if self.selected_phone_id is None:
            self.selected_phone_id = db.insert_phone(data)
            messagebox.showinfo("Sukces", "Nowy telefon został pomyślnie zarejestrowany.")
        else:
            db.update_phone(self.selected_phone_id, data)
            messagebox.showinfo("Sukces", "Dane telefonu zostały zaktualizowane.")

        self.load_phone_list()
        self.tree.selection_set(str(self.selected_phone_id))
        self.refresh_selected_details()

    def soft_delete_phone(self):
        if not self.selected_phone_id:
            messagebox.showwarning("Wybierz telefon", "Nie wybrano telefonu do przeniesienia do kosza.")
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
            self.btn_clear.pack_forget()
            self.btn_restore.pack(side=tk.RIGHT, padx=4)
            self.btn_hard_delete.pack(side=tk.RIGHT, padx=4)
        else:
            self.btn_restore.pack_forget()
            self.btn_hard_delete.pack_forget()
            self.btn_save.pack(side=tk.RIGHT, padx=4)
            self.btn_delete.pack(side=tk.RIGHT, padx=4)
            self.btn_clear.pack(side=tk.RIGHT, padx=4)

        self.load_phone_list()

    def open_employer_dialog(self):
        EmployerDialog(self)

    def open_protocol_dialog(self):
        ProtocolDialog(self, initial_phone_id=self.selected_phone_id)

    def create_db_backup(self):
        default_filename = f"backup_telefony_{datetime.now().strftime('%Y%m%d_%H%M')}.db"

        file_path = filedialog.asksaveasfilename(
            parent=self,
            title="Wybierz miejsce zapisu kopii zapasowej",
            defaultextension=".db",
            initialfile=default_filename,
            filetypes=[("Baza danych SQLite (*.db)", "*.db"), ("Wszystkie pliki", "*.*")],
        )
        if not file_path:
            return

        try:
            db.backup_database(file_path)
            messagebox.showinfo(
                "Kopia zapasowa",
                f"Kopia zapasowa bazy danych została pomyślnie utworzona:\n{file_path}",
                parent=self,
            )
        except PermissionError:
            messagebox.showerror(
                "Błąd zapisu",
                "Brak uprawnień do zapisu we wskazanym folderze lub plik jest zablokowany.",
                parent=self,
            )
        except Exception as e:
            messagebox.showerror(
                "Błąd",
                f"Nie udało się utworzyć kopii zapasowej:\n{e}",
                parent=self,
            )

    def restore_db_backup(self):
        msg = (
            "UWAGA: Ta operacja ZASTĄPI wszystkie obecne dane w programie danymi z wybranego pliku kopii!\n\n"
            "Czy na pewno chcesz kontynuować i wybrać plik kopii zapasowej?"
        )
        if not messagebox.askyesno("Ostrzeżenie", msg, icon=messagebox.WARNING, parent=self):
            return

        file_path = filedialog.askopenfilename(
            parent=self,
            title="Wybierz plik kopii zapasowej bazy danych",
            filetypes=[("Baza danych SQLite (*.db)", "*.db"), ("Wszystkie pliki", "*.*")],
        )
        if not file_path:
            return

        try:
            db.restore_database(file_path)
            self.prepare_new_phone()
            self.load_phone_list()

            messagebox.showinfo(
                "Sukces",
                "Kopia zapasowa została pomyślnie wgrana do bazy danych aplikacji!\nLista urządzeń została zaktualizowana.",
                parent=self,
            )
        except sqlite3.DatabaseError:
            messagebox.showerror(
                "Błąd pliku",
                "Wybrany plik jest uszkodzony lub nie jest poprawną bazą danych SQLite.",
                parent=self,
            )
        except Exception as e:
            messagebox.showerror(
                "Błąd",
                f"Wystąpił nieoczekiwany błąd podczas przywracania bazy:\n{e}",
                parent=self,
            )