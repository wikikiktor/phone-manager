from __future__ import annotations

import tkinter as tk
from collections.abc import Callable, Iterable
from dataclasses import dataclass
from datetime import datetime
from tkinter import filedialog, messagebox, ttk
from typing import Any, Optional

import db
from dialogs import (
    AddEventDialog,
    AddNoteDialog,
    EmployerDialog,
    ExcelImportDialog,
    ProtocolDialog,
)
from login_dialogs import (
    AddUserDialog,
    ChangePasswordDialog,
    RemoveUserDialog,
    Session,
    authenticate,
    regenerate_recovery_code,
)

# ---------------------------------------------------------------------------
# Konfiguracja wyglądu
# ---------------------------------------------------------------------------

APP_TITLE = "Baza Telefonów Służbowych"
WINDOW_SIZE = "1180x720"
WINDOW_MIN_SIZE = (1020, 640)

FONT_FAMILY = "Segoe UI"
FONT = (FONT_FAMILY, 9)
FONT_BOLD = (FONT_FAMILY, 9, "bold")


class Palette:
    BG_APP = "#F1F5F9"
    BG_CARD = "#FFFFFF"
    BG_ROW_ALT = "#F8FAFC"
    BORDER = "#CBD5E1"
    TEXT_MAIN = "#0F172A"
    TEXT_MUTED = "#64748B"
    TEXT_HEADING = "#1E293B"
    BRAND = "#1E3A8A"
    HEADING_BG = "#E2E8F0"
    BADGE_BG = "#E0E7FF"
    BADGE_FG = "#3730A3"
    SELECTION_BG = "#DBEAFE"
    SELECTION_FG = "#1E3A8A"
    WARNING_ROW_BG = "#FEE2E2"
    WARNING_ROW_FG = "#991B1B"
    WARNING_SELECTION_BG = "#FCA5A5"
    WARNING_SELECTION_FG = "#7F1D1D"
    DISABLED_BG = "#94A3B8"
    DISABLED_FG = "#F1F5F9"


BUTTON_STYLES: dict[str, tuple[str, str, tuple[int, int]]] = {
    "Primary.TButton": ("#1E40AF", "#1D4ED8", (10, 6)),
    "Success.TButton": ("#059669", "#047857", (10, 6)),
    "Danger.TButton": ("#EF4444", "#DC2626", (9, 5)),
    "Warning.TButton": ("#FF7300", "#FF9743", (9, 5)),
}

# ---------------------------------------------------------------------------
# Dane formularza
# ---------------------------------------------------------------------------

USAGE_TYPES = ("Montage/Service", "Biurowy", "Zarząd", "Kierownik", "Inny")
BATTERY_STATES = ("Dobry", "Zadowalający", "Słaby", db.DEFAULT_BATTERY_STATE)
EQUIPMENT_OPTIONS = ("Ładowarka", "Etui", "Kabel USB", "Szkło ochronne")
EQUIPMENT_SEPARATOR = ", "

DEVICE_FIELDS: tuple[tuple[str, str], ...] = (
    ("model", "Model tel:"),
    ("nr_tel", "Nr telefonu:"),
    ("imei", "Nr IMEI:"),
    ("nr_sim", "Nr SIM:"),
    ("nr_seryjny", "Nr seryjny:"),
)
PERSON_FIELDS: tuple[tuple[str, str], ...] = (
    ("osoba_uzytkujaca", "Osoba użytkująca:"),
    ("osoba_odpowiedzialna", "Osoba odpowiedzialna:"),
)

NEW_PHONE_STATUS = "Nowy telefon (niewprowadzony)"
NO_USER = "[BRAK]"

# ---------------------------------------------------------------------------
# Tabele
# ---------------------------------------------------------------------------

@dataclass(frozen=True)
class Column:
    key: str
    title: str
    width: int
    min_width: int = 20
    stretch: bool = True


PHONE_COLUMNS: tuple[Column, ...] = (
    Column("nr_tel", "Nr telefonu", 115, 90, stretch=False),
    Column("uzytkownik", "Użytkownik", 135, 100),
    Column("model", "Model", 125, 90),
)
HISTORY_COLUMNS: tuple[Column, ...] = (
    Column("data", "Data", 120, stretch=False),
    Column("kategoria", "Kategoria", 125, stretch=False),
    Column("opis", "Opis zdarzenia", 250),
    Column("uwagi", "Uwagi", 220),
)

TAG_EVEN = "even"
TAG_ODD = "odd"
TAG_WARNING = "warning"
SORT_ASC_ARROW = "  ▲"
SORT_DESC_ARROW = "  ▼"

@dataclass
class SortState:
    column: Optional[str] = None
    reverse: bool = False

    def toggle(self, column: str) -> None:
        if self.column == column:
            self.reverse = not self.reverse
        else:
            self.column = column
            self.reverse = False


def _timestamp() -> str:
    return datetime.now().strftime("%Y%m%d_%H%M")  


def _create_tree(
    parent: tk.Misc, columns: Iterable[Column], on_heading_click: Callable[[str], None], **options: Any
) -> ttk.Treeview:
    columns = tuple(columns)
    tree = ttk.Treeview(
        parent, columns=[c.key for c in columns], show="headings", selectmode="browse", **options
    )
    for column in columns:
        tree.heading(column.key, text=column.title, command=lambda k=column.key: on_heading_click(k))
        tree.column(column.key, width=column.width, minwidth=column.min_width, stretch=column.stretch)

    tree.tag_configure(TAG_EVEN, background=Palette.BG_CARD)
    tree.tag_configure(TAG_ODD, background=Palette.BG_ROW_ALT)

    scrollbar = ttk.Scrollbar(parent, orient=tk.VERTICAL, command=tree.yview)
    tree.configure(yscrollcommand=scrollbar.set)
    scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
    tree.pack(fill=tk.BOTH, expand=True)
    return tree


def _apply_sort(tree: ttk.Treeview, columns: Iterable[Column], state: SortState) -> None:
    for column in columns:
        arrow = ""
        if column.key == state.column:
            arrow = SORT_DESC_ARROW if state.reverse else SORT_ASC_ARROW
        tree.heading(column.key, text=column.title + arrow)

    if state.column is None:
        return

    items = sorted(
        tree.get_children(),
        key=lambda item: tree.set(item, state.column).lower(),
        reverse=state.reverse,
    )
    for index, item in enumerate(items):
        tree.move(item, "", index)


def _apply_row_tags(tree: ttk.Treeview, warning_items: Iterable[str] = ()) -> None:
    warning_items = set(warning_items)
    for index, item in enumerate(tree.get_children()):
        if item in warning_items:
            tag = TAG_WARNING
        else:
            tag = TAG_EVEN if index % 2 == 0 else TAG_ODD
        tree.item(item, tags=(tag,))


# ---------------------------------------------------------------------------
# Okno główne
# ---------------------------------------------------------------------------

class PhoneDatabaseApp(tk.Tk):
    def __init__(self) -> None:
        super().__init__()
        self.withdraw()
        self.title(APP_TITLE)
        self.geometry(WINDOW_SIZE)
        self.minsize(*WINDOW_MIN_SIZE)
        self._setup_styles()

        session = authenticate(self)
        if session is None:
            self.destroy()
            return
        self.session: Session = session

        self.selected_phone_id: Optional[int] = None
        self.show_deleted_var = tk.BooleanVar(value=False)
        self.search_var = tk.StringVar()
        self.var_protocol = tk.BooleanVar(value=False)
        self.equipment_vars = {option: tk.BooleanVar(value=False) for option in EQUIPMENT_OPTIONS}
        self.entries: dict[str, ttk.Entry] = {}

        self._phone_sort = SortState()
        self._history_sort = SortState()
        self._warning_items: set[str] = set()

        db.init_db()
        self._build_ui()
        self.search_var.trace_add("write", lambda *_args: self.load_phone_list())
        self.load_phone_list()
        self.deiconify()

    # =======================================================================
    # Style
    # =======================================================================

    def _setup_styles(self) -> None:
        self.style = ttk.Style(self)
        try:
            self.style.theme_use("clam")
        except tk.TclError:
            pass

        self.configure(bg=Palette.BG_APP)
        configure = self.style.configure

        configure(".", background=Palette.BG_APP, foreground=Palette.TEXT_MAIN, font=FONT)
        configure("TFrame", background=Palette.BG_APP)
        configure("Card.TFrame", background=Palette.BG_CARD)
        configure("TLabel", background=Palette.BG_APP, foreground=Palette.TEXT_MAIN, font=FONT)
        configure("Card.TLabel", background=Palette.BG_CARD, foreground=Palette.TEXT_MAIN, font=FONT)
        configure("CardBold.TLabel", background=Palette.BG_CARD, font=FONT_BOLD)
        configure("CardItalic.TLabel", background=Palette.BG_CARD, font=(FONT_FAMILY, 9, "italic"))
        configure("SearchIcon.TLabel", font=(FONT_FAMILY, 10))

        configure(
            "BrandTitle.TLabel",
            background=Palette.BG_APP,
            foreground=Palette.BRAND,
            font=(FONT_FAMILY, 13, "bold"),
        )
        configure(
            "BrandSub.TLabel",
            background=Palette.BG_APP,
            foreground=Palette.TEXT_MUTED,
            font=(FONT_FAMILY, 8),
        )
        configure(
            "CardTitle.TLabel",
            background=Palette.BG_CARD,
            foreground=Palette.TEXT_HEADING,
            font=(FONT_FAMILY, 10, "bold"),
        )
        configure(
            "StatusInfo.TLabel",
            background=Palette.BG_CARD,
            foreground=Palette.TEXT_MUTED,
            font=(FONT_FAMILY, 8, "italic"),
        )
        configure(
            "Badge.TLabel",
            background=Palette.BADGE_BG,
            foreground=Palette.BADGE_FG,
            font=FONT_BOLD,
            padding=(10, 4),
        )

        configure(
            "Card.TLabelframe",
            background=Palette.BG_CARD,
            bordercolor=Palette.BORDER,
            relief="solid",
            borderwidth=1,
        )
        configure(
            "Card.TLabelframe.Label",
            background=Palette.BG_CARD,
            foreground=Palette.BRAND,
            font=(FONT_FAMILY, 10, "bold"),
        )

        configure("TEntry", fieldbackground=Palette.BG_CARD, padding=5)
        configure("TCombobox", fieldbackground=Palette.BG_CARD, padding=4)
        self.style.map("TCombobox", fieldbackground=[("readonly", Palette.BG_CARD)])

        configure("TCheckbutton", background=Palette.BG_APP, font=FONT)
        configure("Card.TCheckbutton", background=Palette.BG_CARD, font=FONT)

        for style_name, (background, active, padding) in BUTTON_STYLES.items():
            configure(
                style_name,
                background=background,
                foreground="#FFFFFF",
                font=FONT_BOLD,
                borderwidth=0,
                padding=padding,
            )
            self.style.map(
                style_name,
                background=[("active", active), ("disabled", Palette.DISABLED_BG)],
                foreground=[("disabled", Palette.DISABLED_FG)],
            )

        configure(
            "Secondary.TButton",
            background=Palette.BG_CARD,
            foreground=Palette.TEXT_HEADING,
            font=FONT,
            borderwidth=1,
            bordercolor=Palette.BORDER,
            padding=(8, 5),
        )
        self.style.map("Secondary.TButton", background=[("active", Palette.BG_ROW_ALT)])

        configure(
            "Treeview",
            background=Palette.BG_CARD,
            fieldbackground=Palette.BG_CARD,
            foreground=Palette.TEXT_MAIN,
            rowheight=26,
            font=FONT,
            bordercolor=Palette.BORDER,
            borderwidth=1,
        )
        configure(
            "Treeview.Heading",
            background=Palette.HEADING_BG,
            foreground=Palette.TEXT_MAIN,
            font=FONT_BOLD,
            padding=(6, 6),
            relief="flat",
        )
        self.style.map("Treeview.Heading", background=[("active", Palette.BORDER)])
        self.style.map(
            "Treeview",
            background=[("selected", Palette.SELECTION_BG)],
            foreground=[("selected", Palette.SELECTION_FG)],
        )
        self._set_phone_selection_colors(warning=False)

    def _set_phone_selection_colors(self, *, warning: bool) -> None:
        if warning:
            colors = (Palette.WARNING_SELECTION_BG, Palette.WARNING_SELECTION_FG)
        else:
            colors = (Palette.SELECTION_BG, Palette.SELECTION_FG)
        self.style.map(
            "Phones.Treeview",
            background=[("selected", colors[0])],
            foreground=[("selected", colors[1])],
        )

    # =======================================================================
    # Budowa interfejsu
    # =======================================================================

    def _build_ui(self) -> None:
        self._build_header()

        main_paned = ttk.PanedWindow(self, orient=tk.HORIZONTAL)
        main_paned.pack(fill=tk.BOTH, expand=True, padx=14, pady=(4, 12))

        left_frame = ttk.Frame(main_paned, width=390)
        main_paned.add(left_frame, weight=1)
        self._build_phone_list(left_frame)

        right_frame = ttk.Frame(main_paned)
        main_paned.add(right_frame, weight=2)
        self._build_details(right_frame)
        self._build_history(right_frame)

    def _build_header(self) -> None:
        header = ttk.Frame(self, padding=(14, 10, 14, 6))
        header.pack(fill=tk.X)

        top_row = ttk.Frame(header)
        top_row.pack(fill=tk.X, pady=(0, 8))

        brand = ttk.Frame(top_row)
        brand.pack(side=tk.LEFT)
        ttk.Label(brand, text="📱 Baza Telefonów", style="BrandTitle.TLabel").pack(anchor=tk.W)
        ttk.Label(
            brand, text="System ewidencji sprzętu i protokołów", style="BrandSub.TLabel"
        ).pack(anchor=tk.W)

        actions = ttk.Frame(top_row)
        actions.pack(side=tk.RIGHT)
        for text, style, command in (
            ("+ Nowy telefon", "Primary.TButton", self.prepare_new_phone),
            ("📋 Protokół", "Secondary.TButton", self.open_protocol_dialog),
            ("📊 Eksport Excel", "Secondary.TButton", self.export_to_excel),
            ("📥 Import Excel", "Secondary.TButton", self.open_excel_import),
        ):
            ttk.Button(actions, text=text, style=style, command=command).pack(side=tk.LEFT, padx=3)

        self.db_menu = tk.Menu(self, tearoff=0)
        self.db_menu.add_command(label="🏢 Dane firmy / pracodawcy", command=self.open_employer_dialog)
        self.db_menu.add_separator()
        self.db_menu.add_command(
            label="💾 Utwórz kopię zapasową bazy (.db)", command=self.create_db_backup
        )
        self.db_menu.add_command(label="🔄 Przywróć bazę z pliku (.db)", command=self.restore_db_backup)
        self.db_menu.add_separator()
        self.db_menu.add_command(label="🔑 Zmień hasło", command=self.open_change_password)
        self._admin_menu_items: list[int] = []
        for label, command in (
            ("👥 Dodaj użytkownika (administrator)", self.open_add_user),
            ("🚫 Usuń użytkownika (administrator)", self.open_remove_user),
            ("🆘 Wygeneruj nowy kod awaryjny (administrator)", 
            self.regenerate_recovery_code),
        ):
            self.db_menu.add_command(label=label, command=command)
            self._admin_menu_items.append(self.db_menu.index(tk.END))
        self.db_menu.add_separator()
        self.db_menu.add_command(label="🚪 Wyloguj", command=self.logout)

        self.btn_db_tools = ttk.Button(
            actions, text="⚙️ Opcje bazy ▾", style="Secondary.TButton", command=self.show_db_tools_menu
        )
        self.btn_db_tools.pack(side=tk.LEFT, padx=3)

        search_row = ttk.Frame(header)
        search_row.pack(fill=tk.X)

        search_box = ttk.Frame(search_row)
        search_box.pack(side=tk.LEFT, fill=tk.X)

        ttk.Label(search_box, text="🔍", style="SearchIcon.TLabel").pack(side=tk.LEFT, padx=(0, 4))
        self.search_entry = ttk.Entry(search_box, textvariable=self.search_var, width=36)
        self.search_entry.pack(side=tk.LEFT, padx=(0, 4))
        ttk.Button(
            search_box, text="✕", width=2, style="Secondary.TButton", command=self.clear_search
        ).pack(side=tk.LEFT, padx=(0, 10))

        ttk.Checkbutton(
            search_box,
            text="Pokaż kosz (usunięte)",
            variable=self.show_deleted_var,
            command=self.on_toggle_trash_view,
        ).pack(side=tk.LEFT)

        self.lbl_count = ttk.Label(search_row, style="Badge.TLabel")
        self.lbl_count.pack(side=tk.RIGHT)
        self.lbl_user = ttk.Label(search_row, style="BrandSub.TLabel")
        self.lbl_user.pack(side=tk.RIGHT, padx=(0, 6))
        self._apply_session()

    def _apply_session(self) -> None:
        suffix = " (administrator)" if self.session.is_admin else ""
        self.lbl_user.configure(text=f"👤 {self.session.username}{suffix}")
        state = tk.NORMAL if self.session.is_admin else tk.DISABLED
        for index in self._admin_menu_items:
            self.db_menu.entryconfigure(index, state=state)

    def _require_admin(self) -> bool:
        if self.session.is_admin:
            return True
        messagebox.showwarning(
            "Brak uprawnień", "Tę operację może wykonać tylko administrator.", parent=self
        )
        return False

    def _build_phone_list(self, parent: ttk.Frame) -> None:
        card = ttk.LabelFrame(parent, text="  Lista urządzeń  ", style="Card.TLabelframe", padding=6)
        card.pack(fill=tk.BOTH, expand=True)

        self.tree = _create_tree(
            card, PHONE_COLUMNS, self.on_phone_column_click, style="Phones.Treeview"
        )
        self.tree.tag_configure(
            TAG_WARNING, background=Palette.WARNING_ROW_BG, foreground=Palette.WARNING_ROW_FG
        )
        self.tree.bind("<<TreeviewSelect>>", self.on_phone_select)

    def _build_details(self, parent: ttk.Frame) -> None:
        details = ttk.LabelFrame(
            parent, text="  Szczegóły telefonu  ", style="Card.TLabelframe", padding=12
        )
        details.pack(fill=tk.X, pady=(0, 8))

        form = ttk.Frame(details, style="Card.TFrame")
        form.pack(fill=tk.X)
        self._build_device_fields(form)
        self._build_assignment_fields(form)
        self._build_equipment_row(details)
        self._build_form_buttons(details)

    @staticmethod
    def _add_form_label(parent: tk.Misc, row: int, text: str) -> None:
        ttk.Label(parent, text=text, style="Card.TLabel").grid(
            row=row, column=0, sticky=tk.W, pady=3, padx=(0, 6)
        )

    def _build_device_fields(self, parent: ttk.Frame) -> None:
        box = ttk.LabelFrame(
            parent, text=" Identyfikacja urządzenia ", style="Card.TLabelframe", padding=8
        )
        box.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=(0, 6))
        box.columnconfigure(1, weight=1)

        for row, (key, label) in enumerate(DEVICE_FIELDS):
            self._add_form_label(box, row, label)
            entry = ttk.Entry(box)
            entry.grid(row=row, column=1, sticky=tk.EW, pady=3)
            self.entries[key] = entry

        self.entries["nr_tel"].bind("<FocusOut>", self._format_phone_entry)

    def _build_assignment_fields(self, parent: ttk.Frame) -> None:
        box = ttk.LabelFrame(
            parent, text=" Przypisanie i eksploatacja ", style="Card.TLabelframe", padding=8
        )
        box.pack(side=tk.RIGHT, fill=tk.BOTH, expand=True, padx=(6, 0))
        box.columnconfigure(1, weight=1)

        for row, (key, label) in enumerate(PERSON_FIELDS):
            self._add_form_label(box, row, label)
            entry = ttk.Entry(box)
            entry.grid(row=row, column=1, sticky=tk.EW, pady=3)
            self.entries[key] = entry

        row = len(PERSON_FIELDS)
        self._add_form_label(box, row, "Rodzaj użytkowania:")
        self.entries["rodzaj"] = ttk.Combobox(box, values=USAGE_TYPES)
        self.entries["rodzaj"].grid(row=row, column=1, sticky=tk.EW, pady=3)

        row += 1
        self._add_form_label(box, row, "Stan baterii:")
        self.entries["stan_baterii"] = ttk.Combobox(box, values=BATTERY_STATES, state="readonly")
        self.entries["stan_baterii"].set(db.DEFAULT_BATTERY_STATE)
        self.entries["stan_baterii"].grid(row=row, column=1, sticky=tk.EW, pady=3)

        row += 1
        self._add_form_label(box, row, "Wysłany protokół:")
        ttk.Checkbutton(
            box,
            text="Podpisany i dostarczony",
            variable=self.var_protocol,
            style="Card.TCheckbutton",
        ).grid(row=row, column=1, sticky=tk.W, pady=3)

    def _build_equipment_row(self, parent: ttk.Frame) -> None:
        row = ttk.Frame(parent, style="Card.TFrame")
        row.pack(fill=tk.X, pady=(8, 0))

        ttk.Label(row, text="Wyposażenie:", style="CardBold.TLabel").pack(side=tk.LEFT, padx=(4, 12))
        for option, variable in self.equipment_vars.items():
            ttk.Checkbutton(row, text=option, variable=variable, style="Card.TCheckbutton").pack(
                side=tk.LEFT, padx=8
            )

    def _build_form_buttons(self, parent: ttk.Frame) -> None:
        bar = ttk.Frame(parent, style="Card.TFrame")
        bar.pack(fill=tk.X, pady=(12, 0))

        self.lbl_editing_status = ttk.Label(bar, text=NEW_PHONE_STATUS, style="StatusInfo.TLabel")

        def button(text: str, style: str, command: Callable[[], None]) -> ttk.Button:
            return ttk.Button(bar, text=text, style=style, command=command)

        self.btn_save = button("💾 Zapisz zmiany", "Success.TButton", self.save_phone)
        self.btn_delete = button("🗑️ Do kosza", "Danger.TButton", self.soft_delete_phone)
        self.btn_warning = button("⚠️ Oznacz telefon", "Warning.TButton", self.toggle_warning)
        self.btn_clear = button("Wyczyść", "Secondary.TButton", self.prepare_new_phone)
        self.btn_restore = button("↩️ Przywróć telefon", "Success.TButton", self.restore_phone)
        self.btn_hard_delete = button(
            "❌ Usuń trwale z bazy", "Danger.TButton", self.hard_delete_phone
        )
        self._layout_form_buttons(in_trash=False)

    def _layout_form_buttons(self, *, in_trash: bool) -> None:
        all_buttons = (
            self.btn_save,
            self.btn_delete,
            self.btn_warning,
            self.btn_clear,
            self.btn_restore,
            self.btn_hard_delete,
        )
        for widget in (*all_buttons, self.lbl_editing_status):
            widget.pack_forget()

        if in_trash:
            visible = (self.btn_restore, self.btn_hard_delete, self.btn_warning)
        else:
            visible = (self.btn_save, self.btn_delete, self.btn_warning, self.btn_clear)
        for widget in visible:
            widget.pack(side=tk.RIGHT, padx=4)
        self.lbl_editing_status.pack(side=tk.LEFT, padx=4)

    def _build_history(self, parent: ttk.Frame) -> None:
        box = ttk.LabelFrame(
            parent, text="  Dziennik zdarzeń i uwagi  ", style="Card.TLabelframe", padding=10
        )
        box.pack(fill=tk.BOTH, expand=True)

        toolbar = ttk.Frame(box, style="Card.TFrame")
        toolbar.pack(fill=tk.X, pady=(0, 6))
        ttk.Label(
            toolbar, text="Historia operacji i notatki do urządzenia", style="CardItalic.TLabel"
        ).pack(side=tk.LEFT)

        self.btn_add_event = ttk.Button(
            toolbar, text="+ Dodaj wpis", style="Primary.TButton", command=self.open_add_event_popup
        )
        self.btn_add_event.pack(side=tk.RIGHT, padx=(4, 0))
        self.btn_add_note = ttk.Button(
            toolbar, text="✏️ Edytuj uwagę", style="Secondary.TButton", command=self.open_add_note_popup
        )

        self.history_tree = _create_tree(box, HISTORY_COLUMNS, self.on_history_column_click)
        self.history_tree.bind("<<TreeviewSelect>>", self.on_history_select)
        self.history_tree.bind("<Double-1>", lambda _event: self.open_add_note_popup())

    # =======================================================================
    # Lista telefonów
    # =======================================================================

    def load_phone_list(self) -> None:
        query = self.search_var.get().strip()
        in_trash = self.show_deleted_var.get()
        rows = db.search_phones(query, show_deleted=in_trash)

        self.tree.delete(*self.tree.get_children())
        self._warning_items = set()
        for phone_id, nr_tel, user, model, is_damaged in rows:
            item = str(phone_id)
            if is_damaged:
                self._warning_items.add(item)
            self.tree.insert("", tk.END, iid=item, values=(nr_tel, user or NO_USER, model))

        _apply_sort(self.tree, PHONE_COLUMNS, self._phone_sort)
        _apply_row_tags(self.tree, self._warning_items)
        self._update_counter(len(rows), query, in_trash)

    def _update_counter(self, found: int, query: str, in_trash: bool) -> None:
        total = db.get_phones_count(show_deleted=in_trash)
        prefix = "W koszu:" if in_trash else "Aktywnych:"
        if query:
            text = f"{prefix} znaleziono {found} z {total}"
        else:
            text = f"{prefix} {total} telefonów"
        self.lbl_count.configure(text=text)

    def on_phone_column_click(self, column: str) -> None:
        self._phone_sort.toggle(column)
        _apply_sort(self.tree, PHONE_COLUMNS, self._phone_sort)
        _apply_row_tags(self.tree, self._warning_items)

    def on_phone_select(self, _event: Optional[tk.Event] = None) -> None:
        selection = self.tree.selection()
        if not selection:
            return
        self.selected_phone_id = int(selection[0])
        self.refresh_selected_details()

    def _select_phone_in_list(self, phone_id: int) -> None:
        item = str(phone_id)
        if self.tree.exists(item):
            self.tree.selection_set(item)
            self.tree.see(item)

    def _update_selection_colors(self) -> None:
        selection = self.tree.selection()
        if selection:
            self._set_phone_selection_colors(warning=selection[0] in self._warning_items)

    def clear_search(self) -> None:
        self.search_var.set("")
        self.search_entry.focus_set()

    def on_toggle_trash_view(self) -> None:
        self.prepare_new_phone()
        self._layout_form_buttons(in_trash=self.show_deleted_var.get())
        self.load_phone_list()

    # =======================================================================
    # Formularz szczegółów
    # =======================================================================

    def _set_field(self, key: str, value: Any) -> None:
        widget = self.entries[key]
        text = "" if value is None else str(value)
        if isinstance(widget, ttk.Combobox):
            widget.set(text)
        else:
            widget.delete(0, tk.END)
            widget.insert(0, text)

    def _fill_form(self, phone: dict[str, Any]) -> None:
        for key in self.entries:
            if key != "stan_baterii":
                self._set_field(key, phone.get(key))
        self._set_field("stan_baterii", phone.get("stan_baterii") or db.DEFAULT_BATTERY_STATE)

        self.var_protocol.set(bool(phone.get("czy_protokol")))
        equipment = {
            item.strip()
            for item in (phone.get("wyposazenie") or "").split(",")
            if item.strip()
        }
        for option, variable in self.equipment_vars.items():
            variable.set(option in equipment)

    def _collect_form_data(self) -> dict[str, Any]:
        data: dict[str, Any] = {key: widget.get().strip() for key, widget in self.entries.items()}
        data["czy_protokol"] = int(self.var_protocol.get())
        data["wyposazenie"] = EQUIPMENT_SEPARATOR.join(
            option for option, variable in self.equipment_vars.items() if variable.get()
        )
        return data

    def _update_warning_button(self, is_damaged: bool) -> None:
        text = "✅ Odznacz telefon" if is_damaged else "⚠️ Oznacz telefon"
        self.btn_warning.configure(text=text)

    def refresh_selected_details(self) -> None:
        if self.selected_phone_id is None:
            return

        self._update_selection_colors()

        row = db.get_phone_by_id(self.selected_phone_id)
        if row is not None:
            phone = dict(zip(db.PHONE_FIELDS, row))
            self._fill_form(phone)
            self._update_warning_button(bool(phone["czy_uszkodzony"]))
            self.lbl_editing_status.configure(
                text=f"Edytujesz ID #{self.selected_phone_id}: {phone['model']} ({phone['nr_tel']})"
            )

        self._load_history()

    def prepare_new_phone(self) -> None:
        self.selected_phone_id = None
        self.tree.selection_remove(*self.tree.selection())
        self._set_phone_selection_colors(warning=False)

        for key in self.entries:
            self._set_field(key, db.DEFAULT_BATTERY_STATE if key == "stan_baterii" else "")
        self.var_protocol.set(False)
        for variable in self.equipment_vars.values():
            variable.set(False)

        self._update_warning_button(False)
        self.lbl_editing_status.configure(text=NEW_PHONE_STATUS)
        self.history_tree.delete(*self.history_tree.get_children())
        self.btn_add_note.pack_forget()
        self.entries["model"].focus_set()

    def _format_phone_entry(self, _event: Optional[tk.Event] = None) -> None:
        current = self.entries["nr_tel"].get()
        formatted = db.format_phone_number(current)
        if formatted != current:
            self._set_field("nr_tel", formatted)

    # =======================================================================
    # Akcje na telefonie
    # =======================================================================

    def _require_selected_phone(self, message: str) -> bool:
        if self.selected_phone_id is None:
            messagebox.showwarning("Wybierz telefon", message, parent=self)
            return False
        return True

    def _after_phone_removed_from_view(self) -> None:
        self.prepare_new_phone()
        self.load_phone_list()

    def save_phone(self) -> None:
        self._format_phone_entry()
        data = self._collect_form_data()

        if not data["model"] or not data["nr_tel"]:
            messagebox.showwarning(
                "Wymagane pola", "Pola 'Model tel' oraz 'Nr telefonu' są obowiązkowe.", parent=self
            )
            return

        if self.selected_phone_id is None:
            self.selected_phone_id = db.insert_phone(data)
            messagebox.showinfo("Sukces", "Nowy telefon został pomyślnie zarejestrowany.", parent=self)
        else:
            current = db.get_phone_by_id(self.selected_phone_id)
            if current is not None:
                data["czy_uszkodzony"] = dict(zip(db.PHONE_FIELDS, current))["czy_uszkodzony"]
            db.update_phone(self.selected_phone_id, data)
            messagebox.showinfo("Sukces", "Dane telefonu zostały zaktualizowane.", parent=self)

        self.load_phone_list()
        self._select_phone_in_list(self.selected_phone_id)
        self.refresh_selected_details()

    def soft_delete_phone(self) -> None:
        if not self._require_selected_phone("Nie wybrano telefonu do przeniesienia do kosza."):
            return
        if messagebox.askyesno(
            "Potwierdzenie", "Czy na pewno chcesz przenieść ten telefon do kosza?", parent=self
        ):
            db.soft_delete_phone(self.selected_phone_id)
            messagebox.showinfo("Kosz", "Telefon został przeniesiony do kosza.", parent=self)
            self._after_phone_removed_from_view()

    def restore_phone(self) -> None:
        if not self._require_selected_phone("Wybierz telefon z kosza, który chcesz przywrócić."):
            return
        if messagebox.askyesno(
            "Potwierdzenie", "Czy chcesz przywrócić ten telefon do aktywnych urządzeń?", parent=self
        ):
            db.restore_phone(self.selected_phone_id)
            messagebox.showinfo("Sukces", "Telefon został przywrócony.", parent=self)
            self._after_phone_removed_from_view()

    def hard_delete_phone(self) -> None:
        if not self._require_selected_phone("Wybierz telefon do trwałego usunięcia."):
            return
        message = (
            "UWAGA: Ta operacja jest NIEODWRACALNA!\n"
            "Telefon oraz cała jego historia zdarzeń zostaną trwale wykasowane z bazy.\n\n"
            "Czy na pewno chcesz kontynuować?"
        )
        if messagebox.askyesno("Ostrzeżenie", message, icon=messagebox.WARNING, parent=self):
            db.hard_delete_phone(self.selected_phone_id)
            messagebox.showinfo("Usunięto", "Telefon został trwale wykasowany z bazy.", parent=self)
            self._after_phone_removed_from_view()

    def toggle_warning(self) -> None:
        if not self._require_selected_phone("Wybierz telefon z listy, aby go oznaczyć."):
            return

        row = db.get_phone_by_id(self.selected_phone_id)
        is_damaged = bool(row and dict(zip(db.PHONE_FIELDS, row))["czy_uszkodzony"])
        if is_damaged and not messagebox.askyesno(
            "Potwierdzenie odznaczenia", "Czy na pewno chcesz odznaczyć ten telefon?", parent=self
        ):
            return

        new_state = db.toggle_phone_warning(self.selected_phone_id)
        self._update_warning_button(bool(new_state))

        self.load_phone_list()
        self._select_phone_in_list(self.selected_phone_id)
        self.refresh_selected_details()

    # =======================================================================
    # Historia zdarzeń
    # =======================================================================

    def _load_history(self) -> None:
        self.history_tree.delete(*self.history_tree.get_children())
        self.btn_add_note.pack_forget()

        for history_id, date, category, description, note in db.get_phone_history(
            self.selected_phone_id
        ):
            self.history_tree.insert(
                "", tk.END, iid=str(history_id), values=(date, category, description, note)
            )

        _apply_sort(self.history_tree, HISTORY_COLUMNS, self._history_sort)
        _apply_row_tags(self.history_tree)

    def on_history_column_click(self, column: str) -> None:
        self._history_sort.toggle(column)
        _apply_sort(self.history_tree, HISTORY_COLUMNS, self._history_sort)
        _apply_row_tags(self.history_tree)

    def on_history_select(self, _event: Optional[tk.Event] = None) -> None:
        if self.history_tree.selection():
            self.btn_add_note.pack(side=tk.RIGHT, padx=(0, 6))
        else:
            self.btn_add_note.pack_forget()

    def open_add_event_popup(self) -> None:
        if not self._require_selected_phone("Wybierz lub zapisz telefon przed dodaniem zdarzenia."):
            return
        AddEventDialog(self, self.selected_phone_id, on_save_callback=self.refresh_selected_details)

    def open_add_note_popup(self) -> None:
        selection = self.history_tree.selection()
        if not selection:
            messagebox.showwarning(
                "Wybierz wpis",
                "Wybierz wpis z historii, aby dodać lub edytować uwagę.",
                parent=self,
            )
            return

        item = selection[0]
        values = self.history_tree.item(item, "values")
        current_note = values[3] if len(values) > 3 else ""
        AddNoteDialog(
            parent=self,
            history_id=int(item),
            current_note=current_note,
            on_save_callback=lambda: self._after_note_saved(item),
        )

    def _after_note_saved(self, item: str) -> None:
        self.refresh_selected_details()
        if self.history_tree.exists(item):
            self.history_tree.selection_set(item)
            self.btn_add_note.pack(side=tk.RIGHT, padx=(0, 6))

    # =======================================================================
    # Okna dialogowe i operacje na plikach
    # =======================================================================

    def show_db_tools_menu(self) -> None:
        x = self.btn_db_tools.winfo_rootx()
        y = self.btn_db_tools.winfo_rooty() + self.btn_db_tools.winfo_height()
        try:
            self.db_menu.tk_popup(x, y)
        finally:
            self.db_menu.grab_release()

    def logout(self) -> None:
        self._close_child_windows()
        self.search_var.set("")
        if self.show_deleted_var.get():
            self.show_deleted_var.set(False)
            self._layout_form_buttons(in_trash=False)
        self.prepare_new_phone()
        self.tree.delete(*self.tree.get_children())
        self._warning_items = set()
        self.lbl_count.configure(text="")

        db.lock()
        self.withdraw()

        session = authenticate(self)
        if session is None:
            self.destroy()
            return

        self.session = session
        self.lbl_user.configure(text=f"👤 {self.session.username}")
        self._apply_session()
        self.load_phone_list()
        self.deiconify()
        self.entries["model"].focus_set()  

    def _close_child_windows(self) -> None:
        for widget in self.winfo_children():
            if isinstance(widget, tk.Toplevel):
                widget.destroy()

    def open_employer_dialog(self) -> None:
        EmployerDialog(self)

    def open_change_password(self) -> None:
        ChangePasswordDialog(self, self.session)

    def open_add_user(self) -> None:
        if self._require_admin():
            AddUserDialog(self, self.session)

    def regenerate_recovery_code(self) -> None:
        if self._require_admin():
            regenerate_recovery_code(self, self.session)

    def open_remove_user(self) -> None:
        if not self._require_admin():
            return
        others = [
            (name, role)
            for name, role in self.session.keystore.users_with_roles()
            if name != self.session.username
        ]
        if not others:
            messagebox.showinfo(
                "Brak użytkowników", "Nie ma innych użytkowników do usunięcia.", parent=self
            )
            return
        RemoveUserDialog(self, self.session, others)

    def open_protocol_dialog(self) -> None:
        ProtocolDialog(self, initial_phone_id=self.selected_phone_id)

    def open_excel_import(self) -> None:
        file_path = filedialog.askopenfilename(
            parent=self,
            title="Wybierz plik Excel do importu",
            filetypes=[("Pliki Excel", "*.xlsx *.xlsm"), ("Wszystkie pliki", "*.*")],
        )
        if file_path:
            ExcelImportDialog(self, file_path, on_success_callback=self.load_phone_list)

    def export_to_excel(self) -> None:
        visible_ids = [int(item) for item in self.tree.get_children()]
        if not visible_ids:
            messagebox.showwarning(
                "Brak danych", "Lista jest pusta. Brak wyników do wyeksportowania.", parent=self
            )
            return

        prefix = "kosz" if self.show_deleted_var.get() else "telefony"
        file_path = filedialog.asksaveasfilename(
            parent=self,
            title="Wybierz miejsce zapisu pliku Excel",
            defaultextension=".xlsx",
            initialfile=f"{prefix}_eksport_{_timestamp()}.xlsx",
            filetypes=[("Pliki Excel", "*.xlsx"), ("Wszystkie pliki", "*.*")],
        )
        if not file_path:
            return

        try:
            exported = db.export_to_excel(file_path, phone_ids=visible_ids)
        except PermissionError:
            messagebox.showerror(
                "Błąd zapisu",
                "Nie można zapisać pliku. Upewnij się, że plik nie jest obecnie otwarty "
                "w programie Excel.",
                parent=self,
            )
        except Exception as error:
            messagebox.showerror("Błąd eksportu", f"Wystąpił nieoczekiwany błąd:\n{error}", parent=self)
        else:
            messagebox.showinfo(
                "Eksport zakończony",
                f"Pomyślnie wyeksportowano {exported} urządzeń wraz z historią do pliku:\n{file_path}",
                parent=self,
            )

    def create_db_backup(self) -> None:
        file_path = filedialog.asksaveasfilename(
            parent=self,
            title="Wybierz miejsce zapisu kopii zapasowej",
            defaultextension=".db",
            initialfile=f"backup_telefony_{_timestamp()}.db",
            filetypes=[("Baza danych SQLite (*.db)", "*.db"), ("Wszystkie pliki", "*.*")],
        )
        if not file_path:
            return

        try:
            keys_copy = db.backup_database(file_path)
        except PermissionError:
            messagebox.showerror(
                "Błąd zapisu",
                "Brak uprawnień do zapisu we wskazanym folderze lub plik jest zablokowany.",
                parent=self,
            )
        except Exception as error:
            messagebox.showerror(
                "Błąd", f"Nie udało się utworzyć kopii zapasowej:\n{error}", parent=self
            )
        else:
            messagebox.showinfo(
                "Kopia zapasowa",
                f"Kopia zapasowa bazy danych została pomyślnie utworzona:\n{file_path}\n\n"
                f"Obok zapisano kopię pliku kluczy:\n{keys_copy}\n\n"
                "Kopia jest zaszyfrowana. Przechowuj oba pliki razem - bez pliku kluczy "
                "kopii nie da się otworzyć na innym komputerze.",
                parent=self,
            )

    def restore_db_backup(self) -> None:
        message = (
            "UWAGA: Ta operacja ZASTĄPI wszystkie obecne dane w programie danymi z wybranego "
            "pliku kopii!\n\nCzy na pewno chcesz kontynuować i wybrać plik kopii zapasowej?"
        )
        if not messagebox.askyesno("Ostrzeżenie", message, icon=messagebox.WARNING, parent=self):
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
        except db.DatabaseError:
            messagebox.showerror(
                "Błąd pliku",
                "Wybrany plik jest uszkodzony, nie jest bazą danych programu albo został "
                "zaszyfrowany innym kluczem (kopia z innej instalacji programu).",
                parent=self,
            )
            return
        except Exception as error:  
            messagebox.showerror(
                "Błąd",
                f"Wystąpił nieoczekiwany błąd podczas przywracania bazy:\n{error}",
                parent=self,
            )
            return

        self._after_phone_removed_from_view()
        messagebox.showinfo(
            "Sukces",
            "Kopia zapasowa została pomyślnie wgrana do bazy danych aplikacji!\n"
            "Lista urządzeń została zaktualizowana.",
            parent=self,
        )

APP = PhoneDatabaseApp