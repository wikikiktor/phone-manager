from __future__ import annotations

import tkinter as tk
from collections.abc import Callable, Sequence
from datetime import datetime
from pathlib import Path
from tkinter import filedialog, messagebox, ttk
from typing import Any, Optional

import openpyxl

import db
import protocol_generator

Callback = Optional[Callable[[], None]]

FONT_FAMILY = "Segoe UI"
FONT_TEXT = (FONT_FAMILY, 9)
FONT_LABEL = (FONT_FAMILY, 9, "bold")

PHONES_SHEET_NAME = "Telefony"
HISTORY_SHEET_NAME = "Historia zdarzeń"
IGNORE_COLUMN = "[Ignoruj / brak]"

EVENT_CATEGORIES: tuple[str, ...] = (
    "Notatka / Uwaga",
    "Zmiana użytkownika",
    "Awaria / Serwis",
    "Wydanie",
    "Zwrot",
    "Wymiana karty SIM",
    "Inne",
)

PROTOCOL_TYPES: dict[str, tuple[str, str]] = {
    "Protokół przekazania": (protocol_generator.PROTOCOL_HANDOVER, "Protokol_przekazania"),
    "Protokół zwrotu": (protocol_generator.PROTOCOL_RETURN, "Protokol_zwrotu"),
}

PROTOCOL_FORMATS: dict[str, tuple[str, Callable[..., None]]] = {
    "pdf": ("Dokument PDF (*.pdf)", protocol_generator.generate_pdf),
    "docx": ("Dokument Word (*.docx)", protocol_generator.generate_docx),
}


def center_dialog(window: tk.Toplevel, width: int, height: int, parent: tk.Misc) -> None:
    window.update_idletasks()
    x = parent.winfo_rootx() + (parent.winfo_width() - width) // 2
    y = parent.winfo_rooty() + (parent.winfo_height() - height) // 2
    window.geometry(f"{width}x{height}+{max(0, x)}+{max(0, y)}")


def _run_callback(callback: Callback) -> None:
    if callback is not None:
        callback()


def _excel_value_to_text(value: Any) -> str:
    if value is None:
        return ""
    if isinstance(value, float) and value.is_integer():
        value = int(value)
    return str(value).strip()


# ---------------------------------------------------------------------------
# Klasa bazowa
# ---------------------------------------------------------------------------

class _BaseDialog(tk.Toplevel):

    def __init__(self, parent: tk.Misc, title: str) -> None:
        super().__init__(parent)
        self.title(title)
        self.transient(parent)
        self.grab_set()
        self.bind("<Escape>", lambda _event: self.destroy())

        self.body = ttk.Frame(self, padding=16)
        self.body.pack(fill=tk.BOTH, expand=True)

    def _show(self, width: int, height: int) -> None:
        center_dialog(self, width, height, self.master)

    def _add_button_bar(
        self,
        primary_text: str,
        primary_command: Callable[[], None],
        *,
        use_grid_row: Optional[int] = None,
    ) -> ttk.Frame:
        bar = ttk.Frame(self.body)
        if use_grid_row is None:
            bar.pack(fill=tk.X, pady=(10, 0))
        else:
            bar.grid(row=use_grid_row, column=0, columnspan=2, pady=(16, 0), sticky=tk.E)

        ttk.Button(bar, text="Anuluj", style="Secondary.TButton", command=self.destroy).pack(
            side=tk.RIGHT, padx=4
        )
        ttk.Button(bar, text=primary_text, style="Primary.TButton", command=primary_command).pack(
            side=tk.RIGHT, padx=4
        )
        return bar

    def _add_text_area(self, label: str, height: int, initial_text: str = "") -> tk.Text:
        ttk.Label(self.body, text=label, font=FONT_LABEL).pack(anchor=tk.W, pady=(0, 4))
        text = tk.Text(
            self.body, height=height, wrap=tk.WORD, font=FONT_TEXT, relief="solid", borderwidth=1
        )
        text.pack(fill=tk.BOTH, expand=True, pady=(0, 2))
        if initial_text:
            text.insert("1.0", initial_text)
        text.focus_set()
        return text

    @staticmethod
    def _read_text(widget: tk.Text) -> str:
        return widget.get("1.0", tk.END).strip()

# ---------------------------------------------------------------------------
# Historia zdarzeń
# ---------------------------------------------------------------------------

class AddEventDialog(_BaseDialog):
    def __init__(self, parent: tk.Misc, phone_id: int, on_save_callback: Callback = None) -> None:
        super().__init__(parent, "Dodaj zdarzenie / uwagę")
        self.phone_id = phone_id
        self.on_save_callback = on_save_callback

        ttk.Label(self.body, text="Kategoria zdarzenia:", font=FONT_LABEL).pack(
            anchor=tk.W, pady=(0, 4)
        )
        self.category_combo = ttk.Combobox(self.body, values=EVENT_CATEGORIES, state="readonly")
        self.category_combo.set(EVENT_CATEGORIES[0])
        self.category_combo.pack(fill=tk.X, pady=(0, 10))

        self.description_text = self._add_text_area("Opis zdarzenia / notatki:", height=6)
        self._add_button_bar("💾 Zapisz wpis", self.save_event)
        self._show(460, 320)

    def save_event(self) -> None:
        description = self._read_text(self.description_text)
        if not description:
            messagebox.showwarning("Puste pole", "Wpisz treść zdarzenia lub uwagi.", parent=self)
            return

        db.add_history_entry(self.phone_id, self.category_combo.get(), description)
        self.destroy()
        _run_callback(self.on_save_callback)


class AddNoteDialog(_BaseDialog):
    def __init__(
        self,
        parent: tk.Misc,
        history_id: int,
        current_note: str,
        on_save_callback: Callback = None,
    ) -> None:
        super().__init__(parent, "Edycja uwagi do zdarzenia")
        self.history_id = history_id
        self.on_save_callback = on_save_callback

        self.note_text = self._add_text_area(
            "Treść uwagi do zdarzenia:", height=5, initial_text=current_note or ""
        )
        self._add_button_bar("💾 Zapisz uwagę", self.save_note)
        self._show(440, 260)

    def save_note(self) -> None:
        db.update_history_note(self.history_id, self._read_text(self.note_text))
        self.destroy()
        _run_callback(self.on_save_callback)


# ---------------------------------------------------------------------------
# Import z Excela
# ---------------------------------------------------------------------------

class ExcelImportDialog(_BaseDialog):
    """Import telefonów (i opcjonalnie historii) z pliku Excel z mapowaniem kolumn."""

    TARGET_FIELDS: tuple[tuple[str, str], ...] = (
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
    )
    REQUIRED_FIELDS = ("model", "nr_tel")

    HISTORY_COLUMN_KEYWORDS: dict[str, tuple[tuple[str, ...], tuple[str, ...]]] = {
        "data": (("data",), ()),
        "nr_tel": (("nr tel", "numer tel", "nr_tel", "telefon"), ("model",)),
        "kategoria": (("kategoria",), ()),
        "opis": (("opis zdarzenia", "opis"), ()),
        "uwagi": (("uwagi", "uwaga"), ()),
    }

    def __init__(
        self, parent: tk.Misc, file_path: str | Path, on_success_callback: Callback = None
    ) -> None:
        super().__init__(parent, "Import danych z Excela")
        self.file_path = file_path
        self.on_success_callback = on_success_callback

        self.workbook: Any = None
        self.sheet: Any = None
        self.headers: list[str] = []
        self.combos: dict[str, ttk.Combobox] = {}
        self.has_history_sheet = False
        self.import_history_var = tk.BooleanVar(value=False)

        if self._load_workbook():
            self._build_ui()
            self._show(620, 480)

   
# ---------------------------------------------------------------------------
# Wczytywanie pliku
# ---------------------------------------------------------------------------

    def _load_workbook(self) -> bool:
        try:
            self.workbook = openpyxl.load_workbook(self.file_path, data_only=True)
        except Exception as error:  # openpyxl zgłasza różne typy wyjątków dla złych plików
            self._fail(f"Nie można wczytać pliku Excel:\n{error}")
            return False

        sheet_names = self.workbook.sheetnames
        self.sheet = (
            self.workbook[PHONES_SHEET_NAME]
            if PHONES_SHEET_NAME in sheet_names
            else self.workbook.active
        )
        self.has_history_sheet = HISTORY_SHEET_NAME in sheet_names
        self.import_history_var.set(self.has_history_sheet)

        first_row = next(self.sheet.iter_rows(values_only=True), None)
        if not first_row:
            self._fail("Plik Excel jest pusty.")
            return False

        self.headers = [
            _excel_value_to_text(value) or f"Kolumna {index}"
            for index, value in enumerate(first_row, start=1)
        ]
        return True

    def _fail(self, message: str) -> None:
        messagebox.showerror("Błąd", message, parent=self.master)
        self.destroy()

    
# ---------------------------------------------------------------------------
# Interfejs
# ---------------------------------------------------------------------------

    def _build_ui(self) -> None:
        ttk.Label(
            self.body,
            text=(
                "Przypisz kolumny z arkusza Excel do odpowiednich pól w bazie.\n"
                "Pola oznaczone gwiazdką (*) są wymagane."
            ),
            font=FONT_LABEL,
            justify=tk.LEFT,
        ).pack(anchor=tk.W, pady=(0, 10))

        form = ttk.Frame(self.body)
        form.pack(fill=tk.BOTH, expand=True)
        form.columnconfigure(1, weight=1)

        options = [IGNORE_COLUMN, *self.headers]
        for row, (field_key, field_label) in enumerate(self.TARGET_FIELDS):
            ttk.Label(form, text=field_label).grid(row=row, column=0, sticky=tk.W, pady=3, padx=4)
            combo = ttk.Combobox(form, values=options, state="readonly", width=30)
            combo.set(self._guess_column(field_label))
            combo.grid(row=row, column=1, sticky=tk.EW, pady=3, padx=4)
            self.combos[field_key] = combo

        if self.has_history_sheet:
            ttk.Checkbutton(
                self.body,
                text=f"Importuj również historię zdarzeń (znaleziono arkusz '{HISTORY_SHEET_NAME}')",
                variable=self.import_history_var,
            ).pack(anchor=tk.W, pady=8)

        self._add_button_bar("📥 Importuj dane", self.import_data)

    def _guess_column(self, field_label: str) -> str:
        """Dobiera kolumnę Excela, której nagłówek pasuje do etykiety pola."""
        target = field_label.replace("*", "").strip().lower()
        for header in self.headers:
            normalized = header.lower()
            if normalized and (normalized in target or target in normalized):
                return header
        return IGNORE_COLUMN

# ---------------------------------------------------------------------------
# Import
# ---------------------------------------------------------------------------

    def _column_mapping(self) -> dict[str, int]:
        return {
            field_key: self.headers.index(combo.get())
            for field_key, combo in self.combos.items()
            if combo.get() != IGNORE_COLUMN
        }

    def _read_phone_records(self, mapping: dict[str, int]) -> list[dict[str, str]]:
        records = []
        for row in self.sheet.iter_rows(min_row=2, values_only=True):
            if not any(row):
                continue
            records.append(
                {
                    field_key: _excel_value_to_text(row[index]) if index < len(row) else ""
                    for field_key, index in mapping.items()
                }
            )
        return records

    def import_data(self) -> None:
        mapping = self._column_mapping()
        if any(field not in mapping for field in self.REQUIRED_FIELDS):
            messagebox.showwarning(
                "Brakujące mapowanie",
                "Wymagane jest przypisanie kolumn: 'Model telefonu' oraz 'Numer telefonu'.",
                parent=self,
            )
            return

        phone_records = self._read_phone_records(mapping)
        if not phone_records:
            messagebox.showinfo(
                "Brak danych", "Nie znaleziono żadnych danych do zaimportowania.", parent=self
            )
            return

        history_records = self._read_history_records()
        phones_count, history_count = db.bulk_insert_phones(phone_records, history_records)

        message = f"Pomyślnie zaimportowano {phones_count} telefon(ów)."
        if history_records is not None:
            message += f"\nZaimportowano również {history_count} wpis(ów) historii."
        messagebox.showinfo("Sukces", message, parent=self)

        self.destroy()
        _run_callback(self.on_success_callback)

    def _read_history_records(self) -> Optional[list[dict[str, str]]]:
        """Odczytuje arkusz historii; ``None`` oznacza brak historii do importu."""
        if not self.has_history_sheet or not self.import_history_var.get():
            return None

        rows = list(self.workbook[HISTORY_SHEET_NAME].iter_rows(values_only=True))
        if len(rows) < 2:
            return None

        headers = [_excel_value_to_text(value).lower() for value in rows[0]]
        columns = {
            field: self._find_column(headers, keywords, excluded)
            for field, (keywords, excluded) in self.HISTORY_COLUMN_KEYWORDS.items()
        }
        if columns["nr_tel"] is None:
            return None

        def cell(row: Sequence[Any], field: str) -> Any:
            index = columns[field]
            return row[index] if index is not None and index < len(row) else None

        records = []
        for row in rows[1:]:
            if not any(row):
                continue
            date_value = cell(row, "data")
            records.append(
                {
                    "data": (
                        date_value.strftime(db.DATETIME_FORMAT)
                        if isinstance(date_value, datetime)
                        else _excel_value_to_text(date_value)
                    ),
                    "nr_tel": _excel_value_to_text(cell(row, "nr_tel")),
                    "kategoria": _excel_value_to_text(cell(row, "kategoria")) or "Import",
                    "opis": _excel_value_to_text(cell(row, "opis")),
                    "uwagi": _excel_value_to_text(cell(row, "uwagi")),
                }
            )
        return records

    @staticmethod
    def _find_column(
        headers: Sequence[str], keywords: Sequence[str], excluded: Sequence[str] = ()
    ) -> Optional[int]:
        for index, header in enumerate(headers):
            if any(word in header for word in excluded):
                continue
            if any(word in header for word in keywords):
                return index
        return None


# ---------------------------------------------------------------------------
# Dane pracodawcy
# ---------------------------------------------------------------------------

class EmployerDialog(_BaseDialog):
    """Edycja danych pracodawcy wykorzystywanych w protokołach."""

    FIELDS: tuple[tuple[str, str], ...] = (
        ("nazwa_firmy", "Nazwa firmy:"),
        ("adres", "Adres firmy:"),
        ("nip", "NIP:"),
    )

    def __init__(self, parent: tk.Misc) -> None:
        super().__init__(parent, "Dane pracodawcy / firmy")
        self.entries: dict[str, ttk.Entry] = {}

        for row, (key, label) in enumerate(self.FIELDS):
            ttk.Label(self.body, text=label).grid(row=row, column=0, sticky=tk.W, pady=6)
            entry = ttk.Entry(self.body, width=32)
            entry.grid(row=row, column=1, sticky=tk.EW, pady=6, padx=(6, 0))
            self.entries[key] = entry
        self.body.columnconfigure(1, weight=1)

        self._add_button_bar("💾 Zapisz dane", self.save_data, use_grid_row=len(self.FIELDS))
        self._load_data()
        self.entries["nazwa_firmy"].focus_set()
        self._show(440, 250)

    def _load_data(self) -> None:
        employer = db.get_employer()
        if not employer:
            return
        for (key, _label), value in zip(self.FIELDS, employer):
            if value:
                self.entries[key].insert(0, value)

    def save_data(self) -> None:
        values = {key: entry.get().strip() for key, entry in self.entries.items()}
        db.save_employer(values["nazwa_firmy"], values["adres"], values["nip"])
        messagebox.showinfo(
            "Zapisano", "Dane pracodawcy zostały pomyślnie zaktualizowane.", parent=self
        )
        self.destroy()


# ---------------------------------------------------------------------------
# Protokoły
# ---------------------------------------------------------------------------

class ProtocolDialog(_BaseDialog):
    def __init__(self, parent: tk.Misc, initial_phone_id: Optional[int] = None) -> None:
        super().__init__(parent, "Generowanie protokołu telefonu")
        self.initial_phone_id = initial_phone_id
        self._phone_ids: list[int] = []
        self.preview_labels: dict[str, ttk.Label] = {}

        self._build_ui()
        self.update_preview()
        self._show(540, 360)

    def _build_ui(self) -> None:
        body = self.body
        body.columnconfigure(1, weight=1)

        ttk.Label(body, text="Typ protokołu:").grid(row=0, column=0, sticky=tk.W, pady=6)
        type_labels = list(PROTOCOL_TYPES)
        self.type_combo = ttk.Combobox(body, values=type_labels, state="readonly", width=36)
        self.type_combo.set(type_labels[0])
        self.type_combo.grid(row=0, column=1, sticky=tk.EW, pady=6, padx=(6, 0))

        ttk.Label(body, text="Wybierz telefon:").grid(row=1, column=0, sticky=tk.W, pady=6)
        self.phone_combo = ttk.Combobox(body, state="readonly", width=36)
        self.phone_combo.grid(row=1, column=1, sticky=tk.EW, pady=6, padx=(6, 0))
        self.phone_combo.bind("<<ComboboxSelected>>", self.update_preview)
        self._load_phones()

        preview = ttk.LabelFrame(
            body,
            text=" Podgląd danych wybranego telefonu ",
            style="Card.TLabelframe",
            padding=10,
        )
        preview.grid(row=2, column=0, columnspan=2, sticky=tk.EW, pady=12)
        for key in ("user", "model", "equipment", "battery"):
            label = ttk.Label(preview, style="Card.TLabel")
            label.pack(anchor=tk.W, pady=1)
            self.preview_labels[key] = label
        self._set_preview("-", "-", "-", "-", "-")

        buttons = ttk.Frame(body)
        buttons.grid(row=3, column=0, columnspan=2, pady=(10, 0), sticky=tk.E)
        ttk.Button(buttons, text="Anuluj", style="Secondary.TButton", command=self.destroy).pack(
            side=tk.RIGHT, padx=4
        )
        ttk.Button(
            buttons,
            text="📄 Pobierz PDF",
            style="Primary.TButton",
            command=lambda: self.export_protocol("pdf"),
        ).pack(side=tk.RIGHT, padx=4)
        ttk.Button(
            buttons,
            text="📝 Pobierz DOCX",
            style="Secondary.TButton",
            command=lambda: self.export_protocol("docx"),
        ).pack(side=tk.RIGHT, padx=4)

    def _load_phones(self) -> None:
        labels = []
        selected_index = 0
        for index, (phone_id, nr_tel, user, _model, _damaged) in enumerate(db.search_phones("")):
            self._phone_ids.append(phone_id)
            labels.append(f"{nr_tel} | {user or '[Brak użytkownika]'}")
            if phone_id == self.initial_phone_id:
                selected_index = index

        self.phone_combo.configure(values=labels)
        if labels:
            self.phone_combo.current(selected_index)

    def get_selected_phone_id(self) -> Optional[int]:
        index = self.phone_combo.current()
        return self._phone_ids[index] if index >= 0 else None

    def _set_preview(self, user: str, model: str, imei: str, equipment: str, battery: str) -> None:
        self.preview_labels["user"].configure(text=f"Użytkownik: {user}")
        self.preview_labels["model"].configure(text=f"Model: {model} | IMEI: {imei}")
        self.preview_labels["equipment"].configure(text=f"Wyposażenie: {equipment}")
        self.preview_labels["battery"].configure(text=f"Stan baterii: {battery}")

    def update_preview(self, _event: Optional[tk.Event] = None) -> None:
        phone_id = self.get_selected_phone_id()
        row = db.get_phone_by_id(phone_id) if phone_id is not None else None
        if row is None:
            return

        phone = dict(zip(db.PHONE_FIELDS, row))
        self._set_preview(
            user=phone["osoba_uzytkujaca"] or "[BRAK]",
            model=phone["model"],
            imei=phone["imei"] or "[BRAK]",
            equipment=phone["wyposazenie"] or "Brak",
            battery=phone["stan_baterii"] or db.DEFAULT_BATTERY_STATE,
        )

    def export_protocol(self, extension: str) -> None:
        phone_id = self.get_selected_phone_id()
        if phone_id is None:
            messagebox.showwarning("Brak telefonu", "Wybierz telefon z listy.", parent=self)
            return

        phone_row = db.get_phone_by_id(phone_id)
        employer_row = db.get_employer()
        protocol_type, file_prefix = PROTOCOL_TYPES[self.type_combo.get()]
        file_description, generate = PROTOCOL_FORMATS[extension]

        file_path = filedialog.asksaveasfilename(
            parent=self,
            title=f"Zapisz protokół ({extension.upper()})",
            defaultextension=f".{extension}",
            initialfile=self._default_file_name(file_prefix, phone_row[0], extension),
            filetypes=[(file_description, f"*.{extension}")],
        )
        if not file_path:
            return

        try:
            generate(file_path, phone_row, employer_row, protocol_type)
        except PermissionError:
            messagebox.showerror(
                "Błąd zapisu",
                "Plik jest otwarty w innym programie (np. Word / Adobe Reader). "
                "Zamknij go i spróbuj ponownie.",
                parent=self,
            )
            return
        except Exception as error:
            messagebox.showerror(
                "Błąd", f"Wystąpił błąd podczas generowania protokołu:\n{error}", parent=self
            )
            return

        messagebox.showinfo(
            "Sukces",
            f"Protokół został pomyślnie wygenerowany i zapisany:\n{file_path}",
            parent=self,
        )
        self.destroy()

    @staticmethod
    def _default_file_name(prefix: str, model: Optional[str], extension: str) -> str:
        safe_model = "".join(char for char in (model or "") if char.isalnum() or char in "-_")
        date = datetime.now().strftime("%Y%m%d")  # noqa: DTZ005
        return f"{prefix}_{safe_model}_{date}.{extension}"
