"""Logowanie, pierwsze uruchomienie (szyfrowanie bazy) i zarządzanie użytkownikami."""

from __future__ import annotations

import tkinter as tk
from collections.abc import Callable
from dataclasses import dataclass
from tkinter import messagebox, ttk
from typing import Optional

import db
import security
from dialogs import FONT_FAMILY, FONT_LABEL, BaseDialog

ERROR_COLOR = "#B91C1C"
FONT_CODE = ("Consolas", 13, "bold")


@dataclass
class Session:
    """Zalogowany użytkownik wraz z kluczem głównym bazy."""

    username: str
    master_key: bytes
    keystore: security.KeyStore


# ---------------------------------------------------------------------------
# Funkcje pomocnicze
# ---------------------------------------------------------------------------


def _add_field(parent: tk.Misc, row: int, label: str, *, secret: bool = False) -> ttk.Entry:
    """Dodaje do siatki etykietę i pole tekstowe."""
    ttk.Label(parent, text=label).grid(row=row, column=0, sticky=tk.W, pady=5, padx=(0, 10))
    entry = ttk.Entry(parent, width=30, show="•" if secret else "")
    entry.grid(row=row, column=1, sticky=tk.EW, pady=5)
    return entry


def _validate_username(username: str, parent: tk.Misc) -> bool:
    name = security.normalize_username(username)
    if not name:
        messagebox.showwarning("Brak loginu", "Wpisz login.", parent=parent)
        return False
    if name == security.RECOVERY_USER:
        messagebox.showwarning("Niedozwolony login", "Ten login jest zarezerwowany.", parent=parent)
        return False
    return True


def _validate_new_password(password: str, repeated: str, parent: tk.Misc) -> bool:
    if len(password) < security.MIN_PASSWORD_LENGTH:
        messagebox.showwarning(
            "Za krótkie hasło",
            f"Hasło musi mieć co najmniej {security.MIN_PASSWORD_LENGTH} znaków.",
            parent=parent,
        )
        return False
    if password != repeated:
        messagebox.showwarning("Hasła się różnią", "Wpisane hasła nie są identyczne.", parent=parent)
        return False
    return True


# ---------------------------------------------------------------------------
# Okna wyświetlane przed otwarciem okna głównego
# ---------------------------------------------------------------------------


class _StartupDialog(tk.Toplevel):
    """Okno wyśrodkowane na ekranie - okno główne jest wtedy jeszcze ukryte."""

    def __init__(self, parent: tk.Misc, title: str) -> None:
        super().__init__(parent)
        self.title(title)
        self.resizable(False, False)
        self.session: Optional[Session] = None
        self.protocol("WM_DELETE_WINDOW", self.destroy)
        self.bind("<Escape>", lambda _event: self.destroy())

        self.body = ttk.Frame(self, padding=20)
        self.body.pack(fill=tk.BOTH, expand=True)
        self.body.columnconfigure(1, weight=1)

    def _show(self) -> None:
        self.update_idletasks()
        x = (self.winfo_screenwidth() - self.winfo_reqwidth()) // 2
        y = (self.winfo_screenheight() - self.winfo_reqheight()) // 3
        self.geometry(f"+{max(0, x)}+{max(0, y)}")
        self.lift()
        self.focus_force()

    def _add_buttons(self, row: int, buttons: list[tuple[str, str, Callable[[], None]]]) -> None:
        bar = ttk.Frame(self.body)
        bar.grid(row=row, column=0, columnspan=2, sticky=tk.EW, pady=(14, 0))
        for text, style, command in reversed(buttons):
            ttk.Button(bar, text=text, style=style, command=command).pack(side=tk.RIGHT, padx=(6, 0))


class LoginDialog(_StartupDialog):
    """Logowanie: odszyfrowuje klucz główny i odblokowuje bazę."""

    def __init__(self, parent: tk.Misc, keystore: security.KeyStore) -> None:
        super().__init__(parent, "Logowanie - Baza Telefonów")
        self.keystore = keystore

        ttk.Label(self.body, text="📱 Baza Telefonów", style="BrandTitle.TLabel").grid(
            row=0, column=0, columnspan=2, sticky=tk.W, pady=(0, 12)
        )
        self.username_entry = _add_field(self.body, 1, "Login:")
        self.password_entry = _add_field(self.body, 2, "Hasło:", secret=True)

        self.error_label = ttk.Label(self.body, text="", foreground=ERROR_COLOR)
        self.error_label.grid(row=3, column=0, columnspan=2, sticky=tk.W)

        self._add_buttons(
            4,
            [
                ("Nie pamiętam hasła", "Secondary.TButton", self.open_recovery),
                ("Zaloguj", "Primary.TButton", self.login),
            ],
        )
        self.bind("<Return>", lambda _event: self.login())
        self._show()
        self.username_entry.focus_set()

    def login(self) -> None:
        username = security.normalize_username(self.username_entry.get())
        password = self.password_entry.get()

        try:
            if username == security.RECOVERY_USER:
                raise security.InvalidCredentials
            master_key = self.keystore.unlock(username, password)
        except security.InvalidCredentials:
            self.error_label.configure(text="Nieprawidłowy login lub hasło.")
            self.password_entry.delete(0, tk.END)
            self.password_entry.focus_set()
            return

        try:
            db.unlock(master_key)
        except db.DatabaseError:
            messagebox.showerror(
                "Błąd bazy danych",
                "Plik kluczy nie pasuje do pliku bazy danych (np. podmieniono jeden z plików)."
                f"\n\nBaza: {db.DB_NAME}\nKlucze: {db.KEYSTORE_PATH}",
                parent=self,
            )
            return
        except RuntimeError as error:
            messagebox.showerror("Błąd szyfrowania", str(error), parent=self)
            return

        self.session = Session(username, master_key, self.keystore)
        self.destroy()

    def open_recovery(self) -> None:
        RecoveryDialog(self, self.keystore, on_success=self._after_recovery)

    def _after_recovery(self, username: str) -> None:
        self.error_label.configure(text="")
        self.username_entry.delete(0, tk.END)
        self.username_entry.insert(0, username)
        self.password_entry.delete(0, tk.END)
        self.password_entry.focus_set()


class SetupDialog(_StartupDialog):
    """Pierwsze uruchomienie: konto administratora, klucz główny i kod awaryjny."""

    def __init__(self, parent: tk.Misc, keystore: security.KeyStore) -> None:
        super().__init__(parent, "Pierwsze uruchomienie - Baza Telefonów")
        self.keystore = keystore
        self.migrate_existing = db.is_plain_database()

        info = (
            "Utwórz konto administratora. Dane w bazie będą zaszyfrowane,\n"
            "a dostęp do nich będzie wymagał zalogowania."
        )
        if self.migrate_existing:
            info += "\n\nZnaleziono istniejącą bazę danych - zostanie zaszyfrowana."
        ttk.Label(self.body, text="Konfiguracja zabezpieczeń", font=FONT_LABEL).grid(
            row=0, column=0, columnspan=2, sticky=tk.W
        )
        ttk.Label(self.body, text=info, justify=tk.LEFT).grid(
            row=1, column=0, columnspan=2, sticky=tk.W, pady=(4, 10)
        )

        self.username_entry = _add_field(self.body, 2, "Login administratora:")
        self.password_entry = _add_field(self.body, 3, "Hasło:", secret=True)
        self.repeat_entry = _add_field(self.body, 4, "Powtórz hasło:", secret=True)

        self._add_buttons(
            5,
            [
                ("Anuluj", "Secondary.TButton", self.destroy),
                ("Utwórz konto", "Primary.TButton", self.create_account),
            ],
        )
        self.bind("<Return>", lambda _event: self.create_account())
        self._show()
        self.username_entry.focus_set()

    def create_account(self) -> None:
        username = self.username_entry.get()
        password = self.password_entry.get()
        if not _validate_username(username, self):
            return
        if not _validate_new_password(password, self.repeat_entry.get(), self):
            return

        master_key = security.generate_master_key()
        recovery_code = security.generate_recovery_code()

        def save_keys() -> None:
            self.keystore.add_user(master_key, username, password)
            self.keystore.set_recovery_code(master_key, recovery_code)

        # Kolejność: zaszyfrowana kopia -> zapis kluczy -> podmiana pliku bazy.
        # Klucz jest zapisany, zanim baza zostanie zaszyfrowana, więc nie może przepaść.
        try:
            if self.migrate_existing:
                db.encrypt_existing_database(master_key, before_replace=save_keys)
            else:
                save_keys()
            db.unlock(master_key)
        except Exception as error:  # pokazujemy użytkownikowi każdy błąd konfiguracji
            # Gdy baza nie została zaszyfrowana, zapisane klucze do niej nie pasują - wycofujemy je.
            if not db.database_exists() or db.is_plain_database():
                self.keystore.delete()
            self._show_setup_error(error)
            return

        self.wait_window(RecoveryCodeDialog(self, recovery_code))
        self.session = Session(security.normalize_username(username), master_key, self.keystore)
        self.destroy()

    def _show_setup_error(self, error: BaseException) -> None:
        if isinstance(error, PermissionError):
            messagebox.showerror(
                "Plik bazy jest zablokowany",
                "Nie udało się zaszyfrować bazy, ponieważ plik jest używany przez inny program.\n\n"
                "Sprawdź, czy baza nie jest otwarta w innym programie (np. DB Browser for SQLite) "
                "ani w drugim oknie tej aplikacji, a następnie kliknij „Utwórz konto” ponownie.\n\n"
                f"Dotychczasowa baza pozostała nienaruszona.\n\nSzczegóły: {error}",
                parent=self,
            )
        else:
            messagebox.showerror(
                "Błąd", f"Nie udało się skonfigurować szyfrowania:\n{error}", parent=self
            )


def authenticate(root: tk.Misc) -> Optional[Session]:
    """Pokazuje logowanie (lub konfigurację przy pierwszym uruchomieniu).

    :return: sesja zalogowanego użytkownika albo ``None``, gdy anulowano.
    """
    keystore = security.KeyStore(db.KEYSTORE_PATH)

    if keystore.is_initialized():
        dialog: _StartupDialog = LoginDialog(root, keystore)
    elif db.database_exists() and not db.is_plain_database():
        messagebox.showerror(
            "Brak pliku kluczy",
            "Baza danych jest zaszyfrowana, ale nie znaleziono pliku kluczy:\n"
            f"{db.KEYSTORE_PATH}\n\n"
            "Skopiuj plik kluczy z kopii zapasowej (plik *.keys.json) pod podaną "
            "ścieżkę, nadając mu nazwę keys.json, i uruchom program ponownie.",
            parent=root,
        )
        return None
    else:
        dialog = SetupDialog(root, keystore)

    root.wait_window(dialog)
    return dialog.session


# ---------------------------------------------------------------------------
# Kod awaryjny i odzyskiwanie hasła
# ---------------------------------------------------------------------------


class RecoveryCodeDialog(BaseDialog):
    """Jednorazowe wyświetlenie kodu awaryjnego - zamknięcie wymaga potwierdzenia zapisu."""

    def __init__(self, parent: tk.Misc, code: str) -> None:
        super().__init__(parent, "Kod awaryjny")
        self.code = code
        self.confirmed_var = tk.BooleanVar(value=False)
        self.protocol("WM_DELETE_WINDOW", self.close)
        self.bind("<Escape>", lambda _event: self.close())

        ttk.Label(
            self.body,
            text=(
                "Zapisz ten kod i przechowuj go w bezpiecznym miejscu (np. na papierze w sejfie).\n"
                "Pozwala ustawić nowe hasło, gdy wszyscy zapomną swoich haseł.\n"
                "Bez hasła ani kodu awaryjnego odzyskanie danych NIE będzie możliwe.\n"
                "Kod nie zostanie pokazany ponownie."
            ),
            justify=tk.LEFT,
        ).pack(anchor=tk.W, pady=(0, 10))

        code_entry = ttk.Entry(self.body, font=FONT_CODE, justify=tk.CENTER)
        code_entry.insert(0, code)
        code_entry.configure(state="readonly")
        code_entry.pack(fill=tk.X, ipady=4)

        ttk.Button(
            self.body, text="📋 Kopiuj do schowka", style="Secondary.TButton", command=self.copy
        ).pack(anchor=tk.E, pady=6)
        ttk.Checkbutton(
            self.body,
            text="Zapisałem kod awaryjny w bezpiecznym miejscu",
            variable=self.confirmed_var,
            command=self._update_button,
        ).pack(anchor=tk.W, pady=(4, 0))

        bar = ttk.Frame(self.body)
        bar.pack(fill=tk.X, pady=(12, 0))
        self.ok_button = ttk.Button(
            bar, text="Zamknij", style="Primary.TButton", command=self.close, state=tk.DISABLED
        )
        self.ok_button.pack(side=tk.RIGHT)
        self._show(560, 300)

    def copy(self) -> None:
        self.clipboard_clear()
        self.clipboard_append(self.code)

    def _update_button(self) -> None:
        self.ok_button.configure(state=tk.NORMAL if self.confirmed_var.get() else tk.DISABLED)

    def close(self) -> None:
        if not self.confirmed_var.get():
            messagebox.showwarning(
                "Zapisz kod", "Potwierdź, że kod awaryjny został zapisany.", parent=self
            )
            return
        self.destroy()


class RecoveryDialog(BaseDialog):
    """Ustawienie nowego hasła dla istniejącego użytkownika za pomocą kodu awaryjnego."""

    def __init__(
        self,
        parent: tk.Misc,
        keystore: security.KeyStore,
        on_success: Optional[Callable[[str], None]] = None,
    ) -> None:
        super().__init__(parent, "Odzyskiwanie hasła")
        self.keystore = keystore
        self.on_success = on_success

        form = ttk.Frame(self.body)
        form.pack(fill=tk.BOTH, expand=True)
        form.columnconfigure(1, weight=1)
        self.username_entry = _add_field(form, 0, "Login:")
        self.code_entry = _add_field(form, 1, "Kod awaryjny:")
        self.password_entry = _add_field(form, 2, "Nowe hasło:", secret=True)
        self.repeat_entry = _add_field(form, 3, "Powtórz hasło:", secret=True)

        self._add_button_bar("Ustaw nowe hasło", self.reset_password)
        self.username_entry.focus_set()
        self._show(440, 260)

    def reset_password(self) -> None:
        username = security.normalize_username(self.username_entry.get())
        password = self.password_entry.get()
        if not _validate_username(username, self):
            return
        if not self.keystore.has_user(username):
            messagebox.showwarning("Nieznany login", "Nie ma użytkownika o takim loginie.", parent=self)
            return
        if not _validate_new_password(password, self.repeat_entry.get(), self):
            return

        try:
            master_key = self.keystore.unlock_with_recovery_code(self.code_entry.get())
        except security.InvalidCredentials:
            messagebox.showerror("Błędny kod", "Nieprawidłowy kod awaryjny.", parent=self)
            return

        self.keystore.add_user(master_key, username, password)
        messagebox.showinfo(
            "Hasło zmienione", "Nowe hasło zostało ustawione. Zaloguj się nim.", parent=self
        )
        self.destroy()
        if self.on_success is not None:
            self.on_success(username)


# ---------------------------------------------------------------------------
# Zarządzanie użytkownikami (z okna głównego)
# ---------------------------------------------------------------------------


class ChangePasswordDialog(BaseDialog):
    """Zmiana hasła zalogowanego użytkownika."""

    def __init__(self, parent: tk.Misc, session: Session) -> None:
        super().__init__(parent, "Zmiana hasła")
        self.session = session

        form = ttk.Frame(self.body)
        form.pack(fill=tk.BOTH, expand=True)
        form.columnconfigure(1, weight=1)
        ttk.Label(form, text=f"Użytkownik: {session.username}", font=FONT_LABEL).grid(
            row=0, column=0, columnspan=2, sticky=tk.W, pady=(0, 6)
        )
        self.old_entry = _add_field(form, 1, "Obecne hasło:", secret=True)
        self.new_entry = _add_field(form, 2, "Nowe hasło:", secret=True)
        self.repeat_entry = _add_field(form, 3, "Powtórz nowe hasło:", secret=True)

        self._add_button_bar("💾 Zmień hasło", self.change_password)
        self.old_entry.focus_set()
        self._show(440, 240)

    def change_password(self) -> None:
        new_password = self.new_entry.get()
        if not _validate_new_password(new_password, self.repeat_entry.get(), self):
            return
        try:
            self.session.keystore.change_password(
                self.session.username, self.old_entry.get(), new_password
            )
        except security.InvalidCredentials:
            messagebox.showerror("Błędne hasło", "Obecne hasło jest nieprawidłowe.", parent=self)
            return

        messagebox.showinfo("Zapisano", "Hasło zostało zmienione.", parent=self)
        self.destroy()


class AddUserDialog(BaseDialog):
    """Dodanie nowego użytkownika z dostępem do bazy."""

    def __init__(self, parent: tk.Misc, session: Session) -> None:
        super().__init__(parent, "Dodaj użytkownika")
        self.session = session

        form = ttk.Frame(self.body)
        form.pack(fill=tk.BOTH, expand=True)
        form.columnconfigure(1, weight=1)
        self.username_entry = _add_field(form, 0, "Login:")
        self.password_entry = _add_field(form, 1, "Hasło:", secret=True)
        self.repeat_entry = _add_field(form, 2, "Powtórz hasło:", secret=True)

        self._add_button_bar("👥 Dodaj", self.add_user)
        self.username_entry.focus_set()
        self._show(420, 210)

    def add_user(self) -> None:
        username = self.username_entry.get()
        password = self.password_entry.get()
        if not _validate_username(username, self):
            return
        if self.session.keystore.has_user(username):
            messagebox.showwarning("Login zajęty", "Taki użytkownik już istnieje.", parent=self)
            return
        if not _validate_new_password(password, self.repeat_entry.get(), self):
            return

        self.session.keystore.add_user(self.session.master_key, username, password)
        messagebox.showinfo(
            "Dodano", f"Użytkownik „{security.normalize_username(username)}” został dodany.", parent=self
        )
        self.destroy()


class RemoveUserDialog(BaseDialog):
    """Usunięcie innego użytkownika (zablokowanie mu logowania)."""

    def __init__(self, parent: tk.Misc, session: Session, usernames: list[str]) -> None:
        super().__init__(parent, "Usuń użytkownika")
        self.session = session

        ttk.Label(self.body, text="Wybierz użytkownika do usunięcia:", font=FONT_LABEL).pack(
            anchor=tk.W, pady=(0, 6)
        )
        self.user_combo = ttk.Combobox(self.body, values=usernames, state="readonly")
        self.user_combo.current(0)
        self.user_combo.pack(fill=tk.X)
        ttk.Label(
            self.body,
            text="Usunięty użytkownik nie będzie mógł się zalogować.",
            font=(FONT_FAMILY, 8, "italic"),
        ).pack(anchor=tk.W, pady=(6, 0))

        self._add_button_bar("🚫 Usuń", self.remove_user)
        self._show(400, 180)

    def remove_user(self) -> None:
        username = self.user_combo.get()
        if not messagebox.askyesno(
            "Potwierdzenie", f"Czy na pewno usunąć użytkownika „{username}”?", parent=self
        ):
            return
        self.session.keystore.remove_user(username)
        messagebox.showinfo("Usunięto", f"Użytkownik „{username}” został usunięty.", parent=self)
        self.destroy()


def regenerate_recovery_code(parent: tk.Misc, session: Session) -> None:
    """Tworzy nowy kod awaryjny (stary przestaje działać) i pokazuje go użytkownikowi."""
    if not messagebox.askyesno(
        "Nowy kod awaryjny",
        "Zostanie wygenerowany nowy kod awaryjny, a dotychczasowy przestanie działać.\n\n"
        "Czy kontynuować?",
        parent=parent,
    ):
        return
    code = security.generate_recovery_code()
    session.keystore.set_recovery_code(session.master_key, code)
    RecoveryCodeDialog(parent, code)
