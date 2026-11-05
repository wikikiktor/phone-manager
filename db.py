"""Warstwa dostępu do danych aplikacji „Baza Telefonów”.

Moduł odpowiada za:
    * inicjalizację i migrację schematu bazy SQLite,
    * operacje CRUD na telefonach służbowych wraz z historią zdarzeń,
    * import masowy oraz eksport danych do pliku Excel,
    * dane pracodawcy oraz kopie zapasowe bazy,
    * szyfrowanie bazy (SQLCipher) kluczem odblokowywanym przy logowaniu.
"""

from __future__ import annotations

import logging
import os
import re
import shutil
import time
from collections.abc import Callable, Iterable, Iterator, Mapping, Sequence
from contextlib import closing, contextmanager
from datetime import datetime
from pathlib import Path
from typing import Any

import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.worksheet import Worksheet
from sqlcipher3 import dbapi2 as sqlite3

# Wyjątek bazy danych SQLCipher (inna klasa niż ``sqlite3.DatabaseError`` z biblioteki standardowej).
DatabaseError = sqlite3.DatabaseError

logger = logging.getLogger(__name__)

# ---------------------------------------------------------------------------
# Konfiguracja
# ---------------------------------------------------------------------------

APP_DIR_NAME = "BazaTelefonow"
FALLBACK_APP_DIR_NAME = ".baza_telefonow"
DB_FILENAME = "telefony.db"
KEYSTORE_FILENAME = "keys.json"
BACKUP_KEYS_SUFFIX = ".keys.json"

# Nagłówek niezaszyfrowanego pliku SQLite (zaszyfrowany plik wygląda jak losowe dane).
_PLAIN_SQLITE_HEADER = b"SQLite format 3\x00"

DATETIME_FORMAT = "%Y-%m-%d %H:%M"
DEFAULT_BATTERY_STATE = "Brak informacji"
TRUTHY_STRINGS = frozenset({"tak", "1", "true", "t", "x"})

# Kolejność pól odpowiada kolejności kolumn w zapytaniach INSERT/UPDATE/SELECT.
PHONE_FIELDS: tuple[str, ...] = (
    "model",
    "nr_tel",
    "nr_sim",
    "imei",
    "nr_seryjny",
    "rodzaj",
    "osoba_uzytkujaca",
    "osoba_odpowiedzialna",
    "czy_protokol",
    "wyposazenie",
    "stan_baterii",
    "czy_uszkodzony",
)

# Pola logiczne (0/1) prezentowane użytkownikowi jako „Tak”/„Nie”.
FLAG_FIELDS = frozenset({"czy_protokol", "czy_uszkodzony"})

# Wartości domyślne dla pól opcjonalnych.
OPTIONAL_FIELD_DEFAULTS: dict[str, Any] = {
    "czy_protokol": 0,
    "wyposazenie": "",
    "stan_baterii": DEFAULT_BATTERY_STATE,
    "czy_uszkodzony": 0,
}

FIELD_LABELS: dict[str, str] = {
    "model": "Model",
    "nr_tel": "Nr telefonu",
    "nr_sim": "Nr SIM",
    "imei": "IMEI",
    "nr_seryjny": "Nr seryjny",
    "rodzaj": "Rodzaj",
    "osoba_uzytkujaca": "Osoba użytkująca",
    "osoba_odpowiedzialna": "Osoba odpowiedzialna",
    "czy_protokol": "Protokół zdawczo-odbiorczy",
    "wyposazenie": "Wyposażenie dodatkowe",
    "stan_baterii": "Stan baterii",
    "czy_uszkodzony": "Czy uszkodzony",
}

CHANGE_CATEGORIES: dict[str, str] = {
    "osoba_uzytkujaca": "Zmiana użytkownika",
    "osoba_odpowiedzialna": "Zmiana odpowiedzialnego",
    "nr_sim": "Wymiana karty SIM",
    "rodzaj": "Zmiana statusu/rodzaju",
    "czy_protokol": "Zmiana statusu protokołu",
    "wyposazenie": "Zmiana wyposażenia",
    "stan_baterii": "Zmiana stanu baterii",
}
DEFAULT_CHANGE_CATEGORY = "Edycja danych"

# ---------------------------------------------------------------------------
# Schemat bazy danych
# ---------------------------------------------------------------------------

_SCHEMA: tuple[str, ...] = (
    """
    CREATE TABLE IF NOT EXISTS telefony (
        id                   INTEGER PRIMARY KEY AUTOINCREMENT,
        model                TEXT NOT NULL,
        nr_tel               TEXT NOT NULL,
        nr_sim               TEXT,
        imei                 TEXT,
        nr_seryjny           TEXT,
        rodzaj               TEXT,
        osoba_uzytkujaca     TEXT,
        osoba_odpowiedzialna TEXT
    )
    """,
    """
    CREATE TABLE IF NOT EXISTS historia (
        id         INTEGER PRIMARY KEY AUTOINCREMENT,
        telefon_id INTEGER NOT NULL,
        data       TEXT NOT NULL,
        kategoria  TEXT NOT NULL,
        opis       TEXT,
        uwagi      TEXT DEFAULT '',
        FOREIGN KEY (telefon_id) REFERENCES telefony (id)
    )
    """,
    """
    CREATE TABLE IF NOT EXISTS pracodawca (
        id          INTEGER PRIMARY KEY AUTOINCREMENT,
        nazwa_firmy TEXT,
        adres       TEXT,
        nip         TEXT
    )
    """,
)

# Kolumny dodane w kolejnych wersjach aplikacji: (tabela, kolumna, definicja).
_MIGRATIONS: tuple[tuple[str, str, str], ...] = (
    ("telefony", "czy_usuniety", "INTEGER DEFAULT 0"),
    ("telefony", "czy_protokol", "INTEGER DEFAULT 0"),
    ("telefony", "wyposazenie", "TEXT DEFAULT ''"),
    ("telefony", "stan_baterii", f"TEXT DEFAULT '{DEFAULT_BATTERY_STATE}'"),
    ("telefony", "czy_uszkodzony", "INTEGER DEFAULT 0"),
    ("historia", "uwagi", "TEXT DEFAULT ''"),
)

_PHONE_COLUMNS_SQL = f"""
    model, nr_tel, nr_sim, imei, nr_seryjny, rodzaj, osoba_uzytkujaca, osoba_odpowiedzialna,
    COALESCE(czy_protokol, 0),
    COALESCE(wyposazenie, ''),
    COALESCE(stan_baterii, '{DEFAULT_BATTERY_STATE}'),
    COALESCE(czy_uszkodzony, 0)
"""

_INSERT_PHONE_SQL = (
    f"INSERT INTO telefony ({', '.join(PHONE_FIELDS)}) "
    f"VALUES ({', '.join('?' for _ in PHONE_FIELDS)})"
)

_UPDATE_PHONE_SQL = (
    f"UPDATE telefony SET {', '.join(f'{field} = ?' for field in PHONE_FIELDS)} "
    "WHERE id = ?"
)

_INSERT_HISTORY_SQL = """
    INSERT INTO historia (telefon_id, data, kategoria, opis, uwagi)
    VALUES (?, ?, ?, ?, ?)
"""

# ---------------------------------------------------------------------------
# Styl eksportu do Excela
# ---------------------------------------------------------------------------

_HEADER_FONT = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
_PHONES_HEADER_FILL = PatternFill(start_color="1F4E79", end_color="1F4E79", fill_type="solid")
_HISTORY_HEADER_FILL = PatternFill(start_color="2F5597", end_color="2F5597", fill_type="solid")
_THIN_SIDE = Side(style="thin", color="D9D9D9")
_CELL_BORDER = Border(left=_THIN_SIDE, right=_THIN_SIDE, top=_THIN_SIDE, bottom=_THIN_SIDE)
_ALIGN_CENTER = Alignment(horizontal="center", vertical="center")
_ALIGN_LEFT = Alignment(horizontal="left", vertical="center")
_MIN_COLUMN_WIDTH = 12
_COLUMN_PADDING = 4

_PHONES_SHEET_HEADERS: tuple[str, ...] = (
    "ID",
    "Model telefonu",
    "Numer telefonu",
    "Numer SIM",
    "IMEI",
    "Numer seryjny",
    "Rodzaj",
    "Osoba użytkująca",
    "Osoba odpowiedzialna",
    "Protokół zdawczo-odbiorczy",
    "Wyposażenie dodatkowe",
    "Stan baterii",
    "Czy uszkodzony",
)
# Numery kolumn (1-based) wyśrodkowanych: ID, nr tel., SIM, IMEI, protokół, uszkodzony.
_PHONES_CENTERED_COLUMNS = frozenset({1, 3, 4, 5, 10, 13})

_HISTORY_SHEET_HEADERS: tuple[str, ...] = (
    "Data",
    "Model telefonu",
    "Nr telefonu",
    "Kategoria",
    "Opis zdarzenia",
    "Uwagi",
)
# Wyśrodkowane: data, nr tel., kategoria.
_HISTORY_CENTERED_COLUMNS = frozenset({1, 3, 4})


# ---------------------------------------------------------------------------
# Połączenie z bazą
# ---------------------------------------------------------------------------


def get_db_path() -> str:
    """Zwraca ścieżkę do pliku bazy, tworząc katalog aplikacji w razie potrzeby."""
    appdata = os.getenv("APPDATA")
    app_dir = Path(appdata) / APP_DIR_NAME if appdata else Path.home() / FALLBACK_APP_DIR_NAME
    app_dir.mkdir(parents=True, exist_ok=True)
    return str(app_dir / DB_FILENAME)


DB_NAME = get_db_path()
KEYSTORE_PATH = str(Path(DB_NAME).with_name(KEYSTORE_FILENAME))


# ---------------------------------------------------------------------------
# Szyfrowanie (SQLCipher)
# ---------------------------------------------------------------------------

# Klucz główny bazy - ustawiany po zalogowaniu (``unlock``), trzymany tylko w pamięci.
_db_key: bytes | None = None


def _apply_key(conn: sqlite3.Connection, key: bytes) -> None:
    """Ustawia klucz szyfrowania; musi to być pierwsze polecenie na połączeniu.

    PRAGMA nie przyjmuje parametrów ``?``. Klucz jest podawany jako liczba
    szesnastkowa (``x'...'``), więc nie da się w nim przemycić poleceń SQL,
    a SQLCipher pomija swoje (wolne) wyprowadzanie klucza z hasła.
    """
    conn.execute(f"PRAGMA key = \"x'{key.hex()}'\"")


def _connect_encrypted(path: str | Path, key: bytes) -> sqlite3.Connection:
    conn = sqlite3.connect(str(path))
    _apply_key(conn, key)
    return conn


def _verify_readable(conn: sqlite3.Connection) -> None:
    """Zgłasza ``DatabaseError``, jeśli klucz nie pasuje do pliku."""
    conn.execute("SELECT COUNT(*) FROM sqlite_master").fetchone()


def unlock(key: bytes) -> None:
    """Odblokowuje bazę kluczem głównym (po zalogowaniu).

    :raises RuntimeError: gdy biblioteka nie obsługuje szyfrowania.
    :raises DatabaseError: gdy klucz nie pasuje do pliku bazy.
    """
    global _db_key
    with closing(_connect_encrypted(DB_NAME, key)) as conn:
        if not conn.execute("PRAGMA cipher_version").fetchone():
            raise RuntimeError("Brak obsługi SQLCipher - baza nie byłaby szyfrowana.")
        _verify_readable(conn)
    _db_key = key


def lock() -> None:
    """Usuwa klucz z pamięci (np. przy wylogowaniu)."""
    global _db_key
    _db_key = None


def is_unlocked() -> bool:
    return _db_key is not None


def database_exists(path: str | Path = DB_NAME) -> bool:
    file = Path(path)
    return file.exists() and file.stat().st_size > 0


def is_plain_database(path: str | Path = DB_NAME) -> bool:
    """Sprawdza, czy plik jest zwykłą (niezaszyfrowaną) bazą SQLite."""
    if not database_exists(path):
        return False
    with open(path, "rb") as file:
        return file.read(len(_PLAIN_SQLITE_HEADER)) == _PLAIN_SQLITE_HEADER


# Odstępy (w sekundach) między ponownymi próbami operacji na zablokowanym pliku.
_LOCK_RETRY_DELAYS = (0.1, 0.25, 0.5, 1.0, 2.0)


def _retry_if_locked(operation: Callable[..., Any], *args: Any) -> Any:
    """Wykonuje operację na pliku, ponawiając ją, gdy plik jest chwilowo zablokowany.

    W systemie Windows świeżo zapisany plik bywa przez chwilę otwarty przez
    antywirusa lub usługę indeksowania - wtedy przeniesienie lub usunięcie
    kończy się błędem ``[WinError 32]`` (``PermissionError``).
    """
    for delay in _LOCK_RETRY_DELAYS:
        try:
            return operation(*args)
        except PermissionError:
            time.sleep(delay)
    return operation(*args)  # ostatnia próba - błąd trafia do wywołującego


def _remove_quietly(path: str | Path) -> None:
    """Usuwa plik tymczasowy; błąd jest tylko logowany, aby nie przykrył właściwego błędu."""
    try:
        _retry_if_locked(Path(path).unlink, True)
    except OSError:
        logger.warning("Nie udało się usunąć pliku tymczasowego %s", path, exc_info=True)


def _table_counts(conn: sqlite3.Connection) -> dict[str, int]:
    tables = [
        row[0]
        for row in conn.execute(
            "SELECT name FROM sqlite_master WHERE type = 'table' AND name NOT LIKE 'sqlite_%'"
        )
    ]
    return {table: conn.execute(f'SELECT COUNT(*) FROM "{table}"').fetchone()[0] for table in tables}


def _export_plain_to_encrypted(plain_path: str | Path, encrypted_path: str | Path, key: bytes) -> None:
    """Kopiuje niezaszyfrowaną bazę do nowego, zaszyfrowanego pliku i weryfikuje wynik."""
    _retry_if_locked(Path(encrypted_path).unlink, True)  # pozostałość po nieudanej próbie
    with closing(sqlite3.connect(str(plain_path))) as plain:
        plain.execute(
            f"ATTACH DATABASE ? AS encrypted KEY \"x'{key.hex()}'\"", (str(encrypted_path),)
        )
        plain.execute("SELECT sqlcipher_export('encrypted')")
        plain.execute("DETACH DATABASE encrypted")
        expected = _table_counts(plain)

    with closing(_connect_encrypted(encrypted_path, key)) as encrypted:
        if _table_counts(encrypted) != expected:
            raise DatabaseError("Weryfikacja zaszyfrowanej kopii nie powiodła się.")


def encrypt_existing_database(
    key: bytes, before_replace: Callable[[], None] | None = None
) -> None:
    """Jednorazowo szyfruje dotychczasową, niezaszyfrowaną bazę aplikacji.

    Dane są najpierw kopiowane do pliku tymczasowego i sprawdzane (liczba
    wierszy w każdej tabeli); dopiero potem plik zastępuje oryginał.

    :param before_replace: wywoływana po udanym przygotowaniu zaszyfrowanej kopii,
        a przed podmianą pliku - służy do zapisania klucza, zanim baza stanie się
        zaszyfrowana (dzięki temu klucz nigdy nie zostanie utracony).
    """
    temporary = Path(DB_NAME).with_suffix(".encrypting.tmp")
    try:
        _export_plain_to_encrypted(DB_NAME, temporary, key)
        if before_replace is not None:
            before_replace()
        _retry_if_locked(os.replace, temporary, DB_NAME)
    except BaseException:
        # Oryginalna baza pozostaje nietknięta; usuwamy tylko plik tymczasowy.
        _remove_quietly(temporary)
        raise


def get_connection() -> sqlite3.Connection:
    """Otwiera nowe połączenie z zaszyfrowaną bazą aplikacji.

    :raises RuntimeError: gdy baza nie została odblokowana (brak logowania).
    """
    if _db_key is None:
        raise RuntimeError("Baza jest zablokowana - najpierw się zaloguj.")
    return _connect_encrypted(DB_NAME, _db_key)


@contextmanager
def _cursor() -> Iterator[sqlite3.Cursor]:
    """Udostępnia kursor w obrębie jednej transakcji.

    Transakcja jest zatwierdzana po poprawnym wyjściu z bloku, a wycofywana
    w przypadku wyjątku. Połączenie jest zawsze zamykane.
    """
    with closing(get_connection()) as conn, conn:
        yield conn.cursor()


# ---------------------------------------------------------------------------
# Funkcje pomocnicze
# ---------------------------------------------------------------------------


def _now() -> str:
    """Zwraca bieżący czas lokalny w formacie używanym w historii."""
    return datetime.now().strftime(DATETIME_FORMAT)  # noqa: DTZ005


def _to_flag(value: Any) -> int:
    """Konwertuje wartość (bool, liczbę lub tekst typu „Tak”) na 0/1."""
    if isinstance(value, str):
        return int(value.strip().lower() in TRUTHY_STRINGS)
    return int(bool(value))


def _clean_text(value: Any) -> str:
    """Zamienia wartość na tekst bez skrajnych białych znaków (None -> "")."""
    return "" if value is None else str(value).strip()


def _display_value(field: str, value: Any) -> str:
    """Zwraca czytelną reprezentację wartości pola na potrzeby historii."""
    if field in FLAG_FIELDS:
        return "Tak" if _to_flag(value) else "Nie"
    return _clean_text(value)


def _phone_values(data: Mapping[str, Any]) -> tuple[Any, ...]:
    """Buduje krotkę wartości w kolejności ``PHONE_FIELDS``.

    Pola wymagane muszą występować w ``data``; opcjonalne otrzymują wartości domyślne.
    """
    return tuple(
        data.get(field, OPTIONAL_FIELD_DEFAULTS[field])
        if field in OPTIONAL_FIELD_DEFAULTS
        else data[field]
        for field in PHONE_FIELDS
    )


def _insert_history(
    cursor: sqlite3.Cursor,
    phone_id: int,
    category: str,
    description: str,
    note: str = "",
    date: str | None = None,
) -> None:
    """Dodaje wpis do historii zdarzeń telefonu."""
    cursor.execute(
        _INSERT_HISTORY_SQL,
        (phone_id, date or _now(), category, description, note),
    )


def _placeholders(count: int) -> str:
    """Zwraca listę znaczników ``?`` dla klauzuli ``IN``."""
    return ", ".join("?" for _ in range(count))


def format_phone_number(val: Any) -> str:
    """Formatuje numer telefonu do postaci ``XXX-XXX-XXX`` lub ``+48 XXX-XXX-XXX``.

    Numery w nierozpoznanym formacie są zwracane bez zmian (po przycięciu spacji).
    """
    if not val:
        return ""

    val_str = str(val).strip()
    digits = re.sub(r"\D", "", val_str)

    if len(digits) == 9:
        return f"{digits[:3]}-{digits[3:6]}-{digits[6:]}"
    if len(digits) == 11 and digits.startswith("48"):
        local = digits[2:]
        return f"+48 {local[:3]}-{local[3:6]}-{local[6:]}"
    return val_str


# ---------------------------------------------------------------------------
# Inicjalizacja bazy
# ---------------------------------------------------------------------------


def _ensure_column(cursor: sqlite3.Cursor, table: str, column: str, definition: str) -> None:
    """Dodaje kolumnę do tabeli, jeśli jeszcze nie istnieje.

    Nazwy tabel i kolumn pochodzą wyłącznie ze stałych modułu (``_MIGRATIONS``).
    """
    existing_columns = {row[1] for row in cursor.execute(f"PRAGMA table_info({table})")}
    if column not in existing_columns:
        cursor.execute(f"ALTER TABLE {table} ADD COLUMN {column} {definition}")


def _seed_sample_data(cursor: sqlite3.Cursor) -> None:
    """Wstawia przykładowy rekord do pustej bazy."""
    sample = {
        "model": "Przykładowy Model",
        "nr_tel": format_phone_number("123456789"),
        "nr_sim": "SIM123456",
        "imei": "IMEI123456789",
        "nr_seryjny": "SERIAL123456",
        "rodzaj": "Rodzaj1",
        "osoba_uzytkujaca": "Użytkownik1",
        "osoba_odpowiedzialna": "Odpowiedzialny1",
    }
    cursor.execute(_INSERT_PHONE_SQL, _phone_values(sample))
    _insert_history(cursor, cursor.lastrowid, "Dodanie", "Telefon dodany do bazy danych")


def init_db() -> None:
    """Tworzy brakujące tabele, wykonuje migracje i zasila pustą bazę danymi przykładowymi."""
    with _cursor() as cursor:
        for statement in _SCHEMA:
            cursor.execute(statement)

        for table, column, definition in _MIGRATIONS:
            _ensure_column(cursor, table, column, definition)

        cursor.execute("SELECT COUNT(*) FROM telefony")
        if cursor.fetchone()[0] == 0:
            _seed_sample_data(cursor)


# ---------------------------------------------------------------------------
# Odczyt danych
# ---------------------------------------------------------------------------


def search_phones(query_str: str, show_deleted: bool = False) -> list[tuple]:
    """Wyszukuje telefony po numerze, użytkowniku, modelu lub IMEI.

    Zwraca krotki ``(id, nr_tel, osoba_uzytkujaca, model, czy_uszkodzony)``.
    """
    pattern = f"%{query_str}%"
    digits = re.sub(r"\D", "", query_str)
    digits_pattern = f"%{digits}%" if digits else pattern

    with _cursor() as cursor:
        cursor.execute(
            """
            SELECT id, nr_tel, osoba_uzytkujaca, model, COALESCE(czy_uszkodzony, 0)
            FROM telefony
            WHERE (
                   nr_tel LIKE ?
                OR REPLACE(nr_tel, '-', '') LIKE ?
                OR osoba_uzytkujaca LIKE ?
                OR model LIKE ?
                OR imei LIKE ?
            )
            AND czy_usuniety = ?
            ORDER BY id DESC
            """,
            (pattern, digits_pattern, pattern, pattern, pattern, int(show_deleted)),
        )
        return cursor.fetchall()


def get_phone_by_id(phone_id: int) -> tuple | None:
    """Zwraca dane telefonu w kolejności ``PHONE_FIELDS`` lub ``None``."""
    with _cursor() as cursor:
        cursor.execute(f"SELECT {_PHONE_COLUMNS_SQL} FROM telefony WHERE id = ?", (phone_id,))
        return cursor.fetchone()


def get_phone_history(phone_id: int) -> list[tuple]:
    """Zwraca historię telefonu: ``(id, data, kategoria, opis, uwagi)``, od najnowszych."""
    with _cursor() as cursor:
        cursor.execute(
            """
            SELECT id, data, kategoria, opis, COALESCE(uwagi, '')
            FROM historia
            WHERE telefon_id = ?
            ORDER BY id DESC
            """,
            (phone_id,),
        )
        return cursor.fetchall()


def get_phones_count(show_deleted: bool = False) -> int:
    """Zwraca liczbę telefonów aktywnych lub znajdujących się w koszu."""
    with _cursor() as cursor:
        cursor.execute(
            "SELECT COUNT(*) FROM telefony WHERE czy_usuniety = ?",
            (int(show_deleted),),
        )
        return cursor.fetchone()[0]


# ---------------------------------------------------------------------------
# Zapis danych
# ---------------------------------------------------------------------------


def insert_phone(data: Mapping[str, Any]) -> int:
    """Dodaje nowy telefon wraz z wpisem w historii i zwraca jego identyfikator."""
    record = {**data, "nr_tel": format_phone_number(data.get("nr_tel", ""))}

    with _cursor() as cursor:
        cursor.execute(_INSERT_PHONE_SQL, _phone_values(record))
        phone_id = cursor.lastrowid
        _insert_history(
            cursor, phone_id, "Dodanie do bazy", "Zarejestrowano urządzenie w systemie."
        )
        return phone_id


def update_phone(phone_id: int, data: Mapping[str, Any]) -> None:
    """Aktualizuje dane telefonu i rejestruje w historii każdą zmienioną wartość."""
    record = {**data, "nr_tel": format_phone_number(data.get("nr_tel", ""))}
    new_values = _phone_values(record)

    with _cursor() as cursor:
        cursor.execute(f"SELECT {_PHONE_COLUMNS_SQL} FROM telefony WHERE id = ?", (phone_id,))
        old_values = cursor.fetchone()

        cursor.execute(_UPDATE_PHONE_SQL, (*new_values, phone_id))

        if old_values is None:
            return

        timestamp = _now()
        for field, old_raw, new_raw in zip(PHONE_FIELDS, old_values, new_values):
            old_display = _display_value(field, old_raw)
            new_display = _display_value(field, new_raw)
            if old_display == new_display:
                continue

            description = (
                f"Zmieniono {FIELD_LABELS[field]}: "
                f"'{old_display or '[puste]'}' ➔ '{new_display or '[puste]'}'"
            )
            _insert_history(
                cursor,
                phone_id,
                CHANGE_CATEGORIES.get(field, DEFAULT_CHANGE_CATEGORY),
                description,
                date=timestamp,
            )


def update_history_note(history_id: int, note: str) -> None:
    """Ustawia uwagi dla wskazanego wpisu historii."""
    with _cursor() as cursor:
        cursor.execute("UPDATE historia SET uwagi = ? WHERE id = ?", (note, history_id))


def add_history_entry(phone_id: int, category: str, description: str, note: str = "") -> None:
    """Dodaje ręczny wpis do historii telefonu."""
    with _cursor() as cursor:
        _insert_history(cursor, phone_id, category, description, note)


def soft_delete_phone(phone_id: int) -> None:
    """Przenosi telefon do kosza (usuwa go z listy aktywnych)."""
    with _cursor() as cursor:
        cursor.execute("UPDATE telefony SET czy_usuniety = 1 WHERE id = ?", (phone_id,))
        _insert_history(
            cursor,
            phone_id,
            "Kosz / Usunięcie",
            "Telefon przeniesiono do kosza (usunięto z aktywnej listy).",
        )


def restore_phone(phone_id: int) -> None:
    """Przywraca telefon z kosza do listy aktywnych."""
    with _cursor() as cursor:
        cursor.execute("UPDATE telefony SET czy_usuniety = 0 WHERE id = ?", (phone_id,))
        _insert_history(
            cursor, phone_id, "Przywrócenie", "Przywrócono urządzenie z kosza do aktywnych."
        )


def hard_delete_phone(phone_id: int) -> None:
    """Trwale usuwa telefon wraz z całą jego historią."""
    with _cursor() as cursor:
        cursor.execute("DELETE FROM historia WHERE telefon_id = ?", (phone_id,))
        cursor.execute("DELETE FROM telefony WHERE id = ?", (phone_id,))


def toggle_phone_warning(phone_id: int) -> int:
    """Przełącza flagę „uszkodzony” i zwraca jej nową wartość (0/1)."""
    with _cursor() as cursor:
        cursor.execute("SELECT COALESCE(czy_uszkodzony, 0) FROM telefony WHERE id = ?", (phone_id,))
        row = cursor.fetchone()
        if row is None:
            return 0

        new_value = 0 if row[0] == 1 else 1
        cursor.execute(
            "UPDATE telefony SET czy_uszkodzony = ? WHERE id = ?", (new_value, phone_id)
        )
        return new_value


# ---------------------------------------------------------------------------
# Import masowy
# ---------------------------------------------------------------------------


def _normalize_import_record(data: Mapping[str, Any]) -> dict[str, Any] | None:
    """Oczyszcza rekord importu; zwraca ``None``, gdy brakuje modelu lub numeru."""
    model = _clean_text(data.get("model"))
    nr_tel = format_phone_number(_clean_text(data.get("nr_tel")))
    if not model or not nr_tel:
        return None

    return {
        "model": model,
        "nr_tel": nr_tel,
        "nr_sim": _clean_text(data.get("nr_sim")),
        "imei": _clean_text(data.get("imei")),
        "nr_seryjny": _clean_text(data.get("nr_seryjny")),
        "rodzaj": _clean_text(data.get("rodzaj")),
        "osoba_uzytkujaca": _clean_text(data.get("osoba_uzytkujaca")),
        "osoba_odpowiedzialna": _clean_text(data.get("osoba_odpowiedzialna")),
        "czy_protokol": _to_flag(data.get("czy_protokol", 0)),
        "wyposazenie": _clean_text(data.get("wyposazenie")),
        "stan_baterii": _clean_text(data.get("stan_baterii")) or DEFAULT_BATTERY_STATE,
        "czy_uszkodzony": _to_flag(data.get("czy_uszkodzony", 0)),
    }


def bulk_insert_phones(
    phone_records: Iterable[Mapping[str, Any]],
    history_records: Sequence[Mapping[str, Any]] | None = None,
) -> tuple[int, int]:
    """Importuje telefony (i opcjonalnie ich historię) w jednej transakcji.

    Rekordy bez modelu lub numeru telefonu są pomijane. Jeśli nie podano
    historii, każdy telefon otrzymuje wpis „Dodanie do bazy”. Wpisy historii
    są przypisywane do telefonów po numerze telefonu.

    Zwraca krotkę ``(liczba_telefonów, liczba_wpisów_historii)``.
    """
    timestamp = _now()
    phones_inserted = 0
    history_inserted = 0
    phone_id_by_number: dict[str, int] = {}

    with _cursor() as cursor:
        for raw_record in phone_records:
            record = _normalize_import_record(raw_record)
            if record is None:
                continue

            cursor.execute(_INSERT_PHONE_SQL, _phone_values(record))
            phone_id = cursor.lastrowid
            phone_id_by_number[record["nr_tel"]] = phone_id
            phones_inserted += 1

            if not history_records:
                _insert_history(
                    cursor,
                    phone_id,
                    "Dodanie do bazy",
                    "Zarejestrowano urządzenie w systemie.",
                    date=timestamp,
                )

        for entry in history_records or ():
            nr_tel = format_phone_number(entry.get("nr_tel", ""))
            if not nr_tel:
                continue

            phone_id = phone_id_by_number.get(nr_tel)
            if phone_id is None:
                cursor.execute("SELECT id FROM telefony WHERE nr_tel = ? LIMIT 1", (nr_tel,))
                row = cursor.fetchone()
                if row is None:
                    continue
                phone_id = row[0]

            _insert_history(
                cursor,
                phone_id,
                entry.get("kategoria") or "Import",
                entry.get("opis") or "",
                entry.get("uwagi") or "",
                date=entry.get("data") or timestamp,
            )
            history_inserted += 1

    return phones_inserted, history_inserted


# ---------------------------------------------------------------------------
# Eksport do Excela
# ---------------------------------------------------------------------------


def _style_sheet(
    ws: Worksheet,
    header_fill: PatternFill,
    column_count: int,
    centered_columns: frozenset[int],
) -> None:
    """Formatuje nagłówek i komórki danych, zamraża nagłówek i włącza filtr."""
    for cell in ws[1]:
        cell.font = _HEADER_FONT
        cell.fill = header_fill
        cell.alignment = _ALIGN_CENTER

    for row in ws.iter_rows(min_row=2, max_row=ws.max_row, max_col=column_count):
        for cell in row:
            cell.border = _CELL_BORDER
            cell.alignment = _ALIGN_CENTER if cell.column in centered_columns else _ALIGN_LEFT

    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions


def _autofit_columns(ws: Worksheet) -> None:
    """Dopasowuje szerokość kolumn do najdłuższej wartości."""
    for column in ws.columns:
        longest = max((len(str(cell.value)) for cell in column if cell.value is not None), default=0)
        letter = get_column_letter(column[0].column)
        ws.column_dimensions[letter].width = max(longest + _COLUMN_PADDING, _MIN_COLUMN_WIDTH)


def _fetch_export_data(phone_ids: list[int] | None) -> tuple[list[tuple], list[tuple]]:
    """Pobiera telefony i historię do eksportu (wszystkie lub wybrane)."""
    phones_sql = f"SELECT id, {_PHONE_COLUMNS_SQL} FROM telefony"
    history_sql = """
        SELECT h.data, t.model, t.nr_tel, h.kategoria, h.opis, COALESCE(h.uwagi, '')
        FROM historia h
        LEFT JOIN telefony t ON h.telefon_id = t.id
    """
    params: tuple[int, ...] = ()

    if phone_ids is not None:
        in_clause = _placeholders(len(phone_ids))
        phones_sql += f" WHERE id IN ({in_clause})"
        history_sql += f" WHERE h.telefon_id IN ({in_clause})"
        params = tuple(phone_ids)

    with _cursor() as cursor:
        cursor.execute(f"{phones_sql} ORDER BY id ASC", params)
        phone_rows = cursor.fetchall()
        cursor.execute(f"{history_sql} ORDER BY h.id DESC", params)
        history_rows = cursor.fetchall()

    return phone_rows, history_rows


def export_to_excel(file_path: str | Path, phone_ids: Iterable[int] | None = None) -> int:
    """Eksportuje telefony i ich historię do pliku ``.xlsx``.

    :param file_path: ścieżka pliku docelowego.
    :param phone_ids: identyfikatory telefonów do eksportu; ``None`` oznacza wszystkie.
    :return: liczba wyeksportowanych telefonów (0, gdy przekazano pustą listę;
             wtedy plik nie jest tworzony).
    """
    ids = list(phone_ids) if phone_ids is not None else None
    if ids is not None and not ids:
        return 0

    phone_rows, history_rows = _fetch_export_data(ids)

    wb = openpyxl.Workbook()

    ws_phones = wb.active
    ws_phones.title = "Telefony"
    ws_phones.append(_PHONES_SHEET_HEADERS)
    flag_indexes = [
        index for index, field in enumerate(("id", *PHONE_FIELDS)) if field in FLAG_FIELDS
    ]
    for row in phone_rows:
        values = list(row)
        for index in flag_indexes:
            values[index] = "Tak" if values[index] else "Nie"
        ws_phones.append(values)
    _style_sheet(
        ws_phones, _PHONES_HEADER_FILL, len(_PHONES_SHEET_HEADERS), _PHONES_CENTERED_COLUMNS
    )

    ws_history = wb.create_sheet(title="Historia zdarzeń")
    ws_history.append(_HISTORY_SHEET_HEADERS)
    for row in history_rows:
        ws_history.append(list(row))
    _style_sheet(
        ws_history, _HISTORY_HEADER_FILL, len(_HISTORY_SHEET_HEADERS), _HISTORY_CENTERED_COLUMNS
    )

    for ws in (ws_phones, ws_history):
        _autofit_columns(ws)

    wb.save(file_path)
    return len(phone_rows)


# ---------------------------------------------------------------------------
# Dane pracodawcy
# ---------------------------------------------------------------------------


def get_employer() -> tuple[str, str, str] | None:
    """Zwraca ``(nazwa_firmy, adres, nip)`` lub ``None``, jeśli nie zapisano danych."""
    with _cursor() as cursor:
        cursor.execute("SELECT nazwa_firmy, adres, nip FROM pracodawca ORDER BY id DESC LIMIT 1")
        return cursor.fetchone()


def save_employer(nazwa_firmy: str, adres: str, nip: str) -> None:
    """Zapisuje (lub aktualizuje) dane pracodawcy."""
    with _cursor() as cursor:
        cursor.execute("SELECT id FROM pracodawca ORDER BY id DESC LIMIT 1")
        row = cursor.fetchone()
        if row is None:
            cursor.execute(
                "INSERT INTO pracodawca (nazwa_firmy, adres, nip) VALUES (?, ?, ?)",
                (nazwa_firmy, adres, nip),
            )
        else:
            cursor.execute(
                "UPDATE pracodawca SET nazwa_firmy = ?, adres = ?, nip = ? WHERE id = ?",
                (nazwa_firmy, adres, nip, row[0]),
            )


# ---------------------------------------------------------------------------
# Kopie zapasowe
# ---------------------------------------------------------------------------


def backup_keys_path(backup_path: str | Path) -> Path:
    """Ścieżka kopii pliku kluczy zapisywanej obok kopii bazy."""
    path = Path(backup_path)
    return path.with_name(path.stem + BACKUP_KEYS_SUFFIX)


def backup_database(dest_path: str | Path) -> Path:
    """Tworzy zaszyfrowaną kopię bazy oraz kopię pliku kluczy obok niej.

    Kopia jest szyfrowana tym samym kluczem głównym co baza, więc da się ją
    przywrócić także po zmianie haseł. Na nowym komputerze do jej otwarcia
    potrzebny jest skopiowany plik kluczy.

    :return: ścieżka zapisanej kopii pliku kluczy.
    """
    Path(dest_path).unlink(missing_ok=True)
    with closing(get_connection()) as source, closing(
        _connect_encrypted(dest_path, _db_key)
    ) as target:
        source.backup(target)

    keys_copy = backup_keys_path(dest_path)
    shutil.copyfile(KEYSTORE_PATH, keys_copy)
    return keys_copy


def restore_database(backup_path: str | Path) -> None:
    """Przywraca bazę z kopii zapasowej i uzupełnia schemat do bieżącej wersji.

    Obsługuje kopie zaszyfrowane tym samym kluczem oraz stare kopie
    niezaszyfrowane (sprzed wprowadzenia szyfrowania).

    :raises DatabaseError: gdy plik nie jest bazą lub zaszyfrowano go innym kluczem.
    """
    if _db_key is None:
        raise RuntimeError("Baza jest zablokowana - najpierw się zaloguj.")

    temporary: Path | None = None
    source_path: str | Path = backup_path
    if is_plain_database(backup_path):
        # API kopii nie kopiuje między bazą jawną a zaszyfrowaną - najpierw ją szyfrujemy.
        temporary = Path(DB_NAME).with_suffix(".restore.tmp")
        _export_plain_to_encrypted(backup_path, temporary, _db_key)
        source_path = temporary

    try:
        with closing(_connect_encrypted(source_path, _db_key)) as source:
            _verify_readable(source)
            with closing(get_connection()) as target:
                source.backup(target)
    finally:
        if temporary is not None:
            _remove_quietly(temporary)

    init_db()
