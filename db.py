from __future__ import annotations

import logging
import os
import re
import shutil
import time
from collections.abc import Callable, Iterable, Iterator, Mapping, Sequence
from contextlib import closing, contextmanager
from dataclasses import dataclass
from dataclasses import field as dataclass_field
from datetime import datetime
from pathlib import Path
from typing import Any

import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.worksheet import Worksheet
from sqlcipher3 import dbapi2 as sqlite3

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

_PLAIN_SQLITE_HEADER = b"SQLite format 3\x00"

DATETIME_FORMAT = "%Y-%m-%d %H:%M"
DEFAULT_BATTERY_STATE = "Brak informacji"
TRUTHY_STRINGS = frozenset({"tak", "1", "true", "t", "x"})

_PHONE_LIKE_QUERY = re.compile(r"[\d\s()+\-]+")
_SEARCH_TEXT_COLUMNS = ("nr_tel", "osoba_uzytkujaca", "model", "imei")

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

FLAG_FIELDS = frozenset({"czy_protokol", "czy_uszkodzony"})

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
_PHONES_CENTERED_COLUMNS = frozenset({1, 3, 4, 5, 10, 13})

_HISTORY_SHEET_HEADERS: tuple[str, ...] = (
    "Data",
    "Model telefonu",
    "Nr telefonu",
    "Kategoria",
    "Opis zdarzenia",
    "Uwagi",
)
_HISTORY_CENTERED_COLUMNS = frozenset({1, 3, 4})


# ---------------------------------------------------------------------------
# Połączenie z bazą
# ---------------------------------------------------------------------------


def get_db_path() -> str:
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
    conn.execute(f"PRAGMA key = \"x'{key.hex()}'\"")


def _connect_encrypted(path: str | Path, key: bytes) -> sqlite3.Connection:
    conn = sqlite3.connect(str(path))
    _apply_key(conn, key)
    return conn


def _verify_readable(conn: sqlite3.Connection) -> None:
    conn.execute("SELECT COUNT(*) FROM sqlite_master").fetchone()


def unlock(key: bytes) -> None:
    global _db_key
    with closing(_connect_encrypted(DB_NAME, key)) as conn:
        if not conn.execute("PRAGMA cipher_version").fetchone():
            raise RuntimeError("Brak obsługi SQLCipher - baza nie byłaby szyfrowana.")
        _verify_readable(conn)
    _db_key = key


def lock() -> None:
    global _db_key
    _db_key = None


def is_unlocked() -> bool:
    return _db_key is not None


def database_exists(path: str | Path = DB_NAME) -> bool:
    file = Path(path)
    return file.exists() and file.stat().st_size > 0


def is_plain_database(path: str | Path = DB_NAME) -> bool:
    if not database_exists(path):
        return False
    with open(path, "rb") as file:
        return file.read(len(_PLAIN_SQLITE_HEADER)) == _PLAIN_SQLITE_HEADER


_LOCK_RETRY_DELAYS = (0.1, 0.25, 0.5, 1.0, 2.0)


def _retry_if_locked(operation: Callable[..., Any], *args: Any) -> Any:
    for delay in _LOCK_RETRY_DELAYS:
        try:
            return operation(*args)
        except PermissionError:
            time.sleep(delay)
    return operation(*args)  


def _remove_quietly(path: str | Path) -> None:
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
    _retry_if_locked(Path(encrypted_path).unlink, True)  
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
    temporary = Path(DB_NAME).with_suffix(".encrypting.tmp")
    try:
        _export_plain_to_encrypted(DB_NAME, temporary, key)
        if before_replace is not None:
            before_replace()
        _retry_if_locked(os.replace, temporary, DB_NAME)
    except BaseException:
        _remove_quietly(temporary)
        raise


def get_connection() -> sqlite3.Connection:
    if _db_key is None:
        raise RuntimeError("Baza jest zablokowana - najpierw się zaloguj.")
    return _connect_encrypted(DB_NAME, _db_key)


@contextmanager
def _cursor() -> Iterator[sqlite3.Cursor]:
    with closing(get_connection()) as conn, conn:
        yield conn.cursor()


# ---------------------------------------------------------------------------
# Funkcje pomocnicze
# ---------------------------------------------------------------------------


def _now() -> str:
    return datetime.now().strftime(DATETIME_FORMAT)  # noqa: DTZ005


def _to_flag(value: Any) -> int:
    if isinstance(value, str):
        return int(value.strip().lower() in TRUTHY_STRINGS)
    return int(bool(value))


def _clean_text(value: Any) -> str:
    return "" if value is None else str(value).strip()


def _display_value(field: str, value: Any) -> str:
    if field in FLAG_FIELDS:
        return "Tak" if _to_flag(value) else "Nie"
    return _clean_text(value)


def _phone_values(data: Mapping[str, Any]) -> tuple[Any, ...]:
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
    cursor.execute(
        _INSERT_HISTORY_SQL,
        (phone_id, date or _now(), category, description, note),
    )


def _placeholders(count: int) -> str:
    return ", ".join("?" for _ in range(count))


def format_phone_number(val: Any) -> str:
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
# Walidacja IMEI i wykrywanie duplikatów
# ---------------------------------------------------------------------------

IMEI_LENGTH = 15

def normalize_imei(value: Any) -> str:
    text = _clean_text(value)
    compact = re.sub(r"[\s-]", "", text)
    return compact if compact.isdigit() else text

def _luhn_checksum_valid(digits: str) -> bool:
    total = 0
    for position, char in enumerate(reversed(digits)):
        digit = int(char)
        if position % 2 == 1:
            digit *= 2
            if digit > 9:
                digit -= 9
        total += digit
    return total % 10 == 0

def imei_error(value: Any) -> str | None:
    imei = normalize_imei(value)
    if not imei:
        return None
    if not imei.isdigit():
        return "IMEI może zawierać tylko cyfry."
    if len(imei) != IMEI_LENGTH:
        return f"IMEI musi mieć {IMEI_LENGTH} cyfr (wpisano {len(imei)})."
    if not _luhn_checksum_valid(imei):
        return "Nieprawidłowa cyfra kontrolna IMEI - prawdopodobnie literówka."
    return None

def is_valid_imei(value: Any) -> bool:
    return imei_error(value) is None

def _phone_key(value: Any) -> str:
    digits = re.sub(r"\D", "", _clean_text(value))
    if len(digits) == 11 and digits.startswith("48"):
        digits= digits[2:]
    return digits

@dataclass(frozen=True)
class Duplicate:
    phone_id:int
    fields: tuple[str, ...]
    nr_tel: str
    model: str
    user: str
    in_trash: bool

    def describe(self) -> str:
        names = {"nr_tel": "numer telefonu", "imei": "IMEI"}
        what = "ten sam " + " i ".join(names[name] for name in self.fields)
        trash = ", w koszu" if self.in_trash else ""
        return (
            f"ID #{self.phone_id}: {self.model} ({self.nr_tel}), "
            f"użytkownik: {self.user or '[brak]'}{trash} - {what}"
        )

def _identifier_rows() -> list[tuple]:
    with _cursor() as cursor:
        cursor.execute(
            "SELECT id, nr_tel, imei, model, osoba_uzytkujaca, COALESCE(czy_usuniety, 0) "
            "FROM telefony"
        )
        return cursor.fetchall()

def find_duplicates(nr_tel: Any, imei: Any, exclude_id: int | None = None) -> list[Duplicate]:
    phone_key = _phone_key(nr_tel)
    imei_key = normalize_imei(imei)
    if not phone_key and not imei_key:
        return[]

    duplicates = []
    for phone_id, row_nr, row_imei, model, user, deleted in _identifier_rows():
        if phone_id == exclude_id:
            continue
        matching = []
        if phone_key and _phone_key(row_nr) == phone_key:
            matching.append("nr_tel")
        if imei_key and normalize_imei(row_imei) == imei_key:
            matching.append("imei")
        if matching:
            duplicates.append(
                Duplicate(
                    phone_id=phone_id,
                    fields=tuple(matching),
                    nr_tel=row_nr,
                    model=model,
                    user=user or "",
                    in_trash=bool(deleted),
                )
            )
    return duplicates

# ---------------------------------------------------------------------------
# Inicjalizacja bazy
# ---------------------------------------------------------------------------


def _ensure_column(cursor: sqlite3.Cursor, table: str, column: str, definition: str) -> None:
    existing_columns = {row[1] for row in cursor.execute(f"PRAGMA table_info({table})")}
    if column not in existing_columns:
        cursor.execute(f"ALTER TABLE {table} ADD COLUMN {column} {definition}")


def _seed_sample_data(cursor: sqlite3.Cursor) -> None:
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

def _sql_lower(value: Any) -> str:
    return "" if value is None else str(value).lower()


def _sql_digits(value: Any) -> str:
    return "" if value is None else re.sub(r"\D", "", str(value))


def _like_pattern(text: str) -> str:
    escaped = text.replace("\\", "\\\\").replace("%", "\\%").replace("_", "\\_")
    return f"%{escaped}%"


def search_phones(query_str: str, show_deleted: bool = False, only_damaged: bool = False) -> list[tuple]:
    query = query_str.strip()
    conditions: list[str] = []
    params: list[Any] = []

    if query:
        text_pattern = _like_pattern(query.lower())
        for column in _SEARCH_TEXT_COLUMNS:
            conditions.append(f"PL_LOWER({column}) LIKE ? ESCAPE '\\'")
            params.append(text_pattern)

        digits = _sql_digits(query)
        if digits and _PHONE_LIKE_QUERY.fullmatch(query):
            variants = {digits}
            if digits.startswith("48") and (query.startswith("+") or len(digits) == 11):
                variants.add(digits[2:])
            for variant in sorted(v for v in variants if v):
                conditions.append("DIGITS(nr_tel) LIKE ?")
                params.append(f"%{variant}%")
            conditions.append("DIGITS(imei) LIKE ?")
            params.append(f"%{digits}%")

    where = f"({' OR '.join(conditions)}) AND " if conditions else ""

    with _cursor() as cursor:
        cursor.connection.create_function("PL_LOWER", 1, _sql_lower)
        cursor.connection.create_function("DIGITS", 1, _sql_digits)
        cursor.execute(
            f"""
            SELECT id, nr_tel, osoba_uzytkujaca, model, COALESCE(czy_uszkodzony, 0)
            FROM telefony
            WHERE {where}COALESCE(czy_usuniety, 0) = ?
            AND (? = 0 OR COALESCE(czy_uszkodzony, 0) = 1)
            ORDER BY id DESC
            """,
            (*params, int(show_deleted), int(only_damaged)),
        )
        return cursor.fetchall()


def get_phone_by_id(phone_id: int) -> tuple | None:
    with _cursor() as cursor:
        cursor.execute(f"SELECT {_PHONE_COLUMNS_SQL} FROM telefony WHERE id = ?", (phone_id,))
        return cursor.fetchone()


def get_phone_history(phone_id: int) -> list[tuple]:
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
    with _cursor() as cursor:
        cursor.execute(
            "SELECT COUNT(*) FROM telefony WHERE COALESCE(czy_usuniety, 0) = ?",
            (int(show_deleted),),
        )
        return cursor.fetchone()[0]


# ---------------------------------------------------------------------------
# Zapis danych
# ---------------------------------------------------------------------------


def insert_phone(data: Mapping[str, Any]) -> int:
    record = {**data, "nr_tel": format_phone_number(data.get("nr_tel", ""))}

    with _cursor() as cursor:
        cursor.execute(_INSERT_PHONE_SQL, _phone_values(record))
        phone_id = cursor.lastrowid
        _insert_history(
            cursor, phone_id, "Dodanie do bazy", "Zarejestrowano urządzenie w systemie."
        )
        return phone_id


def update_phone(phone_id: int, data: Mapping[str, Any]) -> None:
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
    with _cursor() as cursor:
        cursor.execute("UPDATE historia SET uwagi = ? WHERE id = ?", (note, history_id))


def add_history_entry(phone_id: int, category: str, description: str, note: str = "") -> None:
    with _cursor() as cursor:
        _insert_history(cursor, phone_id, category, description, note)


def soft_delete_phone(phone_id: int) -> None:
    with _cursor() as cursor:
        cursor.execute("UPDATE telefony SET czy_usuniety = 1 WHERE id = ?", (phone_id,))
        _insert_history(
            cursor,
            phone_id,
            "Kosz / Usunięcie",
            "Telefon przeniesiono do kosza (usunięto z aktywnej listy).",
        )


def restore_phone(phone_id: int) -> None:
    with _cursor() as cursor:
        cursor.execute("UPDATE telefony SET czy_usuniety = 0 WHERE id = ?", (phone_id,))
        _insert_history(
            cursor, phone_id, "Przywrócenie", "Przywrócono urządzenie z kosza do aktywnych."
        )


def hard_delete_phone(phone_id: int) -> None:
    with _cursor() as cursor:
        cursor.execute("DELETE FROM historia WHERE telefon_id = ?", (phone_id,))
        cursor.execute("DELETE FROM telefony WHERE id = ?", (phone_id,))


def toggle_phone_warning(phone_id: int) -> int:
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
    model = _clean_text(data.get("model"))
    nr_tel = format_phone_number(_clean_text(data.get("nr_tel")))
    if not model or not nr_tel:
        return None

    return {
        "model": model,
        "nr_tel": nr_tel,
        "nr_sim": _clean_text(data.get("nr_sim")),
        "imei": normalize_imei(data.get("imei")),
        "nr_seryjny": _clean_text(data.get("nr_seryjny")),
        "rodzaj": _clean_text(data.get("rodzaj")),
        "osoba_uzytkujaca": _clean_text(data.get("osoba_uzytkujaca")),
        "osoba_odpowiedzialna": _clean_text(data.get("osoba_odpowiedzialna")),
        "czy_protokol": _to_flag(data.get("czy_protokol", 0)),
        "wyposazenie": _clean_text(data.get("wyposazenie")),
        "stan_baterii": _clean_text(data.get("stan_baterii")) or DEFAULT_BATTERY_STATE,
        "czy_uszkodzony": _to_flag(data.get("czy_uszkodzony", 0)),
    }

@dataclass
class ImportResult:
    phones: int = 0
    history: int = 0
    skipped_duplicates: list[str] = dataclass_field(default_factory=list)
    invalid_imei: list[str] = dataclass_field(default_factory=list)

def bulk_insert_phones(
    phone_records: Iterable[Mapping[str, Any]],
    history_records: Sequence[Mapping[str, Any]] | None = None,
    skip_duplicates: bool = False,
) -> ImportResult:
    timestamp = _now()
    result = ImportResult()
    phone_id_by_number: dict[str, int] = {}
    skipped_phone_keys: set[str] = set()

    known_phones: set[str] = set()
    known_imeis: set[str] = set()
    if skip_duplicates:
        for _id, row_nr, row_imei, *_rest in _identifier_rows():
            known_phones.add(_phone_key(row_nr))
            if normalize_imei(row_imei):
                known_imeis.add(normalize_imei(row_imei))

    with _cursor() as cursor:
        for raw_record in phone_records:
            record = _normalize_import_record(raw_record)
            if record is None:
                continue

            phone_key = _phone_key(record["nr_tel"])
            imei = record["imei"]
            label = f"{record['nr_tel']} ({record['model']})"

            if skip_duplicates:
                if phone_key in known_phones or (imei and imei in known_imeis):
                    result.skipped_duplicates.append(label)
                    skipped_phone_keys.add(phone_key)
                    continue
                known_phones.add(phone_key)
                if imei:
                    known_imeis.add(imei)

            if imei_error(imei):
                result.invalid_imei.append(f"{label}: {imei}")

            cursor.execute(_INSERT_PHONE_SQL, _phone_values(record))
            phone_id = cursor.lastrowid
            phone_id_by_number[record["nr_tel"]] = phone_id
            result.phones += 1

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
            if not nr_tel or _phone_key(nr_tel) in skipped_phone_keys:
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
            result.history += 1

    return result


# ---------------------------------------------------------------------------
# Eksport do Excela
# ---------------------------------------------------------------------------


def _style_sheet(
    ws: Worksheet,
    header_fill: PatternFill,
    column_count: int,
    centered_columns: frozenset[int],
) -> None:
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


def _append_row(ws: Worksheet, values: Iterable[Any]) -> None:
    ws.append(list(values))
    for cell in ws[ws.max_row]:
        if isinstance(cell.value, str) and cell.value.startswith("="):
            cell.data_type = "s"


def _autofit_columns(ws: Worksheet) -> None:
    for column in ws.columns:
        longest = max((len(str(cell.value)) for cell in column if cell.value is not None), default=0)
        letter = get_column_letter(column[0].column)
        ws.column_dimensions[letter].width = max(longest + _COLUMN_PADDING, _MIN_COLUMN_WIDTH)


def _fetch_export_data(phone_ids: list[int] | None) -> tuple[list[tuple], list[tuple]]:
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
        _append_row(ws_phones, values)
    _style_sheet(
        ws_phones, _PHONES_HEADER_FILL, len(_PHONES_SHEET_HEADERS), _PHONES_CENTERED_COLUMNS
    )

    ws_history = wb.create_sheet(title="Historia zdarzeń")
    ws_history.append(_HISTORY_SHEET_HEADERS)
    for row in history_rows:
        _append_row(ws_history, row)
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
    with _cursor() as cursor:
        cursor.execute("SELECT nazwa_firmy, adres, nip FROM pracodawca ORDER BY id DESC LIMIT 1")
        return cursor.fetchone()


def save_employer(nazwa_firmy: str, adres: str, nip: str) -> None:
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
    path = Path(backup_path)
    return path.with_name(path.stem + BACKUP_KEYS_SUFFIX)


def backup_database(dest_path: str | Path) -> Path:
    Path(dest_path).unlink(missing_ok=True)
    with closing(get_connection()) as source, closing(
        _connect_encrypted(dest_path, _db_key)
    ) as target:
        source.backup(target)

    keys_copy = backup_keys_path(dest_path)
    shutil.copyfile(KEYSTORE_PATH, keys_copy)
    return keys_copy


def restore_database(backup_path: str | Path) -> None:
    if _db_key is None:
        raise RuntimeError("Baza jest zablokowana - najpierw się zaloguj.")

    temporary: Path | None = None
    source_path: str | Path = backup_path
    if is_plain_database(backup_path):
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
