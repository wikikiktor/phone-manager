from __future__ import annotations

import json
import logging
from dataclasses import asdict, dataclass, fields
from datetime import datetime, timedelta
from pathlib import Path
from typing import Optional

import db

logger = logging.getLogger(__name__)

SETTINGS_PATH = Path(db.DB_NAME).with_name("auto_backup.json")
FILE_PREFIX = "auto_telefony_"
FILE_TIMESTAMP_FORMAT = "%Y%m%d_%H%M%S"
LAST_BACKUP_FORMAT = "%Y-%m-%d %H:%M:%S"

MIN_INTERVAL_DAYS, MAX_INTERVAL_DAYS = 1, 365
MIN_KEEP_COUNT, MAX_KEEP_COUNT = 1, 100

def default_folder() -> str:
    return str(Path.home() / "Documents" / "Baza Telefonów - kopie")

@dataclass
class AutoBackupSettings:
    enabled: bool = False
    folder: str = ""
    interval_days: int = 7
    keep_count: int = 8
    last_backup: Optional[str] = None

    def __post_init__(self) -> None:
        if not self.folder:
            self.folder = default_folder()

    @property
    def last_backup_time(self) -> Optional[datetime]:
        if not self.last_backup:
            return None
        try:
            return datetime.strptime(self.last_backup, LAST_BACKUP_FORMAT)
        except ValueError:
            return None

    @property
    def next_backup_time(self) -> Optional[datetime]:
        last = self.last_backup_time
        return None if last is None else last + timedelta(days=self.interval_days)

    def is_due(self, now: Optional[datetime] = None) -> bool:
        if not self.enabled:
            return False
        next_time = self.next_backup_time
        return next_time is None or (now or datetime.now()) >= next_time

def validate(interval_days: int, keep_count: int, folder: str) -> list[str]:
    problems = []
    if not MIN_INTERVAL_DAYS <= interval_days <= MAX_INTERVAL_DAYS:
        problems.append(
            f"Liczba dni musi być z zakresu {MIN_INTERVAL_DAYS}-{MAX_INTERVAL_DAYS}."
        )
    if not MIN_KEEP_COUNT <= keep_count <= MAX_KEEP_COUNT:
        problems.append(f"Liczba kopii musi być z zakresu {MIN_KEEP_COUNT}-{MAX_KEEP_COUNT}.")
    if not folder.strip():
        problems.append("Wybierz folder na kopie.")
    return problems

# ---------------------------------------------------------------------------
# Odczyt / zapis ustawień
# ---------------------------------------------------------------------------

def load_settings() -> AutoBackupSettings:
    if not SETTINGS_PATH.exists():
        return AutoBackupSettings()
    try:
        raw = json.loads(SETTINGS_PATH.read_text("utf-8"))
        known = {f.name for f in fields(AutoBackupSettings)}
        return AutoBackupSettings(**{k: v for k, v in raw.items() if k in known})
    except (OSError, ValueError, TypeError):
        logger.warning("Nie udało się odczytać %s - używam domyślnych", SETTINGS_PATH, exc_info=True)
        return AutoBackupSettings()

def save_settings(settings: AutoBackupSettings) -> None:
    temporary = SETTINGS_PATH.with_suffix(".tmp")
    temporary.write_text(json.dumps(asdict(settings), indent=2, ensure_ascii=False), "utf-8")
    temporary.replace(SETTINGS_PATH)

# ---------------------------------------------------------------------------
# Wykonanie kopii
# ---------------------------------------------------------------------------

def list_backups(folder: str | Path) -> list[Path]:
    path = Path(folder)
    if not path.is_dir():
        return []
    return sorted(path.glob(f"{FILE_PREFIX}*.db"))

def remove_old_backups(folder: str | Path, keep_count: int) -> list[Path]:
    backups = list_backups(folder)
    removed = []
    for old in backups[: max(0, len(backups) - keep_count)]:
        try:
            old.unlink(missing_ok=True)
            db.backup_keys_path(old).unlink(missing_ok=True)
            removed.append(old)
        except OSError:
            logger.warning("Nie udało się usunąć starej kopii %s", old, exc_info=True)
    return removed

def run_backup(settings: AutoBackupSettings, now: Optional[datetime] = None) -> Path:
    now = now or datetime.now()
    folder = Path(settings.folder)
    folder.mkdir(parents=True, exist_ok=True)

    dest = folder / f"{FILE_PREFIX}{now.strftime(FILE_TIMESTAMP_FORMAT)}.db"
    db.backup_database(dest)
    remove_old_backups(folder, settings.keep_count)

    settings.last_backup = now.strftime(LAST_BACKUP_FORMAT)
    save_settings(settings)
    logger.info("Utworzono automatyczną kopię %s", dest)
    return dest

def run_if_due(now: Optional[datetime] = None) -> Optional[Path]:
    if not db.is_unlocked():
        return None
    settings = load_settings()
    if not settings.is_due(now):
        return None
    return run_backup(settings, now)
