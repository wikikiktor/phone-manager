from __future__ import annotations

import base64
import hashlib
import json
import os
import secrets
import time
from pathlib import Path

from cryptography.exceptions import InvalidTag
from cryptography.hazmat.primitives.ciphers.aead import AESGCM

# Parametry scrypt: ok. 0,1 s na sprawdzenie hasła, 32 MB pamięci.
_SCRYPT_PARAMS = {"n": 2**15, "r": 8, "p": 1, "maxmem": 64 * 1024 * 1024, "dklen": 32}
_FILE_VERSION = 1
_LOCK_RETRY_DELAYS = (0.1, 0.25, 0.5, 1.0, 2.0)

RECOVERY_USER = "__awaryjny__"
MIN_PASSWORD_LENGTH = 3

_RECOVERY_ALPHABET = "ABCDEFGHJKLMNPQRSTUVWXYZ23456789"  # bez mylących znaków 0/O, 1/I
_RECOVERY_GROUPS = 6
_RECOVERY_GROUP_LENGTH = 4


class InvalidCredentials(Exception):
    """Zły login, hasło lub kod awaryjny."""


def _b64encode(data: bytes) -> str:
    return base64.b64encode(data).decode("ascii")


def _derive_key(secret: str, salt: bytes) -> bytes:
    return hashlib.scrypt(secret.encode("utf-8"), salt=salt, **_SCRYPT_PARAMS)


def normalize_username(username: str) -> str:
    return username.strip().lower()


def generate_master_key() -> bytes:
    return AESGCM.generate_key(bit_length=256)


def generate_recovery_code() -> str:
    groups = (
        "".join(secrets.choice(_RECOVERY_ALPHABET) for _ in range(_RECOVERY_GROUP_LENGTH))
        for _ in range(_RECOVERY_GROUPS)
    )
    return "-".join(groups)


def normalize_recovery_code(code: str) -> str:
    return "".join(char for char in code.upper() if char.isalnum())


class KeyStore:

    def __init__(self, path: str | Path) -> None:
        self.path = Path(path)

    # --- Odczyt / zapis pliku ---------------------------------------------

    def _load(self) -> dict:
        if not self.path.exists():
            return {"version": _FILE_VERSION, "users": {}}
        return json.loads(self.path.read_text("utf-8"))

    def _save(self, data: dict) -> None:
        temporary = self.path.with_suffix(".tmp")
        temporary.write_text(json.dumps(data, indent=2), "utf-8")
        for delay in _LOCK_RETRY_DELAYS:
            try:
                temporary.replace(self.path)
                return
            except PermissionError:
                time.sleep(delay)
        temporary.replace(self.path)

    def delete(self) -> None:
        self.path.unlink(missing_ok=True)

    # --- Informacje --------------------------------------------------------

    def is_initialized(self) -> bool:
        return bool(self._load()["users"])

    def usernames(self) -> list[str]:
        return sorted(name for name in self._load()["users"] if name != RECOVERY_USER)

    def has_user(self, username: str) -> bool:
        return normalize_username(username) in self._load()["users"]

    # --- Operacje ----------------------------------------------------------

    def add_user(self, master_key: bytes, username: str, password: str) -> None:
        name = normalize_username(username)
        salt, nonce = os.urandom(16), os.urandom(12)
        wrapped = AESGCM(_derive_key(password, salt)).encrypt(nonce, master_key, name.encode())

        data = self._load()
        data["users"][name] = {
            "salt": _b64encode(salt),
            "nonce": _b64encode(nonce),
            "key": _b64encode(wrapped),
        }
        self._save(data)

    def unlock(self, username: str, password: str) -> bytes:
        name = normalize_username(username)
        entry = self._load()["users"].get(name)
        if entry is None:
            _derive_key(password, b"\0" * 16)
            raise InvalidCredentials

        salt, nonce, wrapped = (base64.b64decode(entry[field]) for field in ("salt", "nonce", "key"))
        try:
            return AESGCM(_derive_key(password, salt)).decrypt(nonce, wrapped, name.encode())
        except InvalidTag:
            raise InvalidCredentials from None

    def unlock_with_recovery_code(self, code: str) -> bytes:
        return self.unlock(RECOVERY_USER, normalize_recovery_code(code))

    def set_recovery_code(self, master_key: bytes, code: str) -> None:
        self.add_user(master_key, RECOVERY_USER, normalize_recovery_code(code))

    def change_password(self, username: str, old_password: str, new_password: str) -> None:
        """Zmienia hasło; baza nie wymaga ponownego szyfrowania."""
        self.add_user(self.unlock(username, old_password), username, new_password)

    def remove_user(self, username: str) -> None:
        data = self._load()
        data["users"].pop(normalize_username(username), None)
        self._save(data)
