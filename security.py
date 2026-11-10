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

ROLE_ADMIN = "admin"
ROLE_USER = "user"
_ROLE_RECOVERY = "recovery"
ROLE_LABELS = {ROLE_ADMIN: "Administrator", ROLE_USER: "Użytkownik"}

_RECOVERY_ALPHABET = "ABCDEFGHJKLMNPQRSTUVWXYZ23456789"  
_RECOVERY_GROUPS = 6
_RECOVERY_GROUP_LENGTH = 4


class InvalidCredentials(Exception):
    """Zły login, hasło lub kod awaryjny."""


def _b64encode(data: bytes) -> str:
    return base64.b64encode(data).decode("ascii")


def _derive_key(secret: str, salt: bytes) -> bytes:
    return hashlib.scrypt(secret.encode("utf-8"), salt=salt, **_SCRYPT_PARAMS)

def _aad(name: str, role: str | None) -> bytes:
    return name.encode() if role is None else f"{name}\0{role}".encode()


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

    def users_with_roles(self) -> list[tuple[str, str]]:
        users = self._load()["users"]
        return [
            (name, users[name].get("role") or ROLE_ADMIN)
            for name in sorted(users)
            if name != RECOVERY_USER
        ]

    def has_user(self, username: str) -> bool:
        return normalize_username(username) in self._load()["users"]

    # --- Wpisy (wewnętrzne) ------------------------------------------------
    
    def _write_entry(self, master_key: bytes, name: str, secret: str, role: str) -> None:
        salt, nonce = os.urandom(16), os.urandom(12)
        wrapped = AESGCM(_derive_key(secret, salt)).encrypt(nonce, master_key, _aad(name, role))

        data = self._load()
        data["users"][name] = {
            "role": role,
            "salt": _b64encode(salt),
            "nonce": _b64encode(nonce),
            "key": _b64encode(wrapped)
        }
        self._save(data)

    def _open_entry(self, name: str, secret: str) -> tuple[bytes, str]:
        entry = self._load()["users"].get(name)
        if entry is None:
            _derive_key(secret, b"\0" * 16)
            raise InvalidCredentials

        role = entry.get("role")
        salt, nonce, wrapped = (base64.b64decode(entry[field]) for field in ("salt", "nonce", "key"))
        try:
            master_key = AESGCM(_derive_key(secret, salt)).decrypt(nonce, wrapped, _aad(name, role))
        except InvalidTag:
            raise InvalidCredentials from None

        if role is None:
            role = _ROLE_RECOVERY if name == RECOVERY_USER else ROLE_ADMIN
            self._write_entry(master_key, name, secret, role)
        return master_key, role

    # --- Operacje ----------------------------------------------------------

    def add_user(self, master_key: bytes, username: str, password: str, role: str = ROLE_USER) -> None:
        if role not in ROLE_LABELS:
            raise ValueError(f"Nieznana rola: {role}")
        self._write_entry(master_key, normalize_username(username), password, role)

    def unlock_with_role(self, username: str, password: str) -> tuple[bytes, str]:
        name = normalize_username(username)
        if name == RECOVERY_USER:
            raise InvalidCredentials
        return self._open_entry(name, password)

    def unlock(self, username: str, password: str) -> bytes:
        return self.unlock_with_role(username, password)[0]

    def set_password(self, master_key: bytes, username: str, password: str) -> None:
        name = normalize_username(username)
        entry = self._load()["users"].get(name)
        if entry is None or name == RECOVERY_USER:
            raise InvalidCredentials
        self._write_entry(master_key, name, password, entry.get("role") or ROLE_ADMIN)

    def unlock_with_recovery_code(self, code: str) -> bytes:
        return self._open_entry(RECOVERY_USER, normalize_recovery_code(code))[0]

    def set_recovery_code(self, master_key: bytes, code: str) -> None:
        self._write_entry(master_key, RECOVERY_USER, normalize_recovery_code(code), _ROLE_RECOVERY)

    def change_password(self, username: str, old_password: str, new_password: str) -> None:
        master_key, role = self.unlock_with_role(username, old_password)
        self._write_entry(master_key, normalize_username(username), new_password, role)

    def remove_user(self, username: str) -> None:
        data = self._load()
        data["users"].pop(normalize_username(username), None)
        self._save(data)
