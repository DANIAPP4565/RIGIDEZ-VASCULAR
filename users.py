# -*- coding: utf-8 -*-
"""
Gestión de usuarios para Vascular Health Analyzer.

Persistencia en users.json en el mismo directorio que app.py.
Hash de contraseñas con PBKDF2-HMAC-SHA256 + salt único por usuario.
Recuperación con pregunta de seguridad (su respuesta también se hashea).
"""

import hashlib
import json
import os
import secrets
from datetime import datetime
from pathlib import Path
from typing import Optional, Tuple

USERS_FILE = Path(__file__).resolve().parent / "users.json"
PBKDF2_ITERATIONS = 200_000


# ---------- HASHING ----------
def _hash(value: str, salt: str) -> str:
    """PBKDF2-HMAC-SHA256, devuelve hex."""
    dk = hashlib.pbkdf2_hmac(
        "sha256",
        value.encode("utf-8"),
        salt.encode("utf-8"),
        PBKDF2_ITERATIONS,
    )
    return dk.hex()


def _new_salt() -> str:
    return secrets.token_hex(16)


# ---------- USER STORE ----------
class UserStore:
    """Lee/escribe users.json con operaciones atómicas."""

    def __init__(self, path: Path = USERS_FILE):
        self.path = Path(path)
        self._bootstrap_if_empty()

    # --- IO ---
    def _load(self) -> dict:
        if not self.path.exists():
            return {"users": {}}
        try:
            with open(self.path, "r", encoding="utf-8") as f:
                return json.load(f)
        except (json.JSONDecodeError, OSError):
            return {"users": {}}

    def _save(self, data: dict) -> None:
        tmp = self.path.with_suffix(".tmp")
        with open(tmp, "w", encoding="utf-8") as f:
            json.dump(data, f, indent=2, ensure_ascii=False)
        os.replace(tmp, self.path)

    def _bootstrap_if_empty(self) -> None:
        """Crea usuarios admin y medico por defecto en el primer arranque."""
        data = self._load()
        if data.get("users"):
            return
        for u, pw, role in [
            ("admin", "vascular2025", "admin"),
            ("medico", "doctor123", "medico"),
        ]:
            salt = _new_salt()
            qsalt = _new_salt()
            data["users"][u] = {
                "password_hash": _hash(pw, salt),
                "salt": salt,
                "role": role,
                "security_question": "Usuario por defecto - cambiar via 'Recuperar contraseña'",
                "security_answer_hash": _hash("admin", qsalt),
                "security_salt": qsalt,
                "created_at": datetime.now().isoformat(timespec="seconds"),
                "default": True,
            }
        self._save(data)

    # --- API ---
    def existe(self, username: str) -> bool:
        return username.strip().lower() in self._load()["users"]

    def lista_usuarios(self) -> list:
        return sorted(self._load()["users"].keys())

    def registrar(
        self,
        username: str,
        password: str,
        security_question: str,
        security_answer: str,
        role: str = "medico",
        nombre_completo: str = "",
        matricula: str = "",
    ):
        """Registra un nuevo usuario. Devuelve (ok, mensaje)."""
        username = username.strip().lower()
        if not username or len(username) < 3 or not all(c.isalnum() or c in "._-" for c in username):
            return False, "El usuario debe tener al menos 3 caracteres (letras, números, . _ -)."
        if len(password) < 8:
            return False, "La contraseña debe tener al menos 8 caracteres."
        if not security_question.strip():
            return False, "Debe ingresar una pregunta de seguridad."
        if not security_answer.strip():
            return False, "Debe ingresar la respuesta a la pregunta de seguridad."

        data = self._load()
        if username in data["users"]:
            return False, "Ese usuario ya existe. Use 'Recuperar contraseña' si la olvidó."

        salt = _new_salt()
        qsalt = _new_salt()
        data["users"][username] = {
            "password_hash": _hash(password, salt),
            "salt": salt,
            "role": role,
            "nombre_completo": nombre_completo.strip(),
            "matricula": matricula.strip(),
            "security_question": security_question.strip(),
            "security_answer_hash": _hash(security_answer.strip().lower(), qsalt),
            "security_salt": qsalt,
            "created_at": datetime.now().isoformat(timespec="seconds"),
        }
        self._save(data)
        return True, f"Usuario '{username}' creado correctamente. Ya puede iniciar sesión."

    def verificar(self, username: str, password: str):
        """Verifica credenciales. Devuelve (ok, role_or_msg)."""
        username = username.strip().lower()
        data = self._load()
        u = data["users"].get(username)
        if not u:
            return False, "Usuario o contraseña incorrectos."
        if _hash(password, u["salt"]) != u["password_hash"]:
            return False, "Usuario o contraseña incorrectos."
        return True, u.get("role", "medico")

    def obtener_pregunta(self, username: str):
        u = self._load()["users"].get(username.strip().lower())
        return u["security_question"] if u else None

    def recuperar(
        self,
        username: str,
        security_answer: str,
        new_password: str,
    ):
        """Resetea la contraseña validando la pregunta de seguridad."""
        username = username.strip().lower()
        data = self._load()
        u = data["users"].get(username)
        if not u:
            return False, "Usuario no encontrado."
        ans_hash = _hash(security_answer.strip().lower(), u["security_salt"])
        if ans_hash != u["security_answer_hash"]:
            return False, "La respuesta a la pregunta de seguridad no coincide."
        if len(new_password) < 8:
            return False, "La nueva contraseña debe tener al menos 8 caracteres."

        new_salt = _new_salt()
        u["salt"] = new_salt
        u["password_hash"] = _hash(new_password, new_salt)
        u["updated_at"] = datetime.now().isoformat(timespec="seconds")
        data["users"][username] = u
        self._save(data)
        return True, "Contraseña actualizada. Ya puede iniciar sesión."

    def admin_reset(self, admin_username: str, admin_password: str,
                    target_username: str, new_password: str):
        """Permite al admin resetear cualquier contraseña."""
        ok, role = self.verificar(admin_username, admin_password)
        if not ok or role != "admin":
            return False, "Solo el administrador puede resetear contraseñas de terceros."
        target = target_username.strip().lower()
        data = self._load()
        if target not in data["users"]:
            return False, "Usuario objetivo no encontrado."
        if len(new_password) < 8:
            return False, "La nueva contraseña debe tener al menos 8 caracteres."
        new_salt = _new_salt()
        data["users"][target]["salt"] = new_salt
        data["users"][target]["password_hash"] = _hash(new_password, new_salt)
        data["users"][target]["updated_at"] = datetime.now().isoformat(timespec="seconds")
        self._save(data)
        return True, f"Contraseña de '{target}' actualizada."