"""
Persistência de credenciais (login/senha) do iiLex.

No Windows tenta usar a DPAPI (via pywin32) — criptografia atrelada à
conta de usuário do SO. Se pywin32 não estiver instalado, cai para
base64 (= ofuscação, NÃO criptografia) e avisa.

Formato do arquivo:
    {
        "usuario": "caio.romeira",
        "senha_protegida": "<bytes em base64>",
        "modo": "dpapi" | "base64"
    }
"""

from __future__ import annotations

import base64
import json
import sys
from dataclasses import dataclass
from pathlib import Path


CAMINHO_CRED = Path.home() / ".iilex_login.dat"

# Tenta importar pywin32 (só faz sentido no Windows)
_DPAPI_DISPONIVEL = False
if sys.platform == "win32":
    try:
        import win32crypt  # type: ignore
        _DPAPI_DISPONIVEL = True
    except ImportError:
        _DPAPI_DISPONIVEL = False


@dataclass
class Credenciais:
    usuario: str
    senha: str


def dpapi_disponivel() -> bool:
    """True se conseguimos usar DPAPI (Windows + pywin32 instalado)."""
    return _DPAPI_DISPONIVEL


# ----------- criptografia -----------
def _proteger(plain: str) -> tuple[bytes, str]:
    """Retorna (bytes_protegidos, modo)."""
    if _DPAPI_DISPONIVEL:
        try:
            protected = win32crypt.CryptProtectData(
                plain.encode("utf-8"),
                "iilex-senha",  # descrição
                None, None, None, 0
            )
            return protected, "dpapi"
        except Exception:
            pass  # cai pro fallback
    return plain.encode("utf-8"), "base64"


def _desproteger(blob: bytes, modo: str) -> str:
    if modo == "dpapi" and _DPAPI_DISPONIVEL:
        try:
            _desc, plain_bytes = win32crypt.CryptUnprotectData(blob, None, None, None, 0)
            return plain_bytes.decode("utf-8")
        except Exception:
            return ""
    # base64 mode
    try:
        return blob.decode("utf-8")
    except UnicodeDecodeError:
        return ""


# ----------- API pública -----------
def salvar(creds: Credenciais) -> bool:
    """Salva credenciais. Retorna True em caso de sucesso."""
    try:
        blob, modo = _proteger(creds.senha)
        dados = {
            "usuario": creds.usuario,
            "senha_protegida": base64.b64encode(blob).decode("ascii"),
            "modo": modo,
        }
        CAMINHO_CRED.write_text(
            json.dumps(dados), encoding="utf-8"
        )
        return True
    except OSError:
        return False


def carregar() -> Credenciais | None:
    """Carrega credenciais salvas, ou None se não existir / inválido."""
    if not CAMINHO_CRED.exists():
        return None
    try:
        raw = CAMINHO_CRED.read_text(encoding="utf-8")
        dados = json.loads(raw)
    except (OSError, json.JSONDecodeError):
        # Pode ser o formato antigo (tkinter) — só usuario e senha base64
        try:
            dados = json.loads(CAMINHO_CRED.read_text(encoding="utf-8"))
        except Exception:
            return None

    usuario = dados.get("usuario", "")
    if not usuario:
        return None

    # Formato novo: senha_protegida + modo
    if "senha_protegida" in dados:
        try:
            blob = base64.b64decode(dados["senha_protegida"])
            modo = dados.get("modo", "base64")
            senha = _desproteger(blob, modo)
            return Credenciais(usuario=usuario, senha=senha)
        except Exception:
            return None

    # Formato antigo (compatibilidade com versão tkinter): senha em base64
    if "senha" in dados:
        try:
            senha = base64.b64decode(dados["senha"]).decode("utf-8")
            return Credenciais(usuario=usuario, senha=senha)
        except Exception:
            return None

    return None


def apagar() -> bool:
    """Remove o arquivo de credenciais (quando o usuário desmarca 'Lembrar')."""
    try:
        if CAMINHO_CRED.exists():
            CAMINHO_CRED.unlink()
        return True
    except OSError:
        return False
