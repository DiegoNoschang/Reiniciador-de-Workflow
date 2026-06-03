"""
Checkpoint / retomar execução interrompida.

Salva o progresso em ~/.iilex_checkpoint.json após cada processo. Se a
ferramenta for fechada / cair internet / Chrome travar, na próxima
execução podemos perguntar ao usuário se quer continuar de onde parou.

A "identidade" do trabalho é definida por um hash da lista de processos
+ nome do arquivo. Se a planilha mudar, o checkpoint anterior é
descartado.
"""

from __future__ import annotations

import hashlib
import json
from dataclasses import asdict, dataclass, field
from datetime import datetime
from pathlib import Path


CAMINHO_CHECKPOINT = Path.home() / ".iilex_checkpoint.json"


@dataclass
class Checkpoint:
    """Estado de uma execução em andamento."""

    arquivo_excel: str = ""           # caminho da planilha
    hash_processos: str = ""           # hash da lista de processos
    total: int = 0                     # total de processos
    proximo_indice: int = 0            # próximo a processar (0-based)
    salvo_em: str = ""                 # timestamp ISO
    processados: list[str] = field(default_factory=list)  # nºs já processados (audit)
    # Resultados acumulados (lista de dicts: numero/status/mensagem/timestamp).
    # Permite RESTAURAR contadores e relatório ao retomar — sem zerar os números
    # nem perder os que deram erro.
    resultados: list = field(default_factory=list)


def calcular_hash(processos: list[str]) -> str:
    """Hash determinístico da lista de processos."""
    h = hashlib.sha256()
    for p in processos:
        h.update(p.encode("utf-8"))
        h.update(b"\n")
    return h.hexdigest()[:16]


def carregar() -> Checkpoint | None:
    if not CAMINHO_CHECKPOINT.exists():
        return None
    try:
        dados = json.loads(CAMINHO_CHECKPOINT.read_text(encoding="utf-8"))
        return Checkpoint(**{
            k: v for k, v in dados.items()
            if k in {f for f in Checkpoint.__dataclass_fields__}
        })
    except (OSError, json.JSONDecodeError, TypeError):
        return None


def salvar(cp: Checkpoint) -> bool:
    cp.salvo_em = datetime.now().isoformat(timespec="seconds")
    try:
        CAMINHO_CHECKPOINT.write_text(
            json.dumps(asdict(cp), indent=2, ensure_ascii=False),
            encoding="utf-8",
        )
        return True
    except OSError:
        return False


def apagar() -> bool:
    try:
        if CAMINHO_CHECKPOINT.exists():
            CAMINHO_CHECKPOINT.unlink()
        return True
    except OSError:
        return False


def pode_retomar(cp: Checkpoint | None, processos_atuais: list[str], arquivo: str) -> bool:
    """True se o checkpoint corresponde aos processos atuais e ainda há trabalho."""
    if cp is None:
        return False
    if cp.proximo_indice >= cp.total or cp.proximo_indice >= len(processos_atuais):
        return False
    if calcular_hash(processos_atuais) != cp.hash_processos:
        return False
    if arquivo and cp.arquivo_excel and arquivo != cp.arquivo_excel:
        # Planilha diferente — não retomar pra evitar trabalho errado
        return False
    return True
