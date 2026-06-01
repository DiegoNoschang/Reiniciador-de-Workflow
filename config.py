"""
Configurações persistentes da automação iiLex.

Salva em ~/.iilex_config.json. Se o arquivo não existir ou estiver
corrompido, usa os valores padrão e segue sem reclamar.
"""

from __future__ import annotations

import json
from dataclasses import asdict, dataclass, field, fields
from pathlib import Path


CAMINHO_CONFIG = Path.home() / ".iilex_config.json"


@dataclass
class Settings:
    """Configurações persistidas entre execuções."""

    # Selenium
    timeout: int = 20
    headless: bool = False

    # URLs (caso o iiLex mude o domínio futuramente)
    url_login: str = "https://SEU-DOMINIO.iilex.com.br/sistema/login/semacesso"
    url_contencioso: str = "https://SEU-DOMINIO.iilex.com.br/sistema/contencioso/filtro"

    # Excel
    coluna_excel: int = 1          # 1 = coluna A, 2 = B, ...
    aba_excel: str = ""             # vazio = aba ativa
    ignorar_duplicados: bool = True
    validar_cnj: bool = False       # avisar se um nº não está no formato CNJ
    pular_invalidos: bool = False   # pular números que não passam na validação CNJ

    # Agenda — tipo de compromisso que dispara a entrada (match EXATO)
    tipo_compromisso_alvo: str = "Solicitação de Subsídios"

    # WorkFlow — após entrar no compromisso, excluir o WorkFlow.
    # SEGURANÇA: só exclui se o compromisso estiver em negrito (não finalizado)
    # e nunca usa o botão 'Excluir' vermelho (que apaga o compromisso inteiro).
    excluir_workflow: bool = True

    # Comportamento
    pausa_entre_processos: float = 2.0
    retomar_automatico: bool = True   # perguntar se quer retomar checkpoint

    # Persistência
    salvar_log_em_arquivo: bool = True
    pasta_logs: str = ""    # vazio = pasta do script
    pasta_relatorios: str = ""  # vazio = pasta do script

    # Notificações
    notificar_ao_terminar: bool = True

    # Aparência
    tema: str = "escuro"  # "escuro" (padrão) | "claro"

    # ----------- I/O -----------
    @classmethod
    def carregar(cls) -> "Settings":
        """Carrega de ~/.iilex_config.json ou retorna defaults."""
        if not CAMINHO_CONFIG.exists():
            return cls()
        try:
            dados = json.loads(CAMINHO_CONFIG.read_text(encoding="utf-8"))
        except (OSError, json.JSONDecodeError):
            return cls()

        # Aceita apenas chaves que existem na dataclass (resiliente a versões)
        validos = {f.name for f in fields(cls)}
        filtrado = {k: v for k, v in dados.items() if k in validos}
        try:
            return cls(**filtrado)
        except TypeError:
            return cls()

    def salvar(self) -> bool:
        try:
            CAMINHO_CONFIG.write_text(
                json.dumps(asdict(self), indent=2, ensure_ascii=False),
                encoding="utf-8",
            )
            return True
        except OSError:
            return False

    # ----------- conveniências -----------
    def pasta_logs_resolvida(self, fallback: Path) -> Path:
        """Retorna a pasta de logs (configurada ou fallback)."""
        if self.pasta_logs:
            return Path(self.pasta_logs)
        return fallback

    def pasta_relatorios_resolvida(self, fallback: Path) -> Path:
        if self.pasta_relatorios:
            return Path(self.pasta_relatorios)
        return fallback
