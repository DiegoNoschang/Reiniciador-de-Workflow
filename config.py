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


# ----------------------------------------------------------------------
# Tipos de compromisso do iiLex — alimenta o combo pesquisável da TELA
# PRINCIPAL ("Tipo de Compromisso").
#
# Lista COMPLETA extraída do <select id="_field_idtipoagenda1"> da Agenda
# do iiLex (Ramos Advogados). Para atualizar no futuro: abra a Agenda,
# F12 → Console, e rode:
#     copy([...document.querySelectorAll('select')]
#       .find(s => [...s.options].some(o => /SUBS[IÍ]DIO/i.test(o.text)))
#       .options).map... (ver histórico) — depois cole aqui.
#
#  • O match no core é por texto EXATO, mas IGNORA acento e
#    maiúscula/minúscula — então a grafia aqui não precisa ser idêntica
#    à do iiLex, só precisa ser o mesmo "nome".
#  • O combo é EDITÁVEL: mesmo que um tipo não esteja aqui, dá pra digitar.
# ----------------------------------------------------------------------
TIPOS_COMPROMISSO: list[str] = [
    "ABRIR CONSULTA - CEF",
    "ACOMPANHAMENTO DE ASSISTENTE TÉCNICO",
    "ACOMPANHAMENTO DE EXPEDIÇÃO DE CITAÇÃO",
    "ACOMPANHAMENTO OBF – MULTA - CEF",
    "ACOMPANHAMENTO PERÍCIA",
    "ACOMPANHAMENTO PROCESSUAL - CEF",
    "ACOMPANHAMENTO PROCESSUAL - RECURSAL",
    "ACOMPANHAR CONSULTA ABERTA - CEF",
    "ACOMPANHAR HABILITAÇÃO",
    "ACORDO",
    "AGRAVO DE INSTRUMENTO",
    "AGRAVO DE INSTRUMENTO - CEF",
    "AGRAVO DE INSTRUMENTO - CS",
    "AGRAVO DE INSTRUMENTO - TRABALHISTA",
    "AGRAVO DE NEGATIVA DE RESP",
    "AGRAVO DE NEGATIVO DE REXT",
    "AGRAVO DE PETIÇÃO - TRABALHISTA",
    "AGRAVO INTERNO",
    "AGRAVO INTERNO (TST) - TRABALHISTA",
    "AGRAVO REGIMENTAL (TRT) - TRABALHISTA",
    "AGRAVO RETIDO",
    "AGUARDANDO RETORNO CLIENTE",
    "AGUARDA RETORNO CORRESPONDENTE",
    "AJUSTE DE QUALIDADE",
    "ALIMENTAR SIJUR",
    "ALIMENTAR SIJUR + ENCERRAMENTO",
    "ANÁLISE DE CLASSIFICAÇÃO NADA A FAZER - DUPLICADO",
    "ANÁLISE DE CLASSIFICAÇÃO NADA A FAZER - OUTRAS PARTES",
    "ANÁLISE DE CLASSIFICAÇÃO NADA A FAZER - SUSPENSÃO",
    "ANÁLISE DE EXPEDIÇÃO DE CITAÇÃO",
    "ANÁLISE DE RELEVÂNCIA",
    "ANÁLISE PROCEDIMENTOS ENTRANTES",
    "ANÁLISE PUBLICAÇÃO/INTIMAÇÃO",
    "ANÁLISE PUBLICAÇÃO/INTIMAÇÃO - CONTESTAÇÃO",
    "ANÁLISE PUBLICAÇÃO/INTIMAÇÃO - CUMPRIMENTO DE SENTENÇA",
    "ANÁLISE PUBLICAÇÃO/INTIMAÇÃO - LIMINARES",
    "ANÁLISE PUBLICAÇÃO/INTIMAÇÃO - MEIO",
    "ANÁLISE PUBLICAÇÃO/INTIMAÇÃO - RECURSAL",
    "ANÁLISE PUBLICAÇÃO/INTIMAÇÃO - SULACAP",
    "ANÁLISE PUBLICAÇÃO/INTIMAÇÃO - TRABALHISTA",
    "ANÁLISE RECURSAL CEF",
    "ANÁLISE RECURSAL - TUTELA/LIMINAR - CEF",
    "ANOTAR CTPS - 5 DIAS",
    "APELAÇÃO",
    "APELAÇÃO - CS",
    "APELAÇÃO - RECURSAL",
    "APRESENTAÇÃO DE DOCUMENTOS",
    "APRESENTAÇÃO DE DOCUMENTOS - TRABALHISTA",
    "APRESENTAÇÃO DE QUESITOS",
    "APRESENTAR CÁLCULOS - TRABALHISTA",
    "APROPRIAÇÃO DE VALORES - CEF",
    "ATESTE - CEF",
    "ATUALIZAR SISTEMA",
    "CLASSIFICAÇÃO APTO/INAPTO PARA ACORDO - CEF",
    "Complementação De Cadastro",
    "COMPLEMENTAÇÃO DE DEFESA",
    "COMPLEMENTAÇÃO DO PREPARO",
    "COMPROMISSO TESTE",
    "COMPROVAR PAGAGAMENTO DE CONCILIADOR",
    "COMPROVAR PAGAMENTO",
    "COMPROVAR PAGAMENTO - CEF",
    "COMPROVAR PAGAMENTO - CS",
    "COMPROVAR PAGAMENTO DE CUSTAS FINAIS",
    "COMPROVAR PAGAMENTO DE HONORÁRIOS PERICIAIS",
    "CONFERÊNCIA DE VALORES",
    "CONFIRMAÇÃO DE PATROCÍNIO",
    "CONFIRMAR AF",
    "CONTESTAÇÃO",
    "CONTESTAÇÃO - CAUTELAR",
    "CONTESTAÇÃO CAUTELAR",
    "CONTESTAÇÃO - CEF",
    "CONTESTAÇÃO - ESPECIAL",
    "CONTESTAÇÃO - INDENIZATÓRIA",
    "CONTESTAÇÃO - JEF - CEF",
    "CONTESTAÇÃO - REVISIONAL",
    "CONTESTAÇÃO SUPERENDIVIDAMENTO",
    "CONTESTAÇÃO SUPERENDIVIDAMENTO - OUTROS ESTADOS",
    "CONTESTAÇÃO SUPERENDIVIDAMENTO - RS",
    "CONTESTAÇÃO - TRABALHISTA",
    "CONTRARRAZÕES 05 DIAS - CS",
    "CONTRARRAZÕES 08 DIAS - TRABALHISTA",
    "CONTRARRAZÕES - 10 DIAS",
    "CONTRARRAZÕES 10 DIAS - CS",
    "CONTRARRAZÕES - 15 DIAS",
    "CONTRARRAZÕES 15 DIAS - CS",
    "CONTRARRAZÕES - 5 DIAS",
    "CONTRARRAZÕES À IMPUGNAÇÃO - TRABALHISTA",
    "CONTRARRAZÕES AO RO - TRABALHISTA",
    "CONTRARRAZÕES AO RR - TRABALHISTA",
    "CUMPRIMENTO DE ACORDO",
    "CUMPRIMENTO DE ACORDO - 05 DIAS",
    "CUMPRIMENTO DE ACORDO - 10 DIAS",
    "CUMPRIMENTO DE ACORDO - 15 DIAS",
    "CUMPRIMENTO DE ACORDO - 48 HORAS",
    "CUMPRIMENTO DE BLOQUEIO/PENHORA - 05 DIAS",
    "CUMPRIMENTO DE BLOQUEIO/PENHORA - 10 DIAS",
    "CUMPRIMENTO DE BLOQUEIO/PENHORA - 15 DIAS",
    "CUMPRIMENTO DE BLOQUEIO/PENHORA - 48 HORAS",
    "CUMPRIMENTO DE LIMINAR - RECURSAL",
    "CUMPRIMENTO DE SALDO REMANESCENTE - 05 DIAS",
    "CUMPRIMENTO DE SALDO REMANESCENTE - 10 DIAS",
    "CUMPRIMENTO DE SALDO REMANESCENTE - 15 DIAS",
    "CUMPRIMENTO DE SALDO REMANESCENTE - 48 HORAS",
    "CUMPRIMENTO DE SENTENÇA - CEF",
    "CUMPRIMENTO DE SENTENÇA - ESPONTÂNEO",
    "CUMPRIMENTO DE SENTENÇA ESPONTÂNEO - CEF",
    "CUMPRIMENTO DE SENTENÇA - IMPULSIONADO",
    "CUMPRIMENTO DE SENTENÇA IMPULSIONADO - CEF",
    "CUMPRIR LIMINAR",
    "DEFESA ADMINISTRATIVA",
    "DESIGNADA PERÍCIA - PERÍCIA",
    "DISTRIBUIÇÃO DE AÇÃO",
    "EMBARGOS À EXECUÇÃO",
    "EMBARGOS À EXECUÇÃO - CEF",
    "EMBARGOS À EXECUÇÃO - TRABALHISTA",
    "EMBARGOS DE DECLARAÇÃO",
    "EMBARGOS DE DECLARAÇÃO - CEF",
    "EMBARGOS DE DECLARAÇÃO - CS",
    "EMBARGOS DE DECLARAÇÃO - LIMINARES",
    "EMBARGOS DE DECLARAÇÃO - RECURSAL",
    "EMBARGOS DE DECLARAÇÃO - TRABALHISTA",
    "EMBARGOS DE TERCEIROS - CEF",
    "EMBARGOS DO DEVEDOR",
    "EMBARGOS INFRINGENTES",
    "EMBARGOS MONITÓRIOS",
    "EMBARGOS (TST - SDI) - TRABALHISTA",
    "EMENDAR INICIAL",
    "EMISSÃO DE GUIA - CUSTAS FINAIS",
    "EMISSÃO DE GUIA CUSTAS FINAIS - OUTROS ESTADOS",
    "EMISSÃO DE GUIA CUSTAS FINAIS - RS",
    "EMISSÃO DE GUIA - DEPÓSITO",
    "EMISSÃO DE GUIA - IMPUGNAÇÃO",
    "EMISSÃO DE GUIA - RECURSAL",
    "ENCERRAMENTO",
    "ENCERRAMENTO DE PROCESSO",
    "ENCERRAMENTO DE PROCESSO - CEF",
    "ENCERRAMENTO DE PROCESSO - CS",
    "ENCERRAMENTO DE PROCESSO - TRABALHISTA",
    "ENVIO DE NOTIFICAÇÃO",
    "EXCEÇÃO DE INCOMPETÊNCIA",
    "EXCEÇÃO DE PRÉ-EXECUTIVIDADE",
    "EXCEÇÃO DE PRÉ-EXECUTIVIDADE - CEF",
    "EXECUÇÃO DE HONORÁRIOS",
    "FORMULAÇÃO DE QUESITOS - CEF",
    "IILEX TESTE",
    "IMISSÃO OU REINTEGRAÇÃO DE POSSE EFETIVADA - CEF",
    "IMPUGNAÇÃO 15 DIAS - CEF",
    "IMPUGNAÇÃO AO CUMPRIMENTO DE SENTENÇA - CEF",
    "IMPUGNAÇÃO AO LAUDO PERICIAL",
    "IMPUGNAÇÃO AOS CÁLCULOS - TRABALHISTA",
    "IMPUGNAÇÃO À SENTENÇA DE LIQUIDAÇÃO - TRABALHISTA",
    "IMPUGNAÇÃO - CEF",
    "IMPUGNAÇÃO CS - 15 DIAS",
    "IMPUGNAÇÃO CS - 5 DIAS",
    "IMPUGNAÇÃO DE LAUDO PERICIAL - CEF",
    "INTIMAÇÃO SOBRE INTERESSE NA PERÍCIA",
    "LAUDO PERICIAL CONTÁBIL - 05 DIAS",
    "LAUDO PERICIAL GRAFODOCUMENTOSCÓPICO",
    "LAUDO PERICIAL MÉDICO - 05 DIAS",
    "LIGAR PARA CARTÓRIO",
    "MANDADO DE SEGURANÇA",
    "MANDADO DE SEGURANÇA - CEF",
    "MANIFESTAÇÃO - 05 DIAS",
    "MANIFESTAÇÃO 05 DIAS - CS",
    "MANIFESTAÇÃO 05 DIAS - CS (LIMINAR/OBF/MULTA)",
    "MANIFESTAÇÃO 05 DIAS - MEIO",
    "MANIFESTAÇÃO - 10 DIAS",
    "MANIFESTAÇÃO 10 DIAS - CS",
    "MANIFESTAÇÃO 10 DIAS - CS (LIMINAR/OBF/MULTA)",
    "MANIFESTAÇÃO 10 DIAS - MEIO",
    "MANIFESTAÇÃO - 15 DIAS",
    "MANIFESTAÇÃO 15 DIAS - CS",
    "MANIFESTAÇÃO 15 DIAS - CS (LIMINAR/OBF/MULTA)",
    "MANIFESTAÇÃO - 15 DIAS - LIQUIDAÇÃO",
    "MANIFESTAÇÃO 15 DIAS - MEIO",
    "MANIFESTAÇÃO - 48 HORAS",
    "MANIFESTAÇÃO 48 HORAS - CS",
    "MANIFESTAÇÃO 48 HORAS - CS (LIMINAR/OBF/MULTA)",
    "MANIFESTAÇÃO - 48 HORAS- LIQUIDAÇÃO",
    "MANIFESTAÇÃO 48 HORAS - MEIO",
    "MANIFESTAÇÃO - 5 DIAS - LIQUIDAÇÃO",
    "MANIFESTAÇÃO - CS",
    "MANIFESTAÇÃO DE DESABILITAÇÃO",
    "MANIFESTAÇÃO DE HABILITAÇÃO",
    "MANIFESTAÇÃO DE HABILITAÇÃO - CS",
    "MANIFESTAÇÃO DE HABILITAÇÃO FACTA",
    "MANIFESTAÇÃO DE HABILITAÇÃO VTK",
    "MANIFESTAÇÃO DESISTÊNCIA - CEF",
    "MANIFESTAÇÃO LIMINAR",
    "MANIFESTAÇÃO – MULTA EM CURSO - CEF",
    "MANIFESTAÇÃO PENHORA/BLOQUEIO",
    "MANIFESTAÇÃO - PERÍCIA",
    "MANIFESTAÇÃO - PERÍCIA DIGITAL",
    "MANIFESTAÇÃO - PERÍCIA GRAFOTÉCNICA",
    "MANIFESTAÇÃO - PERÍCIA SUPERENDIVIDAMENTO",
    "MANIFESTAÇÃO - RECURSAL",
    "MANIFESTAÇÃO SOBRE CÁLCULOS PERICIAIS - TRABALHISTA",
    "MANIFESTAÇÃO - SUPERENDIVIDAMENTO",
    "MANIFESTAÇÃO - TRABALHISTA",
    "MEMORIAIS",
    "NOTA JURÍDICA DE PROVISIONAMENTO - CEF",
    "OBJETOS E PEDIDOS",
    "OBRIGAÇÃO DE FAZER - CS",
    "PAGAMENTO - 48 HORAS",
    "PAGAMENTO - 5 DIAS",
    "PAGAMENTO CONDENAÇÃO",
    "PAGAMENTO CUSTAS FINAIS - CS",
    "PAGAMENTO DA GUIA",
    "PAGAMENTO DE CONDENAÇÃO - TRABALHISTA",
    "PAGAMENTO DE HONORÁRIOS - CS",
    "PAGAMENTO DE HONORÁRIOS - MEIO",
    "PAGAMENTO DE IMPUGNAÇÃO",
    "PAGAMENTO ESPONTÂNEO DE CONDENAÇÃO",
    "PAGAMENTO GARANTIA DE JUÍZO",
    "PAGAMENTO HONORÁRIOS PERICIAIS - CS",
    "PAGAMENTO PREPARO IMPUGNAÇÃO - CS",
    "PAGAMENTO SALDO REMANESCENTE",
    "PAGAMENTO SALDO REMANESCENTE - CS",
    "PEDIDO DE RECONSIDERAÇÃO",
    "PERÍCIA",
    "PREPARO CUMPRIMENTO DE SENTENÇA",
    "PREPARO CUMPRIMENTO DE SENTENÇA - ESPONTÂNEO",
    "PREPARO CUMPRIMENTO DE SENTENÇA - ESPONTÂNEO - CEF",
    "PREPARO CUMPRIMENTO DE SENTENÇA - IMPULSIONADO",
    "PREPARO CUMPRIMENTO DE SENTENÇA - IMPULSIONADO - CEF",
    "PREPARO DE BLOQUEIO/PENHORA",
    "PREPARO DE CUMPRIMENTO DE ACORDO",
    "PREPARO DE SALDO REMANESCENTE",
    "PREPARO RECURSAL",
    "PRESTAÇÃO DE CONTAS AO CLIENTE",
    "PRODUÇÃO DE PROVAS",
    "PROVIDÊNCIAS PRÉ-AUDIÊNCIA",
    "RAZÕES FINAIS",
    "RAZÕES FINAIS ORAIS - TRABALHISTA",
    "RAZÕES FINAIS POR ESCRITO - TRABALHISTA",
    "REALIZAR PROTOCOLO",
    "RECONVENÇÃO",
    "RECURSO ADESIVO",
    "RECURSO ADMINISTRATIVO",
    "RECURSO DE REVISTA - TRABALHISTA",
    "RECURSO ESPECIAL - CS",
    "RECURSO ESPECIAL - RECURSAL",
    "RECURSO EXTRAORDINÁRIO",
    "RECURSO EXTRAORDINÁRIO - TRABALHISTA",
    "RECURSO INOMINADO",
    "RECURSO INOMINADO - CS",
    "RECURSO ORDINÁRIO",
    "RECURSO ORDINÁRIO - TRABALHISTA",
    "REEMBOLSO - CEF",
    "RÉPLICA",
    "RESPONDER CONSULTA - CEF",
    "RESPONDER CONSULTA - DEPURAÇÃO - CEF",
    "RESPONDER CONSULTA - MULTA - CEF",
    "RESPONDER CONSULTA - ORIENTAÇÃO - CEF",
    "RESPONDER CONSULTA - SIBAJUD - CEF",
    "RESPOSTA À NOTIFICAÇÃO - 15 DIAS",
    "REVELIA",
    "REVERSÃO DE LIMINAR",
    "REVISÃO DE CADASTRO",
    "REVISÃO DE PAUTA",
    "RMC - CEF",
    "RMS - CEF - inativo",
    "SOLICITAÇÃO DE CUMPRIMENTO ESPONTÂNEO",
    "SOLICITAÇÃO DE CUMPRIMENTO ESPONTÂNEO - CEF",
    "SOLICITAÇÃO DE CUMPRIMENTO IMPULSIONADO - CEF",
    "SOLICITAÇÃO DE PROCURAÇÃO",
    "SOLICITAÇÃO DE SUBSIDIOS",
    "SOLICITAÇÃO DE SUBSÍDIOS - CEF",
    "SOLICITAÇÃO DE SUBSÍDIOS – MULTA - CEF",
    "SOLICITAR CÁLCULO",
    "SOLICITAR PAGAMENTO ESPONTÂNEO",
    "SOLICITAR PAGAMENTO HONORÁRIOS PERICIAIS",
    "SOLICITAR PROCURAÇÃO",
    "SUBSÍDIOS - CEF",
    "TESTE IILEX",
    "TRATATIVAS DE ACORDO - CEF",
    "VERIFICAR DECISÃO FAVORÁVEL",
    "VERIFICAR JULGAMENTO",
    "VERIFICAR PRECATÓRIA",
]


@dataclass
class Settings:
    """Configurações persistidas entre execuções."""

    # Selenium
    timeout: int = 20
    headless: bool = False

    # URLs (caso o iiLex mude o domínio futuramente)
    url_login: str = "https://SEU-ESCRITORIO.iilex.com.br/sistema/login/semacesso"
    url_contencioso: str = "https://SEU-ESCRITORIO.iilex.com.br/sistema/contencioso/filtro"

    # Excel
    coluna_excel: int = 1          # 1 = coluna A, 2 = B, ...
    aba_excel: str = ""             # vazio = aba ativa
    ignorar_duplicados: bool = True
    validar_cnj: bool = False       # avisar se um nº não está no formato CNJ
    pular_invalidos: bool = False   # pular números que não passam na validação CNJ

    # Agenda — tipo de compromisso que dispara a entrada (match EXATO)
    tipo_compromisso_alvo: str = "SOLICITAÇÃO DE SUBSIDIOS"

    # WorkFlow — após entrar no compromisso, excluir o WorkFlow.
    # SEGURANÇA: só exclui se o compromisso estiver em negrito (não finalizado)
    # e nunca usa o botão 'Excluir' vermelho (que apaga o compromisso inteiro).
    excluir_workflow: bool = True

    # Comportamento
    pausa_entre_processos: float = 2.0
    retomar_automatico: bool = True   # perguntar se quer retomar checkpoint
    relogin_minutos: int = 30   # re-login proativo a cada N min (0 = desliga)
    max_retentativas_erro: int = 2   # re-tenta processo que deu erro (agenda não carregou etc.)

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
