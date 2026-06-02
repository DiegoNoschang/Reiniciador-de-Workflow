"""
Núcleo da automação iiLex — independente de interface gráfica.

Esse módulo contém TODA a lógica de Selenium (login, pesquisa do processo,
leitura da Agenda e entrada no compromisso-alvo). Ele expõe a classe
IilexAutomation, que recebe callbacks para comunicação com a UI
(PyQt, console, testes...).

Fluxo atual:
    login → Contencioso → pesquisa o processo → abre a seção Agenda →
    localiza o compromisso do tipo configurado (match EXATO) que ainda
    não foi concluído → clica na seta para ENTRAR nele.

Regras de decisão (por processo):
    • 1 compromisso pendente do tipo  → entra (clica na seta)
    • 2+ pendentes do mesmo tipo       → avisa e pula (registra na log)
    • 0 pendentes mas já concluído     → avisa "já havia sido concluída"
    • nenhum do tipo                   → registra "sem compromisso"

Vantagem: bug fix de seletor entra em UM lugar só.
"""

from __future__ import annotations

import re
import time
import unicodedata
from dataclasses import dataclass, field
from datetime import datetime
from enum import Enum
from typing import Callable

from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait
from selenium.common.exceptions import (
    ElementClickInterceptedException,
    TimeoutException,
    WebDriverException,
)
from webdriver_manager.chrome import ChromeDriverManager


# ============================================================
# CONSTANTES — seletores e URLs
# ============================================================
URL_LOGIN_DEFAULT = "https://ramosadv.iilex.com.br/sistema/login/semacesso"
URL_CONTENCIOSO_DEFAULT = "https://ramosadv.iilex.com.br/sistema/contencioso/filtro"

SEL_LOGIN = (By.ID, "texto1")
SEL_SENHA = (By.ID, "texto2")
SEL_BOTAO_LOGIN = (By.ID, "salvaDados")

SEL_CAMPO_PROCESSO = (By.ID, "_field_processo.texto1")
SEL_BOTAO_PESQUISA = (By.ID, "pesquisa")
SEL_MODAL_LOADING = (By.ID, "modalLoading")

# ── Seção "Agenda" do processo ──────────────────────────────
# A tabela da Agenda é localizada de forma ROBUSTA pela presença do
# cabeçalho "Tipo de Compromisso" — não depende de id/classe instáveis.
SEL_TABELA_AGENDA = (
    By.XPATH,
    "//table[.//*[self::th or self::td]"
    "[contains(normalize-space(.), 'Tipo de Compromisso')]]",
)

# Cabeçalho/barra da seção "Agenda" — link de colapso (Bootstrap) que
# expande/recolhe a seção. Confirmado via HTML:
#   <a data-toggle="collapse" href="#..."><i class="fa fa-caret-up"></i> Agenda (...)</a>
SEL_AGENDA_TOGGLE = (
    By.XPATH,
    "//a[@data-toggle='collapse'][contains(normalize-space(.), 'Agenda')]",
)

# Contador do cabeçalho "Agenda ( N )" — <span class="qtd-submodulo">N</span>.
# A seção Agenda é um submódulo carregado de forma ASSÍNCRONA (data-async):
# este número diz quantos compromissos esperar, permitindo distinguir
# "tabela ainda carregando via AJAX" de "Agenda legitimamente vazia".
SEL_AGENDA_QTD = (
    By.XPATH,
    "//a[@data-toggle='collapse'][contains(normalize-space(.), 'Agenda')]"
    "//span[contains(@class, 'qtd-submodulo')]",
)

# Dentro da linha (tr) que casa o tipo de compromisso, o link da SETA que
# abre o compromisso. Confirmado via HTML:
#   <a data-idedicao="editardados" href=".../agenda/editardados/12585817">
#       <i class="fa fa-external-link-square fa-2x"></i></a>
# Sem target="_blank" → navega na MESMA aba.
SEL_REL_SETA_CANDIDATOS = [
    ".//a[@data-idedicao='editardados']",
    ".//a[contains(@href,'/agenda/editardados/')]",
    ".//a[.//i[contains(@class,'fa-external-link')]]",
]

# Lista de RESULTADOS da pesquisa (Contencioso): seta ↗ que abre o processo.
# Confirmado via HTML:  <a href="editardados/593433">
#   <i class="fa fa-external-link-square fa-2x"></i></a>
# href RELATIVO e sem 'agenda' — distingue da seta da Agenda.
XPATH_SETA_RESULTADO = (
    "//a[contains(@href,'editardados') and not(contains(@href,'agenda')) "
    "and .//i[contains(@class,'fa-external-link')]]"
)

# ── Exclusão do WorkFlow (tela do compromisso, após entrar) ─────────
# Botão "Excluir" do submódulo WorkFlow: <a> btn-default (cinza/branco).
# Confirmado via HTML:
#   <a data-idcomponente="botao-excluir-submodulo" data-tabela="historicoworkflow"
#      class="btn btn-default btn-xs pull-right btn-gatilho-excluir" href="#">
#      <i class="glyphicon glyphicon-trash"></i> <span>Excluir</span></a>
# NÃO confundir com <button id="excluir-processo" class="btn-danger"> (VERMELHO),
# que apaga o compromisso inteiro. O par data-idcomponente + data-tabela é
# específico do WorkFlow e impossível de confundir.
SEL_BTN_EXCLUIR_WORKFLOW = (
    By.XPATH,
    "//a[@data-idcomponente='botao-excluir-submodulo']"
    "[@data-tabela='historicoworkflow']",
)
# Modal de confirmação global (#confirm) — botão "Confirmar" (btn-primary).
# Confirmado via HTML: modal id="confirm" com botões Cancelar (btn-default) e
# Confirmar (btn-primary).
SEL_MODAL_BTN_CONFIRMAR = (
    By.XPATH,
    "//div[@id='confirm']//button[contains(@class, 'btn-primary')]"
    "[contains(normalize-space(.), 'Confirmar')]",
)
# Botão "Iniciar Workflow" — aparece quando o WorkFlow fica vazio (após a
# exclusão), para começar um novo fluxo. Confirmado via HTML:
#   <a class="btn btn-default" id="iniciar-workflow" data-modulo="agenda"
#      data-tabela="agendas" href="#"> ... Iniciar Workflow</a>
SEL_BTN_INICIAR_WORKFLOW = (By.ID, "iniciar-workflow")


# ============================================================
# TIPOS PÚBLICOS
# ============================================================
class StatusProcesso(str, Enum):
    WORKFLOW_REINICIADO = "WorkFlow reiniciado"
    ENTROU = "Entrou no compromisso"
    SEM_COMPROMISSO = "Sem compromisso do tipo"
    JA_CONCLUIDO = "Já concluído"
    MULTIPLOS = "Múltiplos pendentes"
    NAO_ENCONTRADO = "Processo não encontrado"
    ERRO = "Erro"
    PULADO = "Pulado"


@dataclass
class ResultadoProcesso:
    numero: str
    status: StatusProcesso
    mensagem: str = ""
    timestamp: str = field(default_factory=lambda: datetime.now().strftime("%H:%M:%S"))


@dataclass
class ConfigAutomacao:
    """Parâmetros configuráveis da automação."""
    timeout: int = 20
    headless: bool = False
    url_login: str = URL_LOGIN_DEFAULT
    url_contencioso: str = URL_CONTENCIOSO_DEFAULT
    pausa_entre_processos: float = 2.0
    # Tipo de compromisso (na Agenda) que dispara a entrada via seta.
    # Comparação EXATA (após normalizar): variações com sufixo NÃO casam
    # (ex.: "Solicitação de Subsídios - CEF" é ignorada).
    tipo_compromisso_alvo: str = "Solicitação de Subsídios"
    # Após entrar no compromisso, excluir o WorkFlow (botão 'Excluir' do
    # submódulo). SEGURANÇA: só exclui se o compromisso estiver em NEGRITO
    # (= não finalizado) e NUNCA usa o 'Excluir' vermelho (que apaga o
    # compromisso inteiro).
    excluir_workflow: bool = True
    # Re-login PROATIVO: reloga a cada N minutos para a sessão não expirar no
    # meio de execuções longas (0 = desliga o re-login por tempo).
    relogin_minutos: int = 30


@dataclass
class CallbacksAutomacao:
    """Callbacks usados pela UI para receber eventos do worker.

    Todos são opcionais — se a UI não precisar de algum, basta deixar
    o default (no-op).
    """
    on_log: Callable[[str, str], None] = field(default=lambda level, msg: None)
    on_progress: Callable[[int, int], None] = field(default=lambda i, t: None)
    on_status: Callable[[str], None] = field(default=lambda s: None)
    on_resultado: Callable[[ResultadoProcesso], None] = field(default=lambda r: None)
    is_stop_requested: Callable[[], bool] = field(default=lambda: False)
    is_pause_requested: Callable[[], bool] = field(default=lambda: False)


# ============================================================
# VALIDAÇÃO DE Nº DE PROCESSO (formato CNJ)
# ============================================================
# NNNNNNN-DD.AAAA.J.TR.OOOO  (com ou sem pontuação)
_CNJ_RE = re.compile(r"^\d{7}-?\d{2}\.?\d{4}\.?\d{1}\.?\d{2}\.?\d{4}$")


def validar_cnj(numero: str) -> bool:
    """Retorna True se o número segue o padrão CNJ."""
    if not numero:
        return False
    numero_normalizado = numero.replace(" ", "").replace("/", "")
    return bool(_CNJ_RE.match(numero_normalizado))


def normalizar_texto(s: str) -> str:
    """Normaliza texto para comparação robusta.

    - remove acentos (SUBSÍDIOS == SUBSIDIOS)
    - minúsculas
    - colapsa espaços em branco

    Usado para o match EXATO do tipo de compromisso, de forma que
    "Solicitação de Subsídios" case independentemente de acento/caixa,
    mas "Solicitação de Subsídios - CEF" continue NÃO casando.
    """
    if not s:
        return ""
    s = unicodedata.normalize("NFKD", s)
    s = "".join(c for c in s if not unicodedata.combining(c))
    return " ".join(s.lower().split())


def normalizar_numero(s: str) -> str:
    """Mantém apenas letras/dígitos para comparar número de processo.

    Robusto a pontuação/espaços, mas preserva a distinção exata:
        '30003333'  != '300033331'   (continuam diferentes)
        '1234567-89.2024...' == '123456789...' (ignora . e -)
    """
    return re.sub(r"[^0-9a-z]", "", normalizar_texto(s))


# ============================================================
# AUTOMAÇÃO PRINCIPAL
# ============================================================
class IilexAutomation:
    """Executa o fluxo de leitura da Agenda e entrada no compromisso no iiLex.

    Uso típico:
        config = ConfigAutomacao(headless=True)
        callbacks = CallbacksAutomacao(on_log=lambda lvl, m: print(m))
        bot = IilexAutomation(config, callbacks)
        bot.executar("user", "senha", ["proc1", "proc2"])
        bot.resultados  # lista de ResultadoProcesso
    """

    def __init__(self, config: ConfigAutomacao, callbacks: CallbacksAutomacao):
        self.config = config
        self.cb = callbacks
        self.driver: webdriver.Chrome | None = None
        self.resultados: list[ResultadoProcesso] = []

    # ---------- helpers de log ----------
    def _info(self, msg: str, *args):
        self.cb.on_log("INFO", msg % args if args else msg)

    def _ok(self, msg: str, *args):
        self.cb.on_log("OK", msg % args if args else msg)

    def _warn(self, msg: str, *args):
        self.cb.on_log("WARNING", msg % args if args else msg)

    def _err(self, msg: str, *args):
        self.cb.on_log("ERROR", msg % args if args else msg)

    # ---------- driver ----------
    def _criar_driver(self) -> webdriver.Chrome:
        options = Options()
        if self.config.headless:
            options.add_argument("--headless=new")
            options.add_argument("--window-size=1920,1080")
        else:
            options.add_argument("--start-maximized")
        options.add_argument("--disable-blink-features=AutomationControlled")
        options.add_argument("--no-sandbox")
        options.add_argument("--disable-dev-shm-usage")
        service = Service(ChromeDriverManager().install())
        driver = webdriver.Chrome(service=service, options=options)
        # Falha em ~60s se uma página travar (ex.: site fora do ar), em vez de
        # ficar pendurado no padrão (~300s) do Chrome.
        try:
            driver.set_page_load_timeout(max(self.config.timeout, 60))
        except WebDriverException:
            pass
        return driver

    # ---------- pausa cooperativa ----------
    def _esperar_pausa(self):
        """Bloqueia enquanto a UI pedir pausa, mas sempre checando stop."""
        if not self.cb.is_pause_requested():
            return
        self.cb.on_status("⏸ Pausado")
        self._info("Pausado pelo usuário.")
        while self.cb.is_pause_requested():
            if self.cb.is_stop_requested():
                return
            time.sleep(0.3)
        self._info("Retomando.")

    def _sleep_cooperativo(self, segundos: float):
        """Sleep que respeita stop e pause."""
        fim = time.time() + segundos
        while time.time() < fim:
            if self.cb.is_stop_requested():
                return
            if self.cb.is_pause_requested():
                self._esperar_pausa()
            time.sleep(min(0.2, max(0.0, fim - time.time())))

    # ---------- fluxo principal ----------
    def executar(
        self,
        usuario: str,
        senha: str,
        processos: list[str],
        retomar_de: int = 0,
    ) -> list[ResultadoProcesso]:
        """Executa o fluxo completo.

        Args:
            usuario: login do iiLex
            senha: senha do iiLex
            processos: lista de números de processo
            retomar_de: índice (0-based) por onde começar
                (útil para checkpoint/retomar)
        """
        self.resultados = []
        self._usuario, self._senha = usuario, senha  # guardado p/ re-login automático
        self._ultimo_login = 0.0  # horário do último login (re-login proativo)
        try:
            self.cb.on_status("Iniciando navegador...")
            self._info("Iniciando automação%s...",
                       " em modo headless" if self.config.headless else "")
            self.driver = self._criar_driver()

            if self.cb.is_stop_requested():
                return self.resultados

            if not self._login(usuario, senha):
                return self.resultados

            self.cb.on_status("Logado. Processando...")
            self._esperar_documento_pronto()

            total = len(processos)
            if retomar_de > 0:
                self._info("Retomando do processo %d/%d", retomar_de + 1, total)
            self.cb.on_progress(retomar_de, total)

            for i in range(retomar_de, total):
                if self.cb.is_stop_requested():
                    self._warn("Parado pelo usuário.")
                    return self.resultados
                self._esperar_pausa()
                self._manter_sessao()  # re-login proativo a cada N min

                numero = processos[i]
                self.cb.on_status(f"Processando {i + 1}/{total}: {numero}")
                self._info("[%d/%d] Processando: %s", i + 1, total, numero)

                resultado = self._processar_um(numero)
                self.resultados.append(resultado)
                self.cb.on_resultado(resultado)
                self.cb.on_progress(i + 1, total)

                self._sleep_cooperativo(self.config.pausa_entre_processos)

            if total > 0:
                resumo = self._gerar_resumo()
                self._ok("Concluído — %s", resumo)
                self.cb.on_status(f"Concluído — {resumo}")
            else:
                self._warn("Nenhum processo recebido.")
                self.cb.on_status("Concluído (sem processos).")

        except WebDriverException as e:
            self._err("Erro no navegador: %s", e.msg)
        except Exception as e:
            self._err("Erro inesperado: %s", str(e))
        finally:
            if self.driver:
                self._info("Fechando navegador.")
                try:
                    self.driver.quit()
                except Exception:
                    pass
                self.driver = None
            self._info("Automação finalizada.")

        return self.resultados

    def _gerar_resumo(self) -> str:
        c = {s: 0 for s in StatusProcesso}
        for r in self.resultados:
            c[r.status] += 1
        return (
            f"{c[StatusProcesso.WORKFLOW_REINICIADO]} workflow(s) reiniciado(s), "
            f"{c[StatusProcesso.ENTROU]} entrou (sem excluir), "
            f"{c[StatusProcesso.SEM_COMPROMISSO]} sem compromisso, "
            f"{c[StatusProcesso.JA_CONCLUIDO]} já concluídos, "
            f"{c[StatusProcesso.MULTIPLOS]} múltiplos, "
            f"{c[StatusProcesso.NAO_ENCONTRADO]} não encontrados, "
            f"{c[StatusProcesso.ERRO]} erros"
        )

    # ---------- login ----------
    def _login(self, usuario: str, senha: str) -> bool:
        self.cb.on_status("Fazendo login...")
        self._info("Abrindo site iiLex...")
        self.driver.get(self.config.url_login)
        wait = WebDriverWait(self.driver, self.config.timeout)

        try:
            # O iiLex exibe um overlay #modalLoading ("Aguarde...") que pode
            # cobrir os campos/botão logo ao abrir a página. Espera sumir.
            self._esperar_loading()

            self._info("Preenchendo login e senha...")
            campo_user = wait.until(EC.presence_of_element_located(SEL_LOGIN))
            campo_user.clear()
            campo_user.send_keys(usuario)

            campo_senha = self.driver.find_element(*SEL_SENHA)
            campo_senha.clear()
            campo_senha.send_keys(senha)

            self._info("Clicando em Login...")
            # Garante que o overlay não intercepte o clique no botão de Login.
            self._esperar_loading()
            botao_login = wait.until(EC.element_to_be_clickable(SEL_BOTAO_LOGIN))
            try:
                botao_login.click()
            except ElementClickInterceptedException:
                self._warn("Clique no Login interceptado; usando JavaScript.")
                self.driver.execute_script("arguments[0].click();", botao_login)

            try:
                wait.until(EC.url_changes(self.config.url_login))
            except TimeoutException:
                self._err("Login falhou — verifique usuário e senha.")
                return False

            if "semacesso" in self.driver.current_url.lower():
                self._err("Login falhou — voltou para a tela de login.")
                return False

            self._ok("Login bem-sucedido!")
            self._ultimo_login = time.time()  # marca p/ o re-login proativo
            return True
        except TimeoutException:
            self._err("Timeout no login.")
            return False

    def _esperar_documento_pronto(self):
        """Espera o documento estar totalmente carregado (sem timer fixo)."""
        try:
            WebDriverWait(self.driver, self.config.timeout).until(
                lambda d: d.execute_script("return document.readyState") == "complete"
            )
        except TimeoutException:
            self._warn("Página demorou a carregar; seguindo mesmo assim.")

    # ---------- detectar sessão expirada ----------
    def _sessao_expirou(self) -> bool:
        """True se fomos jogados de volta pra tela de login (sessão expirada)."""
        try:
            url = (self.driver.current_url or "").lower()
            if "semacesso" in url or "/login" in url:
                return True
            # Fallback: campo de senha presente onde não deveria = caiu no login.
            return bool(self.driver.find_elements(*SEL_SENHA))
        except WebDriverException:
            return False

    def _tentar_religar(self, usuario: str, senha: str) -> bool:
        """Re-login automático se a sessão expirou no meio do processamento."""
        self._warn("Sessão pode ter expirado — tentando relogar...")
        if not self._login(usuario, senha):
            self._err("Não foi possível religar.")
            return False
        self._esperar_documento_pronto()
        return True

    def _manter_sessao(self):
        """Re-login PROATIVO por tempo: se já passou `relogin_minutos` desde o
        último login, reloga ANTES de continuar — evita a sessão do iiLex
        expirar no meio de execuções longas. `relogin_minutos = 0` desliga.
        """
        minutos = getattr(self.config, "relogin_minutos", 30)
        if not minutos or minutos <= 0:
            return
        if not self._ultimo_login:
            return
        if (time.time() - self._ultimo_login) < minutos * 60:
            return
        self._info("Já se passaram ~%d min desde o login — relogando "
                   "preventivamente...", minutos)
        # Marca a tentativa AGORA (evita repetir a cada processo se falhar);
        # se o _login der certo, ele atualiza _ultimo_login de novo.
        self._ultimo_login = time.time()
        self._tentar_religar(self._usuario, self._senha)

    # ---------- resiliência: site fora do ar ----------
    def _site_respondeu(self) -> bool:
        """True se carregou uma página reconhecível do iiLex — a de pesquisa
        (Contencioso) OU a de login. Se não há NENHUM dos dois, o site
        provavelmente está fora do ar / mostrando página de erro.
        """
        try:
            return bool(
                self.driver.find_elements(*SEL_CAMPO_PROCESSO)
                or self.driver.find_elements(*SEL_LOGIN)
                or self.driver.find_elements(*SEL_SENHA)
            )
        except WebDriverException:
            return False

    def _abrir_contencioso_resiliente(self) -> bool:
        """Abre a tela de Contencioso. Se o iiLex estiver FORA DO AR, FICA
        TENTANDO (com espera crescente) até o site voltar. Cooperativo: para
        se o usuário clicar em Parar. Retorna True quando o site respondeu,
        ou False se o usuário parou.
        """
        espera = 15
        avisou = False
        while not self.cb.is_stop_requested():
            try:
                self.driver.get(self.config.url_contencioso)
                self._esperar_loading()
                if self._site_respondeu():
                    if avisou:
                        self._ok("iiLex voltou ao ar — retomando.")
                    return True
            except WebDriverException as e:
                self._warn("iiLex não respondeu (%s).",
                           getattr(e, "msg", "") or type(e).__name__)
            avisou = True
            self.cb.on_status("Site do iiLex fora do ar — aguardando voltar...")
            self._warn("Site do iiLex parece fora do ar — aguardando %ds e "
                       "tentando de novo (deixe rodando, ele retoma sozinho).",
                       espera)
            self._sleep_cooperativo(espera)
            espera = min(espera + 15, 120)  # espera crescente, até 2 min
        return False

    # ---------- processar um único processo ----------
    def _processar_um(self, numero: str, _ja_religou: bool = False) -> ResultadoProcesso:
        wait = WebDriverWait(self.driver, self.config.timeout)
        try:
            self._info("Abrindo tela de Contencioso...")
            # Se o site estiver fora do ar, espera ele voltar (fica tentando).
            if not self._abrir_contencioso_resiliente():
                return ResultadoProcesso(numero, StatusProcesso.PULADO,
                                         "Parado pelo usuário")

            # A sessão do iiLex expira após um tempo logado. Se ao abrir o
            # Contencioso caímos no login, religa automaticamente e tenta de novo.
            if self._sessao_expirou():
                if _ja_religou:
                    msg = "Sessão expirou e o re-login não resolveu"
                    self._err("%s [%s].", msg, numero)
                    return ResultadoProcesso(numero, StatusProcesso.ERRO, msg)
                self._warn("Sessão expirada detectada antes de %s — religando...",
                           numero)
                if self._tentar_religar(self._usuario, self._senha):
                    return self._processar_um(numero, _ja_religou=True)
                return ResultadoProcesso(
                    numero, StatusProcesso.ERRO,
                    "Sessão expirou e não foi possível religar")

            campo_proc = wait.until(
                EC.presence_of_element_located(SEL_CAMPO_PROCESSO)
            )
            campo_proc.clear()
            campo_proc.send_keys(numero)
            self._info("Número preenchido: %s", numero)

            self._esperar_loading()

            botao_pesq = wait.until(EC.element_to_be_clickable(SEL_BOTAO_PESQUISA))
            try:
                botao_pesq.click()
            except ElementClickInterceptedException:
                self._warn("Clique em Pesquisar interceptado; usando JavaScript.")
                self.driver.execute_script("arguments[0].click();", botao_pesq)
            self._ok("Pesquisa realizada para %s", numero)

            time.sleep(1)
            self._esperar_loading()

            if self.cb.is_stop_requested():
                return ResultadoProcesso(numero, StatusProcesso.PULADO, "Parado pelo usuário")

            # A busca do iiLex casa por prefixo ('30003333' traz '300033331').
            # Seleciona, na lista de resultados, o processo de número EXATO.
            erro_sel = self._abrir_processo_correto(numero)
            if erro_sel is not None:
                self._warn("%s [%s].", erro_sel, numero)
                return ResultadoProcesso(numero, StatusProcesso.NAO_ENCONTRADO, erro_sel)
            self._esperar_loading()

            return self._processar_agenda(numero)
        except WebDriverException as e:
            # Pode ter sido sessão expirada bem no meio do processo: religa 1x.
            if not _ja_religou and self._sessao_expirou():
                if self._tentar_religar(self._usuario, self._senha):
                    return self._processar_um(numero, _ja_religou=True)
            # Site caiu no meio do processo? Espera ele voltar (pra não queimar
            # os próximos como erro) e marca ESTE p/ reprocessar depois.
            if not self._site_respondeu():
                self._warn("Site do iiLex caiu durante %s — aguardando voltar...",
                           numero)
                self._abrir_contencioso_resiliente()
                return ResultadoProcesso(
                    numero, StatusProcesso.ERRO,
                    "Site caiu durante o processamento (reprocessar depois)")
            # Mensagem NUNCA vazia (TimeoutException vem sem msg) + URL p/ diagnóstico.
            msg = getattr(e, "msg", "") or str(e) or type(e).__name__
            try:
                msg = f"{msg} [URL: {self.driver.current_url}]"
            except Exception:
                pass
            self._err("Erro processando %s: %s", numero, msg)
            return ResultadoProcesso(numero, StatusProcesso.ERRO, msg)
        except Exception as e:
            self._err("Erro inesperado em %s: %s", numero, str(e))
            return ResultadoProcesso(numero, StatusProcesso.ERRO, str(e))

    # ---------- selecionar o processo certo na lista de resultados ----------
    def _abrir_processo_correto(self, numero: str):
        """Na LISTA de resultados, abre o processo de número EXATO.

        A busca do iiLex casa por prefixo, então pesquisar '30003333'
        também traz '300033331'. Aqui achamos a linha cujo número é
        EXATAMENTE o pesquisado e clicamos na seta ↗ dela.

        Retorna:
            None  -> sucesso (abriu) OU já estávamos na página de detalhe
            str   -> mensagem de falha (número exato não encontrado)
        """
        try:
            url = self.driver.current_url
        except Exception:
            url = ""
        # Se a busca já abriu direto a página de detalhe, não há lista.
        if "editardados" in url:
            return None

        try:
            arrows = self.driver.find_elements(By.XPATH, XPATH_SETA_RESULTADO)
        except WebDriverException:
            arrows = []
        if not arrows:
            # Sem lista detectável — deixa o passo da Agenda decidir/diagnosticar.
            return None

        alvo = normalizar_numero(numero)
        seta_alvo = None
        numeros_vistos: list[str] = []
        for a in arrows:
            try:
                row = a.find_element(By.XPATH, "./ancestor::tr[1]")
                cells = row.find_elements(By.TAG_NAME, "td")
            except WebDriverException:
                continue
            nums = []
            for c in cells:
                txt = (c.text or "").strip()
                if not txt:
                    continue
                norm = normalizar_numero(txt)
                if norm and any(ch.isdigit() for ch in norm):
                    nums.append(norm)
            numeros_vistos.extend(nums)
            if alvo in nums:           # match EXATO do número
                seta_alvo = a
                break

        if seta_alvo is None:
            vistos = ", ".join(dict.fromkeys(numeros_vistos)) or "(nenhum)"
            self._warn("Números vistos na lista: %s", vistos[:200])
            return f"Número exato '{numero}' não encontrado na lista de resultados"

        self._info("Resultado de número exato encontrado — abrindo o processo...")
        try:
            self.driver.execute_script(
                "arguments[0].scrollIntoView({block:'center'});", seta_alvo)
            seta_alvo.click()
        except ElementClickInterceptedException:
            self.driver.execute_script("arguments[0].click();", seta_alvo)
        except Exception as e:
            return f"Erro ao abrir o processo na lista: {e}"
        time.sleep(1)
        self._esperar_loading()
        return None

    # ---------- processar a Agenda do processo ----------
    def _processar_agenda(self, numero: str) -> ResultadoProcesso:
        """Lê a Agenda do processo e age sobre o compromisso-alvo.

        O match do tipo é EXATO (após normalizar acento/caixa), então
        "Solicitação de Subsídios - CEF" NÃO casa com "Solicitação de
        Subsídios".
        """
        alvo = normalizar_texto(self.config.tipo_compromisso_alvo)
        if not alvo:
            msg = "Tipo de compromisso-alvo não configurado"
            self._err("%s.", msg)
            return ResultadoProcesso(numero, StatusProcesso.ERRO, msg)

        tabela = self._localizar_tabela_agenda()
        if tabela is None:
            self._diagnostico_sem_tabela()
            msg = "Tabela da Agenda não encontrada"
            self._warn("%s para %s.", msg, numero)
            return ResultadoProcesso(numero, StatusProcesso.ERRO, msg)

        # Índice da coluna "Concluído" (lido do cabeçalho): leitura precisa
        # do Sim/Não, sem falso-positivo caso outra coluna contenha "sim".
        idx_concluido = self._indice_coluna(tabela, "concluido")

        pendentes, concluidas = [], []
        linhas_com_dados = 0
        tipos_vistos: list[str] = []  # diagnóstico
        try:
            linhas = tabela.find_elements(By.XPATH, ".//tr")
        except WebDriverException:
            linhas = []

        for row in linhas:
            try:
                cells = row.find_elements(By.TAG_NAME, "td")
            except WebDriverException:
                continue
            if not cells:
                continue  # provável linha de cabeçalho (usa <th>) ou vazia
            linhas_com_dados += 1
            textos = [normalizar_texto(c.text) for c in cells]
            # guarda células "de texto" p/ diagnóstico (ignora vazias/sim/não/datas)
            for t in textos:
                if (t and t not in ("sim", "nao") and len(t) > 4
                        and t not in tipos_vistos):
                    tipos_vistos.append(t)
            # Match EXATO: alguma célula é exatamente o tipo-alvo?
            if alvo not in textos:
                continue
            # Concluído? Lê a coluna "Concluído" pelo índice do cabeçalho; se
            # o índice não foi achado, cai na heurística (alguma célula == "sim").
            if 0 <= idx_concluido < len(textos):
                concluido = textos[idx_concluido] == "sim"
            else:
                concluido = "sim" in textos
            if concluido:
                concluidas.append(row)
            else:
                pendentes.append(row)

        self._info("Agenda lida: %d linha(s) com dados.", linhas_com_dados)

        # ---- decisão ----
        if len(pendentes) == 1:
            # Negrito = não finalizado (regra do negócio): lido AGORA, ainda
            # na Agenda, pois após entrar a linha (row) fica obsoleta.
            em_negrito = self._linha_em_negrito(pendentes[0])
            return self._entrar_compromisso(numero, pendentes[0], em_negrito)

        if len(pendentes) >= 2:
            msg = (f"{len(pendentes)} compromissos '{self.config.tipo_compromisso_alvo}' "
                   f"pendentes (mais de um do mesmo tipo) — pulado")
            self._warn("%s [%s].", msg, numero)
            return ResultadoProcesso(numero, StatusProcesso.MULTIPLOS, msg)

        if concluidas:
            msg = f"'{self.config.tipo_compromisso_alvo}' já havia sido concluída"
            self._warn("%s [%s].", msg, numero)
            return ResultadoProcesso(numero, StatusProcesso.JA_CONCLUIDO, msg)

        msg = f"Nenhum compromisso '{self.config.tipo_compromisso_alvo}' encontrado"
        self._info("%s [%s].", msg, numero)
        if tipos_vistos:
            self._warn("Tipos lidos na Agenda: %s", " | ".join(tipos_vistos[:20]))
        else:
            self._warn("Nenhum texto lido nas linhas da Agenda "
                       "(tabela vazia, oculta ou estrutura diferente).")
        return ResultadoProcesso(numero, StatusProcesso.SEM_COMPROMISSO, msg)

    def _diagnostico_sem_tabela(self):
        """Loga pistas quando a tabela da Agenda não é encontrada."""
        try:
            self._warn("Diagnóstico — URL atual: %s", self.driver.current_url)
        except Exception:
            pass
        try:
            n_tabelas = len(self.driver.find_elements(By.TAG_NAME, "table"))
            n_tipo = len(self.driver.find_elements(
                By.XPATH, "//*[contains(normalize-space(.), 'Tipo de Compromisso')]"))
            n_agenda = len(self.driver.find_elements(*SEL_AGENDA_TOGGLE))
            self._warn(
                "Diagnóstico — %d <table>(s) | %d elem. com 'Tipo de Compromisso' "
                "| %d toggle(s) 'Agenda'.", n_tabelas, n_tipo, n_agenda)
        except Exception as e:
            self._warn("Diagnóstico falhou: %s", str(e))

    def _localizar_tabela_agenda(self):
        """Garante a Agenda expandida + linhas carregadas; retorna a tabela.

        A seção Agenda é um submódulo Bootstrap carregado de forma
        ASSÍNCRONA (data-async): o cabeçalho/tabela aparecem com atraso e,
        enquanto a seção está recolhida, a tabela existe no DOM mas com o
        <tbody> VAZIO — as linhas só chegam (via AJAX) DEPOIS de expandir.
        Por isso: (1) esperamos o cabeçalho renderizar e expandimos, e
        (2) esperamos o <tbody> popular antes de devolver a tabela.
        """
        self._expandir_agenda()
        return self._esperar_tabela_agenda_pronta()

    def _agenda_expandida(self, toggle) -> bool:
        """True se o container colapsável da Agenda está aberto (classe 'in').

        O toggle é <a data-toggle="collapse" href="#idContainer">; o
        container correspondente recebe a classe 'in' quando expandido.
        """
        try:
            href = toggle.get_attribute("href") or ""
            alvo = href.split("#", 1)[1] if "#" in href else ""
            if not alvo:
                return False
            cont = self.driver.find_element(By.ID, alvo)
            return "in" in (cont.get_attribute("class") or "").split()
        except WebDriverException:
            return False

    def _qtd_agenda(self) -> int | None:
        """Lê o contador do cabeçalho 'Agenda ( N )' (span.qtd-submodulo).

        Retorna N (compromissos esperados) ou None se não encontrar.
        """
        try:
            el = self.driver.find_element(*SEL_AGENDA_QTD)
            txt = re.sub(r"[^0-9]", "", el.get_attribute("textContent") or "")
            return int(txt) if txt else 0
        except (WebDriverException, ValueError):
            return None

    def _expandir_agenda(self):
        """Espera o cabeçalho 'Agenda' (submódulo data-async) e o expande.

        Espera a PRESENÇA do toggle (a seção renderiza de forma assíncrona)
        e só clica se a seção estiver recolhida — clicar numa já-aberta a
        fecharia.
        """
        try:
            toggle = WebDriverWait(self.driver, self.config.timeout).until(
                EC.presence_of_element_located(SEL_AGENDA_TOGGLE)
            )
        except TimeoutException:
            self._warn("Cabeçalho 'Agenda' não apareceu a tempo.")
            return
        if self._agenda_expandida(toggle):
            return
        self._info("Agenda recolhida — expandindo a seção...")
        try:
            self.driver.execute_script(
                "arguments[0].scrollIntoView({block:'center'});", toggle)
            self.driver.execute_script("arguments[0].click();", toggle)
        except Exception:
            pass
        # Espera o container colapsável abrir (ganhar a classe 'in').
        try:
            WebDriverWait(self.driver, self.config.timeout).until(
                lambda d: self._agenda_expandida(toggle)
            )
        except TimeoutException:
            self._warn("A seção Agenda não abriu no tempo esperado.")
        self._esperar_loading()

    def _esperar_tabela_agenda_pronta(self):
        """Espera a tabela da Agenda popular (AJAX) e a retorna.

        O conteúdo chega de forma assíncrona após expandir. Se o cabeçalho
        'Agenda ( N )' indicar 0 itens, devolve a tabela assim que ela ficar
        visível (sem esperar linhas que nunca virão). Quando há itens
        esperados, espera com MAIS paciência: em processos grandes (ex.:
        federais) o submódulo data-async pode passar dos 20s.
        """
        qtd = self._qtd_agenda()
        # Agenda vazia → espera curta; agenda COM itens (ou qtd desconhecida)
        # → espera generosa (no mínimo 45s), pois o AJAX federal é mais lento.
        espera = self.config.timeout if qtd == 0 else max(self.config.timeout, 45)
        fim = time.time() + espera
        tabela = None
        while time.time() < fim:
            tabela = self._tabela_agenda_visivel()
            if tabela is not None:
                if qtd == 0:
                    return tabela
                try:
                    # ".//tr[td]" (não só dentro de <tbody>): há layouts que
                    # renderizam as linhas fora do <tbody>.
                    if tabela.find_elements(By.XPATH, ".//tr[td]"):
                        return tabela
                except WebDriverException:
                    pass
            time.sleep(0.4)
        # Último recurso: tabela presente no DOM COM linhas de dados, mesmo que
        # o Selenium a julgue "não visível" (alguns layouts confundem o
        # is_displayed). Só aceita se TIVER linhas (evita pegar template vazio).
        if tabela is None:
            try:
                for e in self.driver.find_elements(*SEL_TABELA_AGENDA):
                    if e.find_elements(By.XPATH, ".//tr[td]"):
                        self._warn("Agenda lida via fallback (tabela no DOM, "
                                   "porém marcada como não-visível).")
                        return e
            except WebDriverException:
                pass
        return tabela  # última tentativa (pode estar vazia)

    def _tabela_agenda_visivel(self):
        """Retorna a tabela da Agenda se estiver VISÍVEL no DOM, senão None.

        Não-bloqueante (sem espera longa): é chamado dentro do laço de
        espera do AJAX em _esperar_tabela_agenda_pronta().
        """
        try:
            els = self.driver.find_elements(*SEL_TABELA_AGENDA)
        except WebDriverException:
            return None
        for e in els:
            try:
                if e.is_displayed():
                    return e
            except WebDriverException:
                continue
        return None  # existe no DOM mas oculta (colapsada) ou ainda não renderizou

    def _indice_coluna(self, tabela, alvo_norm: str) -> int:
        """Índice (0-based) da coluna cujo <th> casa `alvo_norm` (texto já
        normalizado). Usa textContent (ignora visibilidade). -1 se não achar.
        """
        try:
            ths = tabela.find_elements(By.XPATH, ".//thead//th")
        except WebDriverException:
            return -1
        for i, th in enumerate(ths):
            try:
                txt = normalizar_texto(th.get_attribute("textContent") or "")
            except WebDriverException:
                txt = ""
            if alvo_norm in txt:
                return i
        return -1

    def _entrar_compromisso(self, numero: str, row,
                            em_negrito: bool = False) -> ResultadoProcesso:
        """Entra no compromisso (clica na seta) e, se habilitado e o
        compromisso estiver em negrito, exclui o WorkFlow."""
        self._info("Compromisso '%s' pendente — entrando via seta...",
                   self.config.tipo_compromisso_alvo)

        seta = None
        for sel_rel in SEL_REL_SETA_CANDIDATOS:
            try:
                cand = row.find_elements(By.XPATH, sel_rel)
            except WebDriverException:
                cand = []
            if cand:
                seta = cand[0]
                break

        if seta is None:
            msg = "Seta para abrir o compromisso não encontrada na linha"
            self._err("%s [%s].", msg, numero)
            return ResultadoProcesso(numero, StatusProcesso.ERRO, msg)

        janelas_antes = set(self.driver.window_handles)
        try:
            self.driver.execute_script(
                "arguments[0].scrollIntoView({block:'center'});", seta)
            seta.click()
        except ElementClickInterceptedException:
            self.driver.execute_script("arguments[0].click();", seta)
        except Exception as e:
            self._err("Erro ao clicar na seta: %s", str(e))
            return ResultadoProcesso(numero, StatusProcesso.ERRO, f"Click seta: {e}")

        time.sleep(1)

        # Se abriu em nova aba, registra e fecha (volta pra aba principal),
        # para que o próximo processo possa ser pesquisado normalmente.
        janelas_depois = set(self.driver.window_handles)
        novas = janelas_depois - janelas_antes
        if novas:
            self._info("Compromisso abriu em nova aba — registrando e fechando.")
            principais = list(janelas_antes & janelas_depois)
            try:
                self.driver.switch_to.window(list(novas)[0])
                time.sleep(0.5)
                self.driver.close()
                destino = principais[0] if principais else self.driver.window_handles[0]
                self.driver.switch_to.window(destino)
            except Exception:
                try:
                    self.driver.switch_to.window(self.driver.window_handles[0])
                except Exception:
                    pass

        self._esperar_loading()
        self._ok("Entrou no compromisso '%s' do processo %s.",
                 self.config.tipo_compromisso_alvo, numero)

        # ---- Exclusão do WorkFlow (se habilitada) ----
        if not self.config.excluir_workflow:
            return ResultadoProcesso(
                numero, StatusProcesso.ENTROU,
                "Entrou no compromisso (exclusão de WorkFlow desligada)")
        # SALVAGUARDA: só exclui se o compromisso estiver em NEGRITO
        # (regra do negócio: negrito = não finalizado).
        if not em_negrito:
            msg = ("Entrou, mas o compromisso NÃO está em negrito "
                   "(pode já ter sido finalizado) — WorkFlow não excluído")
            self._warn("%s [%s].", msg, numero)
            return ResultadoProcesso(numero, StatusProcesso.ENTROU, msg)
        if self._excluir_workflow(numero):
            # Após excluir, reinicia o WorkFlow (botão 'Iniciar Workflow',
            # que aparece quando o fluxo fica vazio).
            reiniciou = self._iniciar_workflow(numero)
            msg = ("Entrou, excluiu e reiniciou o WorkFlow" if reiniciou
                   else "Entrou e excluiu o WorkFlow (falha ao reiniciar)")
            return ResultadoProcesso(numero, StatusProcesso.WORKFLOW_REINICIADO, msg)
        return ResultadoProcesso(numero, StatusProcesso.ERRO,
                                 "Entrou, mas falhou ao excluir o WorkFlow")

    def _linha_em_negrito(self, row) -> bool:
        """True se a linha do compromisso está em NEGRITO (fontWeight >= 600).

        Regra do negócio: na Agenda, compromisso em negrito = NÃO finalizado.
        Verifica a própria <tr> e suas primeiras células (qualquer uma em
        negrito já conta).
        """
        try:
            elementos = [row] + row.find_elements(By.TAG_NAME, "td")[:6]
        except WebDriverException:
            elementos = [row]
        for el in elementos:
            try:
                fw = self.driver.execute_script(
                    "return getComputedStyle(arguments[0]).fontWeight;", el)
            except WebDriverException:
                continue
            if fw in ("bold", "bolder"):
                return True
            try:
                if int(fw) >= 600:
                    return True
            except (ValueError, TypeError):
                continue
        return False

    def _excluir_workflow(self, numero: str) -> bool:
        """Exclui o WorkFlow do compromisso (botão 'Excluir' do submódulo) e
        confirma o diálogo. Retorna True se a confirmação foi clicada.

        SEGURANÇA EM CAMADAS:
          - usa SOMENTE o <a data-idcomponente='botao-excluir-submodulo'
            data-tabela='historicoworkflow'> (btn-default, cinza);
          - antes de clicar, REJEITA o elemento se tiver 'btn-danger', id
            'excluir-processo' ou data-idcomponente diferente — ou seja,
            NUNCA clica no 'Excluir' VERMELHO que apaga o compromisso inteiro.
        """
        wait = WebDriverWait(self.driver, self.config.timeout)
        try:
            btn = wait.until(
                EC.presence_of_element_located(SEL_BTN_EXCLUIR_WORKFLOW))
        except TimeoutException:
            self._err("Botão 'Excluir' do WorkFlow não encontrado.")
            return False

        # Salvaguardas anti-vermelho (dupla checagem antes do clique).
        cls = btn.get_attribute("class") or ""
        idc = btn.get_attribute("data-idcomponente") or ""
        bid = btn.get_attribute("id") or ""
        if ("btn-danger" in cls or bid == "excluir-processo"
                or idc != "botao-excluir-submodulo"):
            self._err("ABORTADO por segurança: botão de exclusão suspeito "
                      "(classe=%r id=%r componente=%r).", cls, bid, idc)
            return False

        self._info("Excluindo o WorkFlow (botão do submódulo, não-vermelho)...")
        try:
            self.driver.execute_script(
                "arguments[0].scrollIntoView({block:'center'});", btn)
            self.driver.execute_script("arguments[0].click();", btn)
        except Exception as e:
            self._err("Erro ao clicar em 'Excluir' do WorkFlow: %s", str(e))
            return False

        # Modal de confirmação (#confirm) -> botão "Confirmar".
        try:
            confirmar = wait.until(
                EC.element_to_be_clickable(SEL_MODAL_BTN_CONFIRMAR))
        except TimeoutException:
            self._warn("Modal de confirmação não apareceu — nada excluído.")
            return False
        self._info("Confirmando a exclusão...")
        try:
            confirmar.click()
        except ElementClickInterceptedException:
            self.driver.execute_script("arguments[0].click();", confirmar)
        except Exception as e:
            self._err("Erro ao confirmar a exclusão: %s", str(e))
            return False

        self._esperar_loading()
        # Confirma a remoção esperando um sinal CLARO (até o timeout): o botão
        # 'Iniciar Workflow' surgir (só aparece com o fluxo vazio) OU a box do
        # fluxo sumir. Evita o falso "não consegui verificar" quando, na prática,
        # a exclusão deu certo — só demorou um instante a refletir na tela.
        if self._esperar_workflow_removido():
            self._ok("WorkFlow excluído do processo %s.", numero)
        else:
            self._warn("Confirmou a exclusão, mas não consegui verificar a "
                       "remoção do WorkFlow [%s].", numero)
        return True  # a confirmação foi clicada — tratamos como excluído

    def _esperar_workflow_removido(self) -> bool:
        """Espera (até ~timeout) um sinal CLARO de que o WorkFlow foi removido,
        checando periodicamente. Retorna True assim que detectar; False só se
        estourar o tempo (aí sim é um caso pra olhar)."""
        fim = time.time() + self.config.timeout
        while time.time() < fim:
            if self._workflow_sumiu():
                return True
            time.sleep(0.4)
        return False

    def _workflow_sumiu(self) -> bool:
        """Sinais de que o WorkFlow foi removido (qualquer um já basta):
          1. surgiu o botão 'Iniciar Workflow' (só aparece com o fluxo VAZIO); OU
          2. a 'box' do fluxo (worflow-box) não está mais visível.
        """
        try:
            # Sinal forte e confiável: o botão de iniciar só aparece quando o
            # fluxo zera — ou seja, a exclusão funcionou.
            if self.driver.find_elements(*SEL_BTN_INICIAR_WORKFLOW):
                return True
            boxes = self.driver.find_elements(
                By.XPATH,
                "//*[contains(@class,'worflow-box') "
                "or contains(@class,'workflow-box-externo')]")
            return not any(b.is_displayed() for b in boxes)
        except WebDriverException:
            return False

    def _iniciar_workflow(self, numero: str) -> bool:
        """Após a exclusão, clica em 'Iniciar Workflow' (id='iniciar-workflow')
        — botão que aparece quando o WorkFlow fica vazio — para começar um
        novo fluxo. Retorna True se conseguiu clicar.
        """
        wait = WebDriverWait(self.driver, self.config.timeout)
        try:
            btn = wait.until(
                EC.element_to_be_clickable(SEL_BTN_INICIAR_WORKFLOW))
        except TimeoutException:
            self._warn("Botão 'Iniciar Workflow' não apareceu após a exclusão.")
            return False
        self._info("Iniciando novo WorkFlow...")
        try:
            self.driver.execute_script(
                "arguments[0].scrollIntoView({block:'center'});", btn)
            btn.click()
        except ElementClickInterceptedException:
            self.driver.execute_script("arguments[0].click();", btn)
        except Exception as e:
            self._err("Erro ao clicar em 'Iniciar Workflow': %s", str(e))
            return False
        self._esperar_loading()
        self._ok("WorkFlow reiniciado no processo %s.", numero)
        return True

    def _esperar_loading(self):
        """Espera o modal de carregamento (#modalLoading) desaparecer."""
        try:
            WebDriverWait(self.driver, self.config.timeout).until(
                EC.invisibility_of_element_located(SEL_MODAL_LOADING)
            )
        except TimeoutException:
            self._warn("Modal de carregamento demorou a sumir; tentando mesmo assim.")
