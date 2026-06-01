# 🤖 Exclusor de WorkFlow — Automação iiLex (RPA)

![Python](https://img.shields.io/badge/Python-3.10+-3776AB?logo=python&logoColor=white)
![Selenium](https://img.shields.io/badge/Selenium-4.x-43B02A?logo=selenium&logoColor=white)
![PyQt6](https://img.shields.io/badge/PyQt6-GUI-41CD52?logo=qt&logoColor=white)
![Platform](https://img.shields.io/badge/Windows-0078D6?logo=windows&logoColor=white)
![Portfolio](https://img.shields.io/badge/projeto-portf%C3%B3lio-blueviolet)

Aplicação **desktop de automação de processos (RPA)** que gerencia *workflows* no
sistema jurídico **iiLex**, substituindo um fluxo manual e repetitivo — abrir
processo por processo, ler a agenda, excluir e reiniciar o workflow — por uma
execução **automática, auditável e retomável** a partir de uma planilha Excel.

> ⚙️ **Nota de portfólio:** projeto desenvolvido em contexto profissional e
> **generalizado para fins de portfólio** — identificadores e dados da empresa
> foram removidos. As URLs de acesso ao iiLex são **configuráveis** (veja
> [Configuração](#%EF%B8%8F-configuração)).

---

## 🎯 O problema que resolve

O time jurídico precisava, para **centenas de processos**, repetir manualmente o
mesmo ciclo no iiLex: pesquisar o processo → abrir a agenda → localizar o
compromisso certo → entrar → excluir o workflow → reiniciá-lo. Minutos por
processo, horas no total, e sujeito a erro humano.

Esta ferramenta recebe uma **planilha com os números dos processos** e executa
todo o ciclo sozinha, gerando um **relatório auditável** ao final.

## ✨ Funcionalidades

- 🔐 **Login automático** e navegação no iiLex via Selenium
- 📑 **Leitura inteligente da Agenda**, incluindo o tratamento de tabelas
  carregadas de forma **assíncrona (AJAX / lazy-load)** — o ponto técnico mais
  difícil do projeto (veja [Destaques técnicos](#-destaques-técnicos))
- 🎯 **Match exato** do tipo de compromisso (normalização de acentos/caixa)
- 🔄 **Exclusão + reinício de workflows** em lote, com **salvaguardas** para
  jamais excluir o registro errado
- 🖥️ **Interface gráfica em PyQt6** — tema claro/escuro, logs coloridos em tempo
  real, barra de progresso e contadores por status
- 📊 **Relatório em Excel** ao final + **planilha separada de pendências**
  (não encontrados / erros / pulados) para reprocessamento rápido
- 💾 **Checkpoint** automático: retoma execuções interrompidas de onde parou
- 📦 **Empacotável como `.exe`** via PyInstaller

## 🖼️ Demonstração

> 📷 _Adicione aqui um print da interface em `docs/screenshot.png`._

![Interface da aplicação](docs/screenshot.png)

## 🏗️ Arquitetura

O projeto separa **lógica de automação** de **interface**, o que facilita testes,
manutenção e reuso (a mesma lógica poderia ser usada por uma CLI ou agendador):

```
core_iilex.py          → TODA a lógica Selenium (login, busca, leitura da
                          agenda, exclusão/reinício do workflow). Não depende
                          da UI: comunica-se por callbacks.
interface_iilex_qt.py  → GUI em PyQt6. Roda a automação em uma QThread para
                          não travar a interface; recebe eventos via callbacks.
config.py              → configurações persistentes (URLs, timeouts, opções)
credentials.py         → credenciais com criptografia DPAPI no Windows
checkpoint.py          → salva/retoma o progresso da execução
report.py              → geração dos relatórios em Excel (openpyxl)
```

A comunicação núcleo → UI é feita por uma estrutura de **callbacks**
(`on_log`, `on_progress`, `on_resultado`, `is_stop_requested`...), mantendo o
núcleo **agnóstico de interface**.

## 🛠️ Stack

`Python` · `Selenium` · `PyQt6` · `openpyxl` · `webdriver-manager` ·
`python-dotenv` · `pywin32 (DPAPI)` · `PyInstaller` · `Git`

Técnicas: Web scraping · Automação de navegador · XPath · análise de DOM
(Bootstrap) · threading · empacotamento de aplicação desktop.

## 🚀 Como rodar

```bash
# 1. Clone o repositório
git clone https://github.com/DiegoNoschang/exclusor-workflow-iilex.git
cd exclusor-workflow-iilex

# 2. (Opcional) crie um ambiente virtual
python -m venv .venv && .venv\Scripts\activate   # Windows

# 3. Instale as dependências
pip install -r requirements.txt

# 4. Configure as URLs do seu iiLex (veja a seção Configuração)

# 5. Execute
python interface_iilex_qt.py
```

> Requer **Google Chrome** instalado (o `webdriver-manager` baixa o driver
> compatível automaticamente).

## ⚙️ Configuração

As URLs do iiLex são placeholders — ajuste para a sua instância em `config.py`
(ou pela tela de **Configurações** na própria interface):

```python
url_login:       "https://SEU-DOMINIO.iilex.com.br/sistema/login/semacesso"
url_contencioso: "https://SEU-DOMINIO.iilex.com.br/sistema/contencioso/filtro"
```

## 🔍 Destaques técnicos

Alguns problemas reais resolvidos no projeto:

- **Tabela carregada via AJAX (`data-async`)** — a seção "Agenda" é um submódulo
  Bootstrap que renderiza com atraso e popula as linhas via AJAX **somente após
  ser expandida**. A primeira versão lia a tabela cedo demais e encontrava
  "0 linhas". A solução espera o cabeçalho renderizar, expande o painel correto
  e **aguarda o `<tbody>` popular** (usando o contador de itens do próprio
  cabeçalho) antes de ler.
- **Ação destrutiva com salvaguardas em camadas** — antes de excluir um
  workflow, o sistema exige que o compromisso esteja sinalizado como pendente,
  valida o **elemento exato** do botão (rejeitando botões parecidos que
  excluiriam o registro inteiro) e só então **confirma** no modal.
- **UI responsiva** — a automação roda em uma `QThread` separada, com `stop`/
  `pause` cooperativos, mantendo a interface fluida e cancelável.
- **Resiliência** — re-login automático em caso de sessão expirada, e
  *checkpoint* para retomar lotes longos interrompidos.

## 📁 Estrutura

```
exclusor-workflow-iilex/
├── core_iilex.py          # núcleo da automação (Selenium)
├── interface_iilex_qt.py  # interface gráfica (PyQt6)
├── config.py              # configurações persistentes
├── credentials.py         # credenciais (DPAPI no Windows)
├── checkpoint.py          # retomada de execução
├── report.py              # relatórios em Excel
├── requirements.txt
├── iilex.spec             # build do .exe (PyInstaller)
└── README.md
```

## 👤 Autor

**Diego Fernandes Noschang** — Estudante de Análise e Desenvolvimento de
Sistemas, com foco em automação de processos (RPA) e Python.

- 💻 GitHub: [@DiegoNoschang](https://github.com/DiegoNoschang)
- 💼 LinkedIn: [diego-fernandes-noschang](https://www.linkedin.com/in/diego-fernandes-noschang-379411300/)
- 📩 dnoschang.27@gmail.com

## 📄 Sobre este repositório

Projeto desenvolvido em contexto profissional e publicado **com autorização**,
em versão **generalizada para portfólio** — sem dados, credenciais ou
identificadores da empresa. Disponibilizado para fins de avaliação e
demonstração de habilidades técnicas.
