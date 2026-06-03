# Fluxograma — Reiniciador de Workflow - iiLex (RPA)

Fluxo da automação, do login até a geração dos relatórios.

```mermaid
flowchart TD
    A([Iniciar]) --> B[Login no iiLex<br/>usuario + senha]
    B --> C{Login OK?}
    C -->|Nao| ZF([Fim: erro de login])
    C -->|Sim| D[Ler planilha Excel<br/>auto-detecta a coluna<br/>Numero do Processo]
    D --> LOOP[/Proximo processo da lista/]

    LOOP --> RL{Passaram 30 min<br/>desde o login?}
    RL -->|Sim| RLX[Re-login automatico]
    RL -->|Nao| OPEN[Abrir tela de Contencioso]
    RLX --> OPEN

    OPEN --> SITE{iiLex no ar?}
    SITE -->|Nao, o site caiu| WAIT[Aguardar de 15s a 2 min<br/>e tentar de novo<br/>ate o site voltar]
    WAIT --> SITE
    SITE -->|Sim| SESS{Sessao expirou?}
    SESS -->|Sim| RLX
    SESS -->|Nao| SEARCH[Pesquisar o numero do processo]

    SEARCH --> EXACT[Abrir o processo<br/>de numero EXATO<br/>resolve colisao de prefixo]
    EXACT --> FOUND{Encontrou?}
    FOUND -->|Nao| O_NF[NAO ENCONTRADO]:::rev
    FOUND -->|Sim| AG[Abrir a Agenda e ESPERAR<br/>carregar por completo<br/>paciencia p/ federais lentos]
    AG --> MATCH[Achar 'Solicitacao de Subsidios'<br/>match EXATO ignora acento/caixa]
    MATCH --> DEC{Situacao do<br/>compromisso-alvo}

    DEC -->|Nenhum do tipo| O_SC[SEM COMPROMISSO]:::rev
    DEC -->|2 ou mais| O_MU[MULTIPLOS - avisa e pula]:::rev
    DEC -->|Ja concluido| O_JC[JA CONCLUIDO]:::rev
    DEC -->|1 pendente| ENTER[Entrar no compromisso<br/>clicar na seta]

    ENTER --> BOLD{Esta em negrito?<br/>= nao finalizado}
    BOLD -->|Nao| O_EN[ENTROU<br/>nao mexe no WorkFlow]:::ok
    BOLD -->|Sim| DEL[Excluir WorkFlow<br/>botao cinza, NUNCA o vermelho<br/>e confirmar a remocao]
    DEL --> NEW[Iniciar novo WorkFlow]
    NEW --> O_WR[WORKFLOW REINICIADO]:::ok

    %% Falha tecnica em qualquer etapa -> auto-retry (ate 2x)
    AG -.->|"falha tecnica<br/>ex.: agenda lenta"| RETRY{Ja tentou<br/>2 vezes?}
    RETRY -->|Nao| WAIT3[Esperar 3s e refazer<br/>o processo do zero]
    WAIT3 --> OPEN
    RETRY -->|Sim| O_ER[ERRO]:::rev

    O_NF --> SAVE
    O_SC --> SAVE
    O_MU --> SAVE
    O_JC --> SAVE
    O_EN --> SAVE
    O_WR --> SAVE
    O_ER --> SAVE
    SAVE[(Salvar checkpoint<br/>resultados + progresso)] --> MORE
    MORE{Tem mais<br/>processos?} -->|Sim| LOOP
    MORE -->|Nao| REP[Gerar relatorios]
    REP --> R1[/Relatorio completo/]
    REP --> R2[/Planilha CONCLUIDOS/]:::ok
    REP --> R3[/Planilha REVISAR/]:::rev
    R1 --> END([Fim])
    R2 --> END
    R3 --> END

    classDef ok fill:#C6EFCE,stroke:#2E7D32,color:#1B5E20;
    classDef rev fill:#FFEB9C,stroke:#B7791F,color:#7A5C00;
```

## Legenda dos resultados

| Resultado | Significado | Vai para |
|---|---|---|
| 🟢 **WORKFLOW REINICIADO** | Entrou no compromisso e reiniciou o WorkFlow | Concluídos |
| 🟢 **ENTROU** | Entrou, mas o compromisso já estava finalizado (não mexe no WF) | Concluídos |
| 🟡 **SEM COMPROMISSO** | O processo não tem o compromisso-alvo | Revisar |
| 🟡 **JÁ CONCLUÍDO** | O compromisso-alvo já havia sido concluído | Revisar |
| 🟡 **MÚLTIPLOS** | 2+ compromissos do mesmo tipo (avisa e pula) | Revisar |
| 🟡 **NÃO ENCONTRADO** | O processo não foi localizado no iiLex | Revisar |
| 🟡 **ERRO** | Falha técnica que **persistiu mesmo após re-tentar** | Revisar |

## Proteções para rodadas longas
- **Re-login proativo** a cada 30 min (a sessão nunca expira)
- **Re-login reativo** se a sessão cair mesmo assim
- **Resiliência a quedas**: se o iiLex sair do ar, espera (15s → até 2 min) e tenta até voltar
- **Paciência ao carregar**: espera a Agenda carregar por completo antes de decidir (processos federais/TRF demoram mais)
- **Auto-retry**: erro técnico transitório (ex.: agenda não carregou a tempo) é re-tentado **até 2x** antes de virar ERRO
- **Checkpoint**: dá para parar e retomar de onde parou — **preservando os resultados já obtidos** (contadores e erros não se perdem)

## Segurança (regras do negócio)
- Só reinicia o WorkFlow se o compromisso estiver **em negrito** (= não finalizado)
- Usa **somente** o botão "Excluir" cinza do submódulo — **nunca** o "Excluir" vermelho (que apagaria o compromisso inteiro)
- Match **EXATO** do tipo: "Solicitação de Subsídios - CEF" **não** é confundido com "Solicitação de Subsídios"
