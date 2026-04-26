# Ferias Automacao

Geracao automatica de Markdown + HTML do Planejamento de Ferias a partir de uma planilha Excel.

Stack: **PowerShell + Pandoc + Mermaid** (sem dependencia de Python).

![Screenshot da janela](docs/screenshot.png)

> **Usuario final:** consulte o **[Manual do Usuario (Word)](docs/MANUAL.docx?raw=true)**
> com o passo-a-passo completo (download -> extracao -> preencher planilha ->
> gerar relatorio -> troubleshooting).

---

## O que faz

1. Le a planilha `Ferias-template.xlsx` (aba `Ferias`)
2. Preenche um template Markdown com:
   - **Dashboard mensal** (qtd. de pessoas + squads + status geral)
   - **Cronograma detalhado** (linha por colaborador)
   - **Gantt** (Mermaid, gerado dinamicamente)
   - **Rodape de autoria** (`Criado e compilado por <Autor> em dd/MM/yyyy as HH:mm`)
3. Roda **Pandoc** para gerar o HTML estilizado (CSS embutido + script do Mermaid no `<head>`)
4. Os arquivos sao gerados sempre na **mesma pasta** (`results/` por padrao,
   ou outra que voce escolher na GUI), com nomes fixos. Cada execucao
   **sobrescreve** a anterior:

```
<pasta-de-saida>/
  Ferias.md     <- relatorio em Markdown
  Ferias.html   <- relatorio em HTML (CSS+Mermaid embutido)
  Ferias.xlsx   <- copia da planilha que gerou o relatorio
  Ferias.pdf    <- (opcional, com -Pdf)
```

> **GUI:** a pasta escolhida pelo usuario fica persistida em
> `HKCU:\Software\FeriasAutomacao\OutputDir`, entao da proxima vez a
> janela ja abre apontando pra ela.

---

## Estrutura

```
ferias-automacao/
  Gerar Relatorio.lnk     -> ATALHO: duplo-clique para abrir a janela
  Gerar Relatorio.bat     -> alternativa ao atalho (mesmo comportamento)
  gui.ps1                 -> janela WinForms (Ano + Planilha + Gerar)
  executar.ps1            -> script principal (uso CLI/automacao)
  Ferias-template.xlsx        -> planilha com os dados (editar aqui)
  template.md             -> template Markdown com placeholders
  README.md               -> este arquivo
  results/                -> pasta de saida (criada na 1a execucao)
  assets/
    style.css             -> CSS aplicado no HTML
    header.html           -> tag <script> do Mermaid no <head>
    mermaid.lua           -> filtro Pandoc: ```mermaid``` -> <div class="mermaid">
  bin/
    pandoc-installer.msi  -> instalador offline do Pandoc (auto-instala)
  assets/
    icon.ico              -> icone do app (palmeira em fundo azul)
    generate-icon.ps1     -> regenera o icon.ico (rodar so se mudar o design)
```

---

## Personalizar o icone

O `assets/icon.ico` e usado pelo Form (barra de titulo + taskbar) e pelo
atalho `Gerar Relatorio.lnk`. Pra trocar o design:

1. Editar `assets/generate-icon.ps1` (cores, formas, tamanhos)
2. Rodar: `.\assets\generate-icon.ps1`
3. Pra atualizar o icone do atalho, recriar o `.lnk` com a flag `-IconLocation`
   apontando pra `assets\icon.ico,0`

---

## Requisitos

- Windows 10 / 11
- PowerShell 5.1+ (ja vem no Windows)
- Pandoc (o script instala automaticamente)
- Modulo PowerShell `ImportExcel` (o script instala automaticamente)

---

## Modo facil (GUI)

1. Editar `Ferias-template.xlsx` na aba `Ferias`. Colunas:

| Coluna | Tipo | Exemplo |
|---|---|---|
| `Mes` | texto | `Janeiro`, `Fevereiro`, `Marco`, ... |
| `Colaborador` | texto | `Joao Silva` |
| `Squad` | texto | `Pedido/OMS - Integracao` |
| `Inicio` | data | `dd/MM/yyyy` |
| `Fim` | data | `dd/MM/yyyy` |
| `Dias` | inteiro | `15` |
| `Status` | texto | `Aprovada`, `Solicitada` ou `Planejada` |

2. Salvar e fechar a planilha
3. Duplo-clique em **`Gerar Relatorio.lnk`**
4. Janela abre com:
   - **Ano** (NumericUpDown, default = ano atual)
   - **Planilha** (caminho + botao Procurar)
   - **Pasta de saida** (caminho + botao Procurar; persistida em registry)
   - **Abrir HTML apos gerar** (checkbox marcado)
   - Botao **Gerar Relatorio** (atalho `Enter`)
   - Botao **Fechar** (atalho `Esc`)

> **Instancia unica:** o app so permite uma janela aberta por vez no mesmo PC.
> Se voce der duplo-clique no atalho com a janela ja aberta, a janela existente
> volta pra frente e uma mensagem avisa que ja esta em execucao.
5. Apertar `Gerar Relatorio` -> aguardar progress bar -> HTML abre automaticamente

---

## Modo CLI (avancado)

```powershell
.\executar.ps1                                  # ano atual, autor = $env:USERNAME
.\executar.ps1 -Ano 2027 -Autor "Carlos Seidl"
.\executar.ps1 -OpenAfter                       # abre o HTML depois de gerar
.\executar.ps1 -OutputDir C:\temp               # muda pasta de saida
```

Parametros completos:

| Parametro | Default | Descricao |
|---|---|---|
| `-XlsxPath` | `.\Ferias-template.xlsx` | Planilha de entrada (precisa ser o template oficial) |
| `-OutputDir` | `.\results` | Pasta de saida |
| `-Autor` | `$env:USERNAME` | Nome no rodape e no `<meta name="author">` |
| `-Ano` | `(Get-Date).Year` | Ano usado no titulo (h1, `<title>`, Gantt) |
| `-OpenAfter` | `$false` | Abre o HTML automaticamente apos gerar |
| `-Pdf` | `$false` | Gera tambem `Ferias.pdf` (compativel com SharePoint) |

> **Regra do template:** o app SO funciona com a planilha-template oficial.
> A validacao roda antes de qualquer leitura e checa **duas** coisas:
> 1. **Nome do arquivo:** precisa ser `Ferias-template.xlsx` (case-insensitive)
> 2. **Estrutura:** o workbook precisa ter as 5 abas
>    `Ferias`, `Squads`, `Integrantes`, `Status`, `Instrucoes`
>
> Qualquer arquivo que falhe em (1) ou (2) e rejeitado com mensagem clara
> tanto no CLI quanto na GUI (MessageBox).

---

## Auto-instalacao do Pandoc

Na primeira execucao, se o Pandoc nao estiver no PATH, o script tenta nesta ordem:

1. **`bin/pandoc-installer.msi`** (offline, instala per-user sem admin)
2. **winget** (`JohnMacFarlane.Pandoc`, requer internet)
3. Aborta com instrucoes para instalacao manual

---

## Como funciona

```
Ferias-template.xlsx
       |
       v
   executar.ps1 ---- le aba "Ferias", normaliza datas/tipos
       |
       v
   template.md ---- substitui <!-- DASHBOARD -->, <!-- CRONOGRAMA -->,
       |             <!-- GANTT -->, <!-- AUTOR --> e <!-- ANO -->
       v
   <pasta-de-saida>/Ferias.md
       |
       v
     Pandoc ---- --standalone --embed-resources
       |          --css assets/style.css
       |          --include-in-header assets/header.html
       |          --lua-filter assets/mermaid.lua
       v
   <pasta-de-saida>/Ferias.html  (HTML standalone)
   <pasta-de-saida>/Ferias.xlsx  (copia da planilha-fonte)
```

Mermaid renderiza o Gantt no navegador via CDN (`cdn.jsdelivr.net`) — precisa de internet ao **abrir** o HTML.

---

## Troubleshooting

| Problema | Solucao |
|---|---|
| `Planilha nao encontrada` | Conferir caminho de `Ferias-template.xlsx` |
| `Planilha invalida: '...'` | Usar a planilha-template oficial (nome `Ferias-template.xlsx` + as 5 abas Ferias/Squads/Integrantes/Status/Instrucoes) |
| `Abas faltando: ...` | Sua copia da planilha foi mexida e perdeu uma das abas. Pegar o template original do release e copiar seus dados pra ele |
| HTML sem Gantt | Verificar internet ao abrir o HTML (CDN do Mermaid) |
| `ImportExcel` falha ao instalar | Rodar PowerShell como admin: `Install-Module ImportExcel -Scope CurrentUser -Force` |
| Erro de execucao de script (politica) | `Set-ExecutionPolicy -Scope CurrentUser -ExecutionPolicy RemoteSigned` |
| Warning `Could not fetch resource ... certificate has unknown CA` | Esperado em rede corporativa com MITM. O HTML carrega o Mermaid no navegador, fora do proxy — funciona normal ao abrir |
| App parou de abrir / arquivo da pasta de instalacao foi deletado | Reparar a instalacao: rodar de novo o `.msi` do release (oferece "Repair") **ou** abrir um console e executar `msiexec /f "<caminho-do-FeriasAutomacao.msi>"`. Como alternativa, desinstalar pelo Painel de Controle e reinstalar |

---

## Melhorias futuras

Lista de ideias mapeadas mas nao priorizadas: **[TODO.md](TODO.md)**.

Item atual em estudo:

- **Self-healing real do MSI** — hoje o MSI nao auto-recupera arquivos
  deletados da pasta de instalacao (atalhos sao non-advertised). Workaround
  via `msiexec /f` esta documentado na tabela de Troubleshooting acima. Pra
  habilitar self-heal automatico precisaria de um launcher EXE + atalhos
  advertised. Detalhes e plano completo no [TODO.md](TODO.md).

---

## Licenca

[MIT](LICENSE)
