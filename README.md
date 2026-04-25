# Ferias Automacao

Geracao automatica de Markdown + HTML do Planejamento de Ferias a partir de uma planilha Excel.

Stack: **PowerShell + Pandoc + Mermaid** (sem dependencia de Python).

![Screenshot da janela](docs/screenshot.png)

---

## O que faz

1. Le a planilha `ferias-2026.xlsx` (aba `Ferias`)
2. Preenche um template Markdown com:
   - **Dashboard mensal** (qtd. de pessoas + squads + status geral)
   - **Cronograma detalhado** (linha por colaborador)
   - **Gantt** (Mermaid, gerado dinamicamente)
   - **Rodape de autoria** (`Criado e compilado por <Autor> em dd/MM/yyyy as HH:mm`)
3. Roda **Pandoc** para gerar o HTML estilizado (CSS embutido + script do Mermaid no `<head>`)
4. Cada execucao gera um arquivo novo com timestamp, preservando historico:

```
results/Ferias-yyyyMMdd-HHmmss.md
results/Ferias-yyyyMMdd-HHmmss.html
```

---

## Estrutura

```
ferias-automacao/
  Gerar Relatorio.lnk     -> ATALHO: duplo-clique para abrir a janela
  Gerar Relatorio.bat     -> alternativa ao atalho (mesmo comportamento)
  gui.ps1                 -> janela WinForms (Ano + Planilha + Gerar)
  executar.ps1            -> script principal (uso CLI/automacao)
  ferias-2026.xlsx        -> planilha com os dados (editar aqui)
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

1. Editar `ferias-2026.xlsx` na aba `Ferias`. Colunas:

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
   - **Abrir HTML apos gerar** (checkbox marcado)
   - Botao **Gerar Relatorio** (atalho `Enter`)
   - Botao **Fechar** (atalho `Esc`)
5. Apertar `Gerar Relatorio` -> aguardar progress bar -> HTML abre automaticamente

---

## Modo CLI (avancado)

```powershell
.\executar.ps1                                  # ano atual, autor = $env:USERNAME
.\executar.ps1 -Ano 2027 -Autor "Carlos Seidl"
.\executar.ps1 -OpenAfter                       # abre o HTML depois de gerar
.\executar.ps1 -CsvPath .\ferias.csv            # usa CSV (delimitador ';')
.\executar.ps1 -OutputDir C:\temp               # muda pasta de saida
```

Parametros completos:

| Parametro | Default | Descricao |
|---|---|---|
| `-XlsxPath` | `.\ferias-2026.xlsx` | Planilha de entrada |
| `-CsvPath` | (vazio) | Alternativa ao xlsx, separador `;` |
| `-OutputDir` | `.\results` | Pasta de saida |
| `-Autor` | `$env:USERNAME` | Nome no rodape e no `<meta name="author">` |
| `-Ano` | `(Get-Date).Year` | Ano usado no titulo (h1, `<title>`, Gantt) |
| `-OpenAfter` | `$false` | Abre o HTML automaticamente apos gerar |

---

## Auto-instalacao do Pandoc

Na primeira execucao, se o Pandoc nao estiver no PATH, o script tenta nesta ordem:

1. **`bin/pandoc-installer.msi`** (offline, instala per-user sem admin)
2. **winget** (`JohnMacFarlane.Pandoc`, requer internet)
3. Aborta com instrucoes para instalacao manual

---

## Como funciona

```
ferias-2026.xlsx
       |
       v
   executar.ps1 ---- le aba "Ferias", normaliza datas/tipos
       |
       v
   template.md ---- substitui <!-- DASHBOARD -->, <!-- CRONOGRAMA -->,
       |             <!-- GANTT -->, <!-- AUTOR --> e <!-- ANO -->
       v
   results/Ferias-{timestamp}.md
       |
       v
     Pandoc ---- --standalone --embed-resources
       |          --css assets/style.css
       |          --include-in-header assets/header.html
       |          --lua-filter assets/mermaid.lua
       v
   results/Ferias-{timestamp}.html  (HTML standalone com CSS embutido)
```

Mermaid renderiza o Gantt no navegador via CDN (`cdn.jsdelivr.net`) — precisa de internet ao **abrir** o HTML.

---

## Troubleshooting

| Problema | Solucao |
|---|---|
| `Planilha nao encontrada` | Conferir caminho de `ferias-2026.xlsx` |
| Acentos quebrados | Garantir que a planilha esta salva em UTF-8 |
| HTML sem Gantt | Verificar internet ao abrir o HTML (CDN do Mermaid) |
| `ImportExcel` falha ao instalar | Rodar PowerShell como admin: `Install-Module ImportExcel -Scope CurrentUser -Force` |
| Erro de execucao de script (politica) | `Set-ExecutionPolicy -Scope CurrentUser -ExecutionPolicy RemoteSigned` |
| Warning `Could not fetch resource ... certificate has unknown CA` | Esperado em rede corporativa com MITM. O HTML carrega o Mermaid no navegador, fora do proxy — funciona normal ao abrir |

---

## Licenca

[MIT](LICENSE)
