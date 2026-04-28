# Changelog

Todas as mudancas notaveis do **Ferias Automacao** ficam aqui.

Formato baseado em [Keep a Changelog](https://keepachangelog.com/pt-BR/1.1.0/),
versionamento segue [SemVer](https://semver.org/lang/pt-BR/) com formato
estendido `MAJOR.MINOR.PATCH.HOTFIX`.

Para visao tecnica completa, ver [README.md](README.md). Para passo-a-passo
do usuario final, ver [docs/MANUAL.docx](docs/MANUAL.docx).

---

## [1.0.0.1] - 2026-04-27

Primeira release com a feature de **Controle de Vencimento de Ferias**.

### Adicionado

- **Controle de Vencimento de Ferias (CLT 12/24 meses)** -- nova secao no
  relatorio HTML, logo apos o Dashboard mensal, que destaca quem precisa
  tirar ferias antes de vencer. Tres sub-blocos coloridos:
  - **CRITICO** (vermelho) -- 2o vencimento proximo (<= 6 meses) ou ja
    vencido (em dobro pela CLT)
  - **ATENCAO** (laranja) -- 1o vencimento proximo ou ja atingido
  - **Dados incompletos** (cinza) -- pessoas sem `Data de inicio na AIR`
    preenchida na aba Integrantes
- **Aba `Integrantes` com 4 colunas:** alem de `Integrante` e `Squad`,
  agora tem `Data de inicio na AIR` (admissao) e `Data das ultimas ferias`
  (data fim do ultimo periodo tirado, vazio = nunca tirou). Veja
  [docs/MANUAL.docx](docs/MANUAL.docx) secao 2.4 pra como preencher.
- **Apresentacao executiva** atualizada com slides "One More Thing"
  estilo Apple keynote (`docs/Apresentacao-FeriasAutomacao.pptx`,
  16 slides).
- **Logo AI/R** no titulo principal do relatorio, em SVG inline vetorial,
  herda a cor do `h1` via `currentColor`. Substitui o emoji de palmeira
  da v1.0.0.

### Corrigido

- **Atalhos do MSI quebrados:** atalhos do Menu Iniciar/Desktop/SendTo
  apontavam pra `[%SystemRoot%]\System32\powershell.exe`. WiX 4 nao
  expande `%SystemRoot%` em `Target` de Shortcut nao-advertised, entao
  o Windows abria "Atalho Nao Encontrado". Trocado por
  `[System64Folder]WindowsPowerShell\v1.0\powershell.exe`.
- **Janela cortando controles em DPI alto:** `Form.Size = 580x400`
  setava o tamanho total (com borda); area util ficava ~360px, apertada.
  Trocado por `ClientSize = 580x405` + `AutoScaleMode = Dpi`.
- **Label "Pasta de saida" truncado verticalmente:** label tinha
  `Width=75`, texto nao cabia em DPI alto e quebrava em 2 linhas com
  `Height=22` mostrando so a primeira. Aumentado pra `Width=105`,
  textboxes/botoes movidos pra `x=120`.
- **Cabecalho duplicado no HTML:** Pandoc gerava automaticamente um
  `<header id="title-block-header">` com titulo + autor que duplicava
  o `# titulo` do `template.md`. Escondido via CSS (`display: none`),
  mantendo os metadados no `<head>`.
- **Manual desatualizado:** 6 paragrafos duplicados de regra de filtro
  (bug de idempotencia), referencia a "tela de selecao de features tipo
  arvore" (UI antiga, agora e InstallDir), "icone da palmeira" (agora
  e logo AI/R). Tudo limpo e atualizado.

### Modificado

- **UI do instalador:** trocada de `WixUI_FeatureTree` para
  `WixUI_InstallDir` -- agora mostra uma tela dedicada com botao
  "Change..." pra escolher a pasta de instalacao no fluxo principal,
  em vez de escondida atras de "Customize".
- **Pasta no Menu Iniciar:** renomeada de "Ferias Automacao" para
  **"AI R"**, com 4 atalhos (Gerar Relatorio, Manual do usuario, Pasta
  de instalacao, Desinstalar Ferias Automacao).
- **Atalho "Gerar Relatorio" na Area de Trabalho:** agora e **obrigatorio**
  (parte da feature `Main` com `AllowAbsent="no"`), sem opcao de desmarcar.
- **`installer/build.ps1`:** copia automaticamente `docs/MANUAL.docx` pra
  `installer/output/` ao final do build, garantindo que o pacote de
  distribuicao tem MSI + Manual juntos sem caca-arquivos.

### Removido

- **Checkbox "Iniciar Ferias Automacao agora" no fim do instalador:**
  4 abordagens diferentes (CustomAction tipo 18 com `.bat`, tipo 50 com
  PowerShell direto, `WixSilentExec` da Util ext, `cmd.exe /c`) falharam
  silenciosamente em fazer o app abrir ao clicar Finish. Removido
  temporariamente -- usuario abre manualmente pelo atalho da Area de
  Trabalho. Item rastreado em [TODO.md](TODO.md).

---

## [1.0.0] - 2026-04-25

Primeira release publica. Aplicativo desktop pra gerar relatorio
executivo de planejamento de ferias a partir de uma planilha Excel.

### Adicionado

- **Pipeline de geracao do relatorio:** le `Ferias-template.xlsx`,
  preenche template Markdown com dashboard mensal + cronograma detalhado
  + Gantt visual, e roda Pandoc pra gerar HTML standalone (CSS + Mermaid
  embutidos) + Markdown + copia da planilha + PDF opcional.
- **GUI WinForms:** janela com Ano (year picker estilo calendario),
  Planilha (caminho + Procurar), Pasta de saida (persistida em registry),
  checkbox "Abrir apos gerar" e checkbox "Gerar PDF tambem".
- **Single-instance lock:** so permite uma janela do app por PC.
  Tentou abrir duas vezes? A janela existente volta pra frente.
- **Validacao estrita do template:** app so aceita
  `Ferias-template.xlsx` com as 5 abas obrigatorias (`Ferias`, `Squads`,
  `Integrantes`, `Status`, `Instrucoes`). Qualquer outro arquivo e
  rejeitado com mensagem clara.
- **Auto-instalacao do Pandoc:** na primeira execucao instala via
  `bin/pandoc-installer.msi` (offline, per-user) ou winget. Modulo
  `ImportExcel` instalado via `Install-Module -Scope CurrentUser`.
- **Status geral por mes (Dashboard):** Atencao (5+ pessoas),
  Em Progresso (alguem Solicitada), Alinhado (resto).
- **Filtro estrito por ano:** o relatorio so traz ferias que comecam
  no ano selecionado no picker, mesmo que a planilha tenha varios anos.
- **Mensagens de erro claras** em todos os caminhos de falha (planilha
  nao encontrada, sem permissao na pasta de saida, Pandoc faltando, etc).
- **Instalador MSI** (WiX 4, per-user, sem prompt de UAC) com atalhos
  no Menu Iniciar, Area de Trabalho e SendTo.
- **Manual do usuario completo** em [docs/MANUAL.docx](docs/MANUAL.docx)
  -- passo-a-passo do download ate gerar o relatorio.

---

## Comparando versoes

- [v1.0.0...v1.0.0.1](https://github.com/ceseidl/FeriasAutomacao/compare/v1.0.0...v1.0.0.1)

## Convencoes

Os commits seguem o formato [Conventional Commits](https://www.conventionalcommits.org/pt-br/v1.0.0/):

| Tipo | Quando usar |
|---|---|
| `feat` | feature nova |
| `fix` | correcao de bug |
| `revert` | desfaz uma mudanca anterior |
| `docs` | mudanca de documentacao |
| `chore` | build, deps, configs |
| `release` | promocao de versao |

Versionamento estendido `MAJOR.MINOR.PATCH.HOTFIX` -- o quarto numero
permite hotfixes sem disputar com a proxima `PATCH` (v1.0.0.1 ainda
e a primeira `1.0.0`, mas com hotfixes acumulados ate virar release
oficial).
