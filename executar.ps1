#requires -Version 5.1
<#
.SYNOPSIS
    Gera Ferias.md, Ferias.html e Ferias.xlsx (snapshot da planilha) na
    pasta de saida configurada. Cada execucao sobrescreve os anteriores.

.DESCRIPTION
    Le a planilha (aba "Ferias"), preenche o template.md com dashboard + cronograma + gantt,
    e roda pandoc pra gerar o HTML estilizado (CSS + Mermaid embutido).
    A pasta de saida e fixa entre execucoes (default: .\results, ou o que a GUI
    salvou no registry). Os arquivos tem nomes fixos e sao sobrescritos a cada
    execucao - assim a pasta nao acumula historico:
        $OutputDir\Ferias.md     <- relatorio em Markdown
        $OutputDir\Ferias.html   <- relatorio em HTML (CSS+Mermaid embutido)
        $OutputDir\Ferias.xlsx   <- copia da planilha-fonte usada
        $OutputDir\Ferias.pdf    <- (opcional, com -Pdf)

.PARAMETER XlsxPath
    Caminho para a planilha. Default: .\Ferias-template.xlsx
    REGRA: o app SO funciona com a planilha-template oficial
    (Ferias-template.xlsx, com as 5 abas Ferias/Squads/Integrantes/Status/
    Instrucoes). Qualquer outro arquivo e rejeitado com erro claro.

.PARAMETER OutputDir
    Pasta de saida para os arquivos gerados. Default: .\results (criada automaticamente).

.PARAMETER Autor
    Nome a ser exibido no rodape do HTML e no metadado <meta name="author">.
    Default: $env:USERNAME.

.PARAMETER Ano
    Ano do planejamento (aparece no titulo do HTML, do Gantt e no <title>).
    Default: ano atual.

.PARAMETER Pdf
    Gera tambem uma versao PDF (Ferias-{timestamp}.pdf) usando Edge/Chrome
    em modo headless. Necessario para SharePoint, que faz preview nativo de
    PDF mas pode bloquear ou forcar download de HTML com scripts externos.

.EXAMPLE
    .\executar.ps1                                  # ano atual, autor = usuario do Windows
    .\executar.ps1 -Ano 2027 -Autor "Carlos Seidl"
    .\executar.ps1 -OpenAfter                       # abre o HTML depois de gerar
    .\executar.ps1 -Pdf                             # gera HTML + PDF (SharePoint)
#>
[CmdletBinding()]
param(
    [string]$XlsxPath = (Join-Path $PSScriptRoot 'Ferias-template.xlsx'),
    [string]$OutputDir = (Join-Path $PSScriptRoot 'results'),
    [string]$TemplatePath = (Join-Path $PSScriptRoot 'template.md'),
    [string]$CssPath = (Join-Path $PSScriptRoot 'assets\style.css'),
    [string]$HeaderPath = (Join-Path $PSScriptRoot 'assets\header.html'),
    [string]$LuaFilter = (Join-Path $PSScriptRoot 'assets\mermaid.lua'),
    [string]$Autor = $env:USERNAME,
    [int]$Ano = (Get-Date).Year,
    [switch]$OpenAfter,
    [switch]$Pdf
)

$ErrorActionPreference = 'Stop'
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8

# ============================================================
# Template oficial - regra de validacao
# ------------------------------------------------------------
# O app SO funciona com a planilha-template oficial. Validamos:
#   1. Nome do arquivo (case-insensitive) = Ferias-template.xlsx
#   2. Workbook contem TODAS as 5 abas esperadas
# Qualquer arquivo que falhe em (1) ou (2) e rejeitado com
# mensagem clara. Isso garante que o pipeline (dropdowns,
# filtros por ano, gantt) sempre encontre os dados onde espera.
# ============================================================
$TEMPLATE_FILENAME       = 'Ferias-template.xlsx'
$TEMPLATE_REQUIRED_SHEETS = @('Ferias', 'Squads', 'Integrantes', 'Status', 'Instrucoes')

function Test-FeriasTemplate {
    param(
        [Parameter(Mandatory)] [string]$Path
    )

    if (-not (Test-Path $Path)) {
        throw "Planilha nao encontrada: $Path"
    }

    $actualName = Split-Path -Leaf $Path
    if ($actualName -ne $TEMPLATE_FILENAME) {
        throw @"
Planilha invalida: '$actualName'.
O app so funciona com a planilha-template oficial '$TEMPLATE_FILENAME'.
Use o template original que vem junto com o app (ou renomeie sua copia
pra esse nome, mantendo a estrutura de 5 abas).
"@
    }

    if (-not (Get-Module -ListAvailable -Name ImportExcel)) {
        Write-Host "Modulo ImportExcel nao encontrado. Instalando no escopo do usuario..." -ForegroundColor Yellow
        Install-Module ImportExcel -Scope CurrentUser -Force -ErrorAction Stop
    }
    Import-Module ImportExcel -ErrorAction Stop | Out-Null

    $actualSheets = @(Get-ExcelSheetInfo -Path $Path | ForEach-Object { $_.Name })
    $missing = @($TEMPLATE_REQUIRED_SHEETS | Where-Object { $_ -notin $actualSheets })
    if ($missing.Count -gt 0) {
        throw @"
Planilha invalida: '$actualName' nao tem a estrutura esperada.
Abas faltando: $($missing -join ', ').
O app so funciona com a planilha-template oficial, que precisa conter
todas as 5 abas: $($TEMPLATE_REQUIRED_SHEETS -join ', ').
"@
    }
}

function Update-PathFromEnvironment {
    $machinePath = [System.Environment]::GetEnvironmentVariable('Path','Machine')
    $userPath    = [System.Environment]::GetEnvironmentVariable('Path','User')
    $env:Path = ($machinePath, $userPath, 'C:\Program Files\Pandoc') -join ';'
}

function Install-PandocFromMsi([string]$MsiPath) {
    Write-Host "Rodando instalador MSI bundled: $MsiPath" -ForegroundColor Yellow
    $logPath = Join-Path $env:TEMP "pandoc-install.log"
    # Tenta per-user (sem admin). Se falhar, cai pra per-machine (que pode exigir UAC).
    $argsPerUser = "/i `"$MsiPath`" /qn /norestart MSIINSTALLPERUSER=1 ALLUSERS=2 /l*v `"$logPath`""
    $p = Start-Process msiexec.exe -ArgumentList $argsPerUser -Wait -PassThru -NoNewWindow
    if ($p.ExitCode -ne 0) {
        Write-Host "Install per-user retornou $($p.ExitCode). Tentando per-machine..." -ForegroundColor DarkYellow
        $argsAll = "/i `"$MsiPath`" /qn /norestart /l*v `"$logPath`""
        $p = Start-Process msiexec.exe -ArgumentList $argsAll -Wait -PassThru -Verb RunAs
    }
    return $p.ExitCode
}

function Ensure-Pandoc {
    Update-PathFromEnvironment
    if (Get-Command pandoc -ErrorAction SilentlyContinue) { return }

    Write-Host "Pandoc nao encontrado no PATH." -ForegroundColor Yellow

    # 1. Instalador MSI bundled (prioridade: funciona offline)
    $bundledMsi = Join-Path $PSScriptRoot 'bin\pandoc-installer.msi'
    if (Test-Path $bundledMsi) {
        $code = Install-PandocFromMsi -MsiPath $bundledMsi
        Update-PathFromEnvironment
        if (Get-Command pandoc -ErrorAction SilentlyContinue) {
            Write-Host "Pandoc instalado via MSI bundled." -ForegroundColor Green
            return
        }
        Write-Host "MSI bundled nao resolveu (exit $code). Tentando winget..." -ForegroundColor DarkYellow
    } else {
        Write-Host "MSI bundled nao encontrado em bin/. Tentando winget..." -ForegroundColor DarkYellow
    }

    # 2. winget (requer internet + loja de pacotes)
    if (Get-Command winget -ErrorAction SilentlyContinue) {
        & winget install --id JohnMacFarlane.Pandoc -e --accept-source-agreements --accept-package-agreements --silent | Out-Null
        Update-PathFromEnvironment
        if (Get-Command pandoc -ErrorAction SilentlyContinue) {
            Write-Host "Pandoc instalado via winget." -ForegroundColor Green
            return
        }
    }

    throw @"
Pandoc nao ficou disponivel. Opcoes:
  1. Feche e reabra o PowerShell (pode ter instalado, mas o PATH da sessao atual esta defasado), e rode novamente.
  2. Coloque o instalador em bin\pandoc-installer.msi e rode novamente.
  3. Instale manualmente de https://pandoc.org/installing.html.
"@
}

function Find-EdgeOrChrome {
    $candidates = @(
        "${env:ProgramFiles}\Microsoft\Edge\Application\msedge.exe",
        "${env:ProgramFiles(x86)}\Microsoft\Edge\Application\msedge.exe",
        "${env:LocalAppData}\Microsoft\Edge\Application\msedge.exe",
        "${env:ProgramFiles}\Google\Chrome\Application\chrome.exe",
        "${env:ProgramFiles(x86)}\Google\Chrome\Application\chrome.exe",
        "${env:LocalAppData}\Google\Chrome\Application\chrome.exe"
    )
    return $candidates | Where-Object { Test-Path $_ } | Select-Object -First 1
}

function Convert-HtmlToPdf {
    param(
        [Parameter(Mandatory)][string]$HtmlPath,
        [Parameter(Mandatory)][string]$PdfPath
    )

    $browser = Find-EdgeOrChrome
    if (-not $browser) {
        throw 'Microsoft Edge ou Google Chrome nao encontrado. Necessario para gerar PDF.'
    }

    # IMPORTANTE: Edge headless interpreta mal paths com espacos quando passados
    # via Start-Process/call operator (OneDrive tem espaco no nome). Trabalhamos
    # tudo em $env:TEMP (sem espaco) e copiamos os arquivos pro destino final.
    $stamp     = [guid]::NewGuid().ToString('N').Substring(0, 8)
    $tempBase  = Join-Path $env:TEMP "ferias-pdf-$stamp"
    $tempUserDir = Join-Path $tempBase 'profile'
    $tempHtml  = Join-Path $tempBase 'input.html'
    $tempPdf   = Join-Path $tempBase 'output.pdf'
    New-Item -ItemType Directory -Path $tempUserDir -Force | Out-Null

    try {
        # Copia o HTML pra um caminho sem espacos pra evitar problemas de quoting
        Copy-Item -Path $HtmlPath -Destination $tempHtml -Force

        $tempHtmlUri = (New-Object System.Uri($tempHtml)).AbsoluteUri

        $browserArgs = @(
            '--headless=new',
            '--disable-gpu',
            '--no-pdf-header-footer',
            '--virtual-time-budget=15000',
            "--user-data-dir=$tempUserDir",
            "--print-to-pdf=$tempPdf",
            $tempHtmlUri
        )

        Write-Host "Gerando PDF via $browser ..."
        $proc = Start-Process -FilePath $browser -ArgumentList $browserArgs -Wait -PassThru -NoNewWindow

        if ($proc.ExitCode -ne 0) {
            throw "Browser headless retornou exit code $($proc.ExitCode)."
        }
        if (-not (Test-Path $tempPdf)) {
            throw "PDF nao foi criado em $tempPdf (Edge nao escreveu o arquivo)."
        }

        # Move pro destino final (que pode ter espaco/acento)
        Copy-Item -Path $tempPdf -Destination $PdfPath -Force
    }
    finally {
        Remove-Item -Path $tempBase -Recurse -Force -ErrorAction SilentlyContinue
    }
}

Ensure-Pandoc

$monthOrder = @('Janeiro','Fevereiro','Marco','Março','Abril','Maio','Junho','Julho','Agosto','Setembro','Outubro','Novembro','Dezembro')
$statusIcon = @{ 'Aprovada' = '🟢'; 'Solicitada' = '🟡'; 'Planejada' = '⚪' }

function Get-MonthIndex([string]$m) {
    $i = $monthOrder.IndexOf($m)
    if ($i -lt 0) { $i = $monthOrder.IndexOf($m.Replace('ç','c')) }
    if ($i -lt 0) { $i = 99 }
    return $i
}

function Read-XlsxRows([string]$Path) {
    # A validacao de Path + estrutura ja foi feita por Test-FeriasTemplate
    # antes desta funcao ser chamada. Aqui so faz o load dos dados da aba
    # Ferias (que sabemos que existe).
    Import-Module ImportExcel -ErrorAction Stop | Out-Null
    $data = Import-Excel -Path $Path -WorksheetName 'Ferias' -ErrorAction Stop
    return , $data
}

function Read-IntegrantesRows([string]$Path) {
    # Le a aba 'Integrantes' (cadastro mestre dos colaboradores). Usado pela
    # secao "Controle de Vencimento de Ferias" pra calcular 1o e 2o vencimento
    # CLT por pessoa. A aba sempre existe (validada por Test-FeriasTemplate),
    # mas pode estar com a estrutura antiga (sem as colunas Data de inicio
    # na AIR / Data das ultimas ferias) -- nesse caso a secao mostra um aviso.
    Import-Module ImportExcel -ErrorAction Stop | Out-Null
    $data = Import-Excel -Path $Path -WorksheetName 'Integrantes' -ErrorAction Stop
    return , $data
}

function ConvertTo-DateOrNull([object]$v) {
    # Converte valor de celula (datetime, OLE serial, string varias) pra
    # [datetime] ou $null se vazio/invalido. Usado pra normalizar as colunas
    # de data da aba Integrantes que podem vir em formatos diferentes
    # dependendo de como o usuario digitou.
    if ($null -eq $v -or "$v" -eq '') { return $null }
    if ($v -is [datetime]) { return $v }
    if ($v -is [double] -or $v -is [int]) {
        try { return [datetime]::FromOADate([double]$v) } catch { return $null }
    }
    $s = "$v".Trim()
    if ($s -eq '') { return $null }
    $dt = [datetime]::MinValue
    $formats = 'dd/MM/yyyy','yyyy-MM-dd','MM/dd/yyyy','dd-MM-yyyy','dd/MM/yy'
    if ([datetime]::TryParseExact($s, $formats, [System.Globalization.CultureInfo]::InvariantCulture, [System.Globalization.DateTimeStyles]::None, [ref]$dt)) {
        return $dt
    }
    if ([datetime]::TryParse($s, [ref]$dt)) { return $dt }
    return $null
}

function Format-Duration([int]$dias) {
    # Formata uma quantidade de dias como "X dia(s)" ou "~X mes(es)",
    # dependendo da magnitude. Negativo vira "ha", positivo vira "em".
    # Pluraliza corretamente (1 dia, 2 dias / 1 mes, 2 meses).
    $abs = [Math]::Abs($dias)
    if ($abs -le 30) {
        $u = if ($abs -eq 1) { 'dia' } else { 'dias' }
        if ($dias -lt 0) { return "ha $abs $u" } else { return "em $abs $u" }
    }
    $meses = [int][Math]::Round($abs / 30.44)
    if ($meses -le 0) { $meses = 1 }
    $u = if ($meses -eq 1) { 'mes' } else { 'meses' }
    if ($dias -lt 0) { return "ha $meses $u" } else { return "em $meses $u" }
}

function Build-VencimentosSection {
    # Monta o markdown da secao "Controle de Vencimento de Ferias" a partir
    # do cadastro de integrantes. Logica:
    #
    #   Para cada pessoa, determina o inicio do periodo aquisitivo atual:
    #     * Se "Data das ultimas ferias" preenchida -> usa essa data + 1 dia
    #       (ciclo aquisitivo novo comeca no dia seguinte ao fim das ferias)
    #     * Senao se "Data de inicio na AIR" preenchida -> usa essa data
    #       (pessoa nunca tirou ferias, ciclo conta da admissao)
    #     * Senao -> bloco "Dados incompletos"
    #
    #   1o vencimento = inicio + 12 meses (CLT: pessoa adquire o direito)
    #   2o vencimento = inicio + 24 meses (CLT: ferias vencem em dobro)
    #
    # Janela de alerta: 6 meses antes de cada vencimento.
    #
    # Classificacao (mais critico primeiro):
    #   * 2o vencimento <= hoje + 6 meses (incluindo passado) -> CRITICO
    #   * 1o vencimento <= hoje + 6 meses (incluindo passado) -> ATENCAO
    #   * Caso contrario -> nao aparece (em dia)
    #
    # Pessoa entra so no bloco mais critico aplicavel (sem duplicar).
    param(
        [object[]]$Integrantes,
        [datetime]$Hoje
    )

    $LIMITE_DIAS = 183  # ~6 meses

    # Detecta se a planilha esta na versao antiga (sem as colunas novas).
    # Se nao tiver as 2 colunas, todo mundo cai em dados incompletos e
    # mostramos um aviso no topo da secao.
    $temColunaAdm = $false
    $temColunaUlt = $false
    if ($Integrantes -and $Integrantes.Count -gt 0) {
        $cols = ($Integrantes[0] | Get-Member -MemberType NoteProperty).Name
        $temColunaAdm = $cols -contains 'Data de inicio na AIR'
        $temColunaUlt = $cols -contains 'Data das ultimas ferias'
    }

    $criticos    = @()
    $atencao     = @()
    $incompletos = @()

    foreach ($pessoa in $Integrantes) {
        $nome = "$($pessoa.Integrante)".Trim()
        if (-not $nome) { continue }
        $squad = "$($pessoa.Squad)".Trim()

        $dataAdm = if ($temColunaAdm) { ConvertTo-DateOrNull $pessoa.'Data de inicio na AIR' } else { $null }
        $dataUlt = if ($temColunaUlt) { ConvertTo-DateOrNull $pessoa.'Data das ultimas ferias' } else { $null }

        $referencia = $null
        $tipoRef    = ''
        if ($dataUlt) {
            $referencia = $dataUlt.AddDays(1)
            $tipoRef    = "Ult.ferias " + $dataUlt.ToString('dd/MM/yyyy')
        } elseif ($dataAdm) {
            $referencia = $dataAdm
            $tipoRef    = "Adm. " + $dataAdm.ToString('dd/MM/yyyy')
        } else {
            # Sem nenhuma data -> dados incompletos. So consideramos "Data de
            # admissao ausente" porque a coluna de ultimas ferias vazia eh
            # esperada pra quem nunca tirou ferias.
            $faltando = @()
            if (-not $temColunaAdm) { $faltando += 'Coluna "Data de inicio na AIR" ausente na planilha' }
            elseif (-not $dataAdm)  { $faltando += 'Data de admissao ausente' }
            if (-not $temColunaUlt) { $faltando += 'Coluna "Data das ultimas ferias" ausente na planilha' }
            $incompletos += [pscustomobject]@{
                Colaborador     = $nome
                Squad           = $squad
                ColunasAusentes = ($faltando -join '; ')
            }
            continue
        }

        $venc1 = $referencia.AddYears(1)
        $venc2 = $referencia.AddYears(2)
        $diasAteVenc1 = [int]([Math]::Floor(($venc1 - $Hoje).TotalDays))
        $diasAteVenc2 = [int]([Math]::Floor(($venc2 - $Hoje).TotalDays))

        $row = [pscustomobject]@{
            Colaborador = $nome
            Squad       = $squad
            Referencia  = $tipoRef
            Venc1       = $venc1.ToString('dd/MM/yyyy')
            Venc2       = $venc2.ToString('dd/MM/yyyy')
            DiasOrdem   = 0
            Situacao    = ''
        }

        if ($diasAteVenc2 -le $LIMITE_DIAS) {
            # 2o vencimento atingido ou dentro da janela -> CRITICO
            $row.DiasOrdem = $diasAteVenc2
            if ($diasAteVenc2 -lt 0) {
                $row.Situacao = "EM DOBRO " + (Format-Duration $diasAteVenc2)
            } else {
                $row.Situacao = "2o venc. " + (Format-Duration $diasAteVenc2)
            }
            $criticos += $row
        } elseif ($diasAteVenc1 -le $LIMITE_DIAS) {
            # 1o vencimento atingido ou dentro da janela -> ATENCAO
            $row.DiasOrdem = $diasAteVenc1
            if ($diasAteVenc1 -lt 0) {
                $row.Situacao = "1o venc. atingido " + (Format-Duration $diasAteVenc1)
            } else {
                $row.Situacao = "1o venc. " + (Format-Duration $diasAteVenc1)
            }
            $atencao += $row
        }
        # else: em dia -- nao aparece na secao
    }

    # Ordena cada bloco do mais critico (menor DiasOrdem) pro menos
    $criticos    = @($criticos    | Sort-Object DiasOrdem)
    $atencao     = @($atencao     | Sort-Object DiasOrdem)
    $incompletos = @($incompletos | Sort-Object Colaborador)

    # Monta o markdown da secao.
    $sb = New-Object System.Text.StringBuilder
    [void]$sb.AppendLine('## 🚨 Controle de Vencimento de Férias')
    [void]$sb.AppendLine('')
    [void]$sb.AppendLine("Referencia: hoje = **" + $Hoje.ToString('dd/MM/yyyy') + "**. Janela de alerta: **6 meses** antes de cada vencimento.")
    [void]$sb.AppendLine('')

    if (-not $temColunaAdm -or -not $temColunaUlt) {
        [void]$sb.AppendLine('> ⚠️ **Atualize a planilha.** A aba `Integrantes` nao tem as colunas')
        [void]$sb.AppendLine('> `Data de inicio na AIR` e/ou `Data das ultimas ferias`. Sem elas o app nao')
        [void]$sb.AppendLine('> consegue calcular os vencimentos. Baixe o template oficial atualizado e')
        [void]$sb.AppendLine('> copie seus dados pra ele.')
        [void]$sb.AppendLine('')
    }

    $algumBloco = $false

    if ($criticos.Count -gt 0) {
        $algumBloco = $true
        [void]$sb.AppendLine('### 🔴 CRITICO -- Ferias com 2o vencimento proximo ou vencido')
        [void]$sb.AppendLine('')
        [void]$sb.AppendLine('*Em dobro / urgente -- precisam tirar ferias agora.*')
        [void]$sb.AppendLine('')
        [void]$sb.AppendLine('| Colaborador | Squad | Ult. ferias / Adm. | 1o venc. | 2o venc. | Situacao |')
        [void]$sb.AppendLine('| :--- | :--- | :--- | :---: | :---: | :--- |')
        foreach ($r in $criticos) {
            [void]$sb.AppendLine("| **$($r.Colaborador)** | $($r.Squad) | $($r.Referencia) | $($r.Venc1) | **$($r.Venc2)** | **$($r.Situacao)** |")
        }
        [void]$sb.AppendLine('')
    }

    if ($atencao.Count -gt 0) {
        $algumBloco = $true
        [void]$sb.AppendLine('### 🟠 ATENCAO -- Ferias com 1o vencimento proximo ou atingido')
        [void]$sb.AppendLine('')
        [void]$sb.AppendLine('*Ja pode (e deve) comecar a tirar ferias.*')
        [void]$sb.AppendLine('')
        [void]$sb.AppendLine('| Colaborador | Squad | Ult. ferias / Adm. | 1o venc. | 2o venc. | Situacao |')
        [void]$sb.AppendLine('| :--- | :--- | :--- | :---: | :---: | :--- |')
        foreach ($r in $atencao) {
            [void]$sb.AppendLine("| **$($r.Colaborador)** | $($r.Squad) | $($r.Referencia) | **$($r.Venc1)** | $($r.Venc2) | $($r.Situacao) |")
        }
        [void]$sb.AppendLine('')
    }

    if ($incompletos.Count -gt 0) {
        $algumBloco = $true
        [void]$sb.AppendLine('### ⚪ Dados incompletos -- atualizar planilha')
        [void]$sb.AppendLine('')
        [void]$sb.AppendLine('*Colaboradores sem cadastro completo na aba `Integrantes`. Preencha as datas e regere o relatorio pra ver os vencimentos.*')
        [void]$sb.AppendLine('')
        [void]$sb.AppendLine('| Colaborador | Squad | O que falta |')
        [void]$sb.AppendLine('| :--- | :--- | :--- |')
        foreach ($r in $incompletos) {
            [void]$sb.AppendLine("| $($r.Colaborador) | $($r.Squad) | $($r.ColunasAusentes) |")
        }
        [void]$sb.AppendLine('')
    }

    if (-not $algumBloco) {
        [void]$sb.AppendLine('Todos os colaboradores estao em dia. Nenhum vencimento proximo ou atingido.')
        [void]$sb.AppendLine('')
    }

    return $sb.ToString().TrimEnd()
}

function Format-DateBr([object]$v) {
    if ($null -eq $v -or "$v" -eq '') { return '' }
    if ($v -is [datetime]) { return $v.ToString('dd/MM/yyyy') }
    $s = "$v".Trim()
    $dt = [datetime]::MinValue
    $formats = 'dd/MM/yyyy','yyyy-MM-dd','MM/dd/yyyy','dd-MM-yyyy'
    if ([datetime]::TryParseExact($s, $formats, [System.Globalization.CultureInfo]::InvariantCulture, [System.Globalization.DateTimeStyles]::None, [ref]$dt)) {
        return $dt.ToString('dd/MM/yyyy')
    }
    if ([datetime]::TryParse($s, [ref]$dt)) { return $dt.ToString('dd/MM/yyyy') }
    return $s
}

function ConvertTo-IsoDate([string]$br) {
    $dt = [datetime]::MinValue
    if ([datetime]::TryParseExact($br, 'dd/MM/yyyy', [System.Globalization.CultureInfo]::InvariantCulture, [System.Globalization.DateTimeStyles]::None, [ref]$dt)) {
        return $dt.ToString('yyyy-MM-dd')
    }
    return $br
}

# 1. Le os dados
# Antes de qualquer leitura, valida que e a planilha-template oficial:
#   - nome do arquivo = Ferias-template.xlsx
#   - workbook tem as 5 abas esperadas
# Se falhar, joga uma excecao com mensagem auto-explicativa que e
# capturada pela GUI e exibida em MessageBox.
Test-FeriasTemplate -Path $XlsxPath
Write-Host "Lendo planilha: $XlsxPath"
$rowsRaw = Read-XlsxRows -Path $XlsxPath

# Normaliza (nomes de coluna, datas, tipos)
$rows = foreach ($r in $rowsRaw) {
    if (-not $r.Colaborador -and -not $r.Mes) { continue }
    [pscustomobject]@{
        Mes         = "$($r.Mes)".Trim()
        Colaborador = "$($r.Colaborador)".Trim()
        Squad       = "$($r.Squad)".Trim()
        Inicio      = Format-DateBr $r.Inicio
        Fim         = Format-DateBr $r.Fim
        Dias        = [int]("$($r.Dias)".Trim())
        Status      = "$($r.Status)".Trim()
    }
}

if (-not $rows -or $rows.Count -eq 0) {
    throw "Nenhuma linha encontrada na fonte de dados."
}

Write-Host ("Linhas carregadas: {0}" -f $rows.Count)

# Filtro por ano selecionado: independente do que a planilha contem, mantem
# apenas as ferias que COMECAM no ano $Ano. Garante que selecionar 2027 nao
# misture com dados de 2026 caso o arquivo tenha varios anos.
$rowsAntesFiltro = $rows.Count
$rows = @($rows | Where-Object {
    $iso = ConvertTo-IsoDate $_.Inicio
    $iso -and $iso.StartsWith("$Ano-")
})
Write-Host ("Filtro por ano {0}: {1} -> {2} linhas" -f $Ano, $rowsAntesFiltro, $rows.Count)

if ($rows.Count -eq 0) {
    throw "Nenhuma linha de ferias para o ano $Ano. A planilha tem $rowsAntesFiltro linhas, mas nenhuma comeca em $Ano."
}

# 2. Dashboard por mes
$dashboardLines = @()
$grouped = $rows | Group-Object Mes | Sort-Object { Get-MonthIndex $_.Name }
foreach ($g in $grouped) {
    $count = $g.Count
    $squads = ($g.Group | Select-Object -ExpandProperty Squad -Unique | Sort-Object) -join ', '
    $hasSolicitada = ($g.Group | Where-Object { $_.Status -eq 'Solicitada' }).Count -gt 0
    $status = if ($count -ge 5) { '🟠 Atenção' }
              elseif ($hasSolicitada) { '🟡 Em Progresso' }
              else { '🟢 Alinhado' }
    $dashboardLines += "| **$($g.Name)** | $count | $squads | $status |"
}
$dashboardTable = @"
| Mês | Qtd. Pessoas | Squads Afetadas | Status Geral |
| :--- | :---: | :--- | :--- |
$($dashboardLines -join "`n")
"@

# 2.5 Controle de Vencimento de Ferias (CLT 12/24 meses)
# Le a aba Integrantes -- cadastro mestre com Data de admissao e Data fim
# das ultimas ferias por pessoa. A funcao Build-VencimentosSection classifica
# cada um em CRITICO (2o vencimento proximo/passado), ATENCAO (1o vencimento
# proximo/atingido) ou Dados incompletos (sem data de admissao).
$integrantesRaw = Read-IntegrantesRows -Path $XlsxPath
$vencimentosSection = Build-VencimentosSection -Integrantes $integrantesRaw -Hoje (Get-Date)

# 3. Cronograma detalhado
$cronogramaLines = foreach ($r in $rows) {
    $icon = $statusIcon[$r.Status]; if (-not $icon) { $icon = '⚪' }
    "| **$($r.Mes)** | $($r.Colaborador) | $($r.Squad) | $($r.Inicio) | $($r.Fim) | $($r.Dias) | $icon $($r.Status) |"
}
$cronogramaTable = @"
| Mês | Colaborador | Squad | Início | Fim | Dias | Status |
| :--- | :--- | :--- | :---: | :---: | :---: | :--- |
$($cronogramaLines -join "`n")
"@

# 4. Gantt (Mermaid)
$ganttBuilder = [System.Collections.Generic.List[string]]::new()
$ganttBuilder.Add('gantt')
$ganttBuilder.Add("    title Planejamento de Ferias $Ano")
$ganttBuilder.Add('    dateFormat YYYY-MM-DD')
$ganttBuilder.Add('    axisFormat %m/%y')
$idx = 1
foreach ($g in $grouped) {
    $ganttBuilder.Add("    section $($g.Name)")
    foreach ($r in $g.Group) {
        $iso = ConvertTo-IsoDate $r.Inicio
        $firstName = ($r.Colaborador -split '\s+')[0]
        $squadShort = (($r.Squad -split '[\s/\-]+')[0]).Trim()
        $label = "$firstName - $squadShort"
        $ganttBuilder.Add("    $label`t:t$idx, $iso, $($r.Dias)d")
        $idx++
    }
}
$ganttBlock = "```````mermaid`n$($ganttBuilder -join "`n")`n``````"

# 5. Renderiza template
if (-not (Test-Path $TemplatePath)) { throw "Template nao encontrado: $TemplatePath" }
$template = Get-Content -Path $TemplatePath -Raw -Encoding UTF8
$now = Get-Date
$autorLine = "*Criado e compilado por **$Autor** em $($now.ToString('dd/MM/yyyy')) às $($now.ToString('HH:mm')).*"
$md = $template.Replace('<!-- DASHBOARD -->', $dashboardTable).
                Replace('<!-- VENCIMENTOS -->', $vencimentosSection).
                Replace('<!-- CRONOGRAMA -->', $cronogramaTable).
                Replace('<!-- GANTT -->', $ganttBlock).
                Replace('<!-- AUTOR -->', $autorLine).
                Replace('<!-- ANO -->', "$Ano")

if (-not (Test-Path $OutputDir)) { New-Item -ItemType Directory -Path $OutputDir | Out-Null }

# Nomes fixos: cada execucao SOBRESCREVE a anterior. Sem timestamp, sem
# subpastas. Assim a pasta de saida fica enxuta com sempre o relatorio
# "atual" e a planilha-fonte que o gerou.
$outMd   = Join-Path $OutputDir 'Ferias.md'
$outHtml = Join-Path $OutputDir 'Ferias.html'
$outXlsx = Join-Path $OutputDir 'Ferias.xlsx'
Write-Host "Pasta de saida: $OutputDir"

# Copia da planilha-fonte como snapshot. A regra do template garante
# que XlsxPath aponta pra Ferias-template.xlsx valido, entao nao ha
# fallback CSV nem extensao variavel - sempre Ferias.xlsx.
Copy-Item -Path $XlsxPath -Destination $outXlsx -Force
Write-Host "Planilha-fonte copiada: $outXlsx"

$utf8NoBom = New-Object System.Text.UTF8Encoding($false)
[System.IO.File]::WriteAllText($outMd, $md, $utf8NoBom)
Write-Host "Markdown gerado: $outMd"

# 6. Pandoc -> HTML
$pandocArgs = @(
    $outMd,
    '-o', $outHtml,
    '--standalone',
    '--embed-resources',
    '--css', $CssPath,
    '--include-in-header', $HeaderPath,
    '--lua-filter', $LuaFilter,
    '--metadata', "title=Planejamento de Férias $Ano",
    '--metadata', 'lang=pt-BR',
    '--metadata', "author=$Autor"
)

# Pandoc emite warnings em stderr (ex: 'Could not fetch' por cert MITM corporativo).
# No PS 5.1 com ErrorActionPreference=Stop, stderr de nativo pode virar ErrorRecord.
# Rodamos com ErrorActionPreference=Continue e validamos o resultado pelo arquivo + exit code.
$prevEAP = $ErrorActionPreference
$ErrorActionPreference = 'Continue'
try {
    & pandoc @pandocArgs 2>&1 | ForEach-Object {
        if ($_ -is [System.Management.Automation.ErrorRecord]) {
            Write-Host "[pandoc] $($_.Exception.Message)" -ForegroundColor DarkGray
        } else {
            Write-Host "[pandoc] $_" -ForegroundColor DarkGray
        }
    }
    $pandocExit = $LASTEXITCODE
} finally {
    $ErrorActionPreference = $prevEAP
}

if ($pandocExit -ne 0 -or -not (Test-Path $outHtml)) {
    throw "Pandoc falhou (exit code $pandocExit). HTML nao gerado."
}
Write-Host "HTML gerado: $outHtml" -ForegroundColor Green

# 7. PDF (opcional, para SharePoint)
$outPdf = $null
if ($Pdf) {
    $outPdf = Join-Path $OutputDir 'Ferias.pdf'
    Convert-HtmlToPdf -HtmlPath $outHtml -PdfPath $outPdf
    Write-Host "PDF gerado: $outPdf" -ForegroundColor Green
}

if ($OpenAfter) {
    if ($Pdf -and $outPdf) { Start-Process $outPdf }
    else { Start-Process $outHtml }
}
