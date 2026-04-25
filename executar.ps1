#requires -Version 5.1
<#
.SYNOPSIS
    Gera Ferias-{timestamp}.md e Ferias-{timestamp}.html a partir de ferias-2026.xlsx.

.DESCRIPTION
    Le a planilha (aba "Ferias"), preenche o template.md com dashboard + cronograma + gantt,
    e roda pandoc pra gerar o HTML estilizado (CSS + Mermaid embutido).
    Cada execucao gera arquivos novos com timestamp (yyyyMMdd-HHmmss), preservando historico.

.PARAMETER XlsxPath
    Caminho para a planilha. Default: .\ferias-2026.xlsx

.PARAMETER CsvPath
    Alternativa ao xlsx. Se informado, sobrescreve XlsxPath. Separador esperado: ';'.

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
    [string]$XlsxPath = (Join-Path $PSScriptRoot 'ferias-2026.xlsx'),
    [string]$CsvPath,
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
    if (-not (Test-Path $Path)) { throw "Planilha nao encontrada: $Path" }
    if (-not (Get-Module -ListAvailable -Name ImportExcel)) {
        Write-Host "Modulo ImportExcel nao encontrado. Instalando no escopo do usuario..." -ForegroundColor Yellow
        Install-Module ImportExcel -Scope CurrentUser -Force -ErrorAction Stop
    }
    Import-Module ImportExcel -ErrorAction Stop | Out-Null
    $data = Import-Excel -Path $Path -WorksheetName 'Ferias' -ErrorAction Stop
    return , $data
}

function Read-CsvRows([string]$Path) {
    if (-not (Test-Path $Path)) { throw "CSV nao encontrado: $Path" }
    return , (Import-Csv -Path $Path -Delimiter ';' -Encoding UTF8)
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
if ($CsvPath) {
    Write-Host "Lendo CSV: $CsvPath"
    $rowsRaw = Read-CsvRows -Path $CsvPath
} else {
    Write-Host "Lendo planilha: $XlsxPath"
    $rowsRaw = Read-XlsxRows -Path $XlsxPath
}

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
                Replace('<!-- CRONOGRAMA -->', $cronogramaTable).
                Replace('<!-- GANTT -->', $ganttBlock).
                Replace('<!-- AUTOR -->', $autorLine).
                Replace('<!-- ANO -->', "$Ano")

if (-not (Test-Path $OutputDir)) { New-Item -ItemType Directory -Path $OutputDir | Out-Null }

$timestamp = Get-Date -Format 'yyyyMMdd-HHmmss'
$outMd   = Join-Path $OutputDir "Ferias-$timestamp.md"
$outHtml = Join-Path $OutputDir "Ferias-$timestamp.html"
Write-Host "Timestamp desta execucao: $timestamp"

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
    $outPdf = Join-Path $OutputDir "Ferias-$timestamp.pdf"
    Convert-HtmlToPdf -HtmlPath $outHtml -PdfPath $outPdf
    Write-Host "PDF gerado: $outPdf" -ForegroundColor Green
}

if ($OpenAfter) {
    if ($Pdf -and $outPdf) { Start-Process $outPdf }
    else { Start-Process $outHtml }
}
