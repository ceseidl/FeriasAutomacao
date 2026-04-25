#requires -Version 5.1
<#
.SYNOPSIS
    Gera localmente o ZIP de distribuicao do FeriasAutomacao, mesma
    estrutura do que o workflow .github/workflows/build-zip.yml produz
    no GitHub Actions.

.DESCRIPTION
    Util pra distribuir pra outras maquinas sem depender da release do
    GitHub. O ZIP final tem so os arquivos que o usuario final precisa
    (atalhos, scripts, planilha, recursos, manual, instalador offline
    do Pandoc).

    Saida: dist\FeriasAutomacao-<data>-<sha>.zip
           dist\FeriasAutomacao-latest.zip (copia com nome estavel)

.EXAMPLE
    .\build-local-zip.ps1
#>

$ErrorActionPreference = 'Stop'
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8

$repoRoot = $PSScriptRoot
$distDir  = Join-Path $repoRoot 'dist'
$stageDir = Join-Path $distDir   'FeriasAutomacao'

# Limpa stage anterior
if (Test-Path $stageDir) { Remove-Item $stageDir -Recurse -Force }
if (-not (Test-Path $distDir)) { New-Item -ItemType Directory -Path $distDir | Out-Null }

# Cria estrutura
New-Item -ItemType Directory -Path (Join-Path $stageDir 'assets') | Out-Null
New-Item -ItemType Directory -Path (Join-Path $stageDir 'bin')    | Out-Null
New-Item -ItemType Directory -Path (Join-Path $stageDir 'docs')   | Out-Null

# Lista de copias: source -> destino (relativo ao stage)
$arquivos = @(
    @{ Src = 'Gerar Relatorio.lnk';    Dst = 'Gerar Relatorio.lnk' }
    @{ Src = 'Gerar Relatorio.bat';    Dst = 'Gerar Relatorio.bat' }
    @{ Src = 'gui.ps1';                Dst = 'gui.ps1' }
    @{ Src = 'executar.ps1';           Dst = 'executar.ps1' }
    @{ Src = 'template.md';            Dst = 'template.md' }
    @{ Src = 'ferias-2026.xlsx';       Dst = 'ferias-2026.xlsx' }
    @{ Src = 'LICENSE';                Dst = 'LICENSE' }
    @{ Src = 'assets\icon.ico';        Dst = 'assets\icon.ico' }
    @{ Src = 'assets\style.css';       Dst = 'assets\style.css' }
    @{ Src = 'assets\header.html';     Dst = 'assets\header.html' }
    @{ Src = 'assets\mermaid.lua';     Dst = 'assets\mermaid.lua' }
    @{ Src = 'bin\pandoc-installer.msi'; Dst = 'bin\pandoc-installer.msi' }
    @{ Src = 'docs\MANUAL.docx';       Dst = 'docs\MANUAL.docx' }
)

foreach ($a in $arquivos) {
    $src = Join-Path $repoRoot $a.Src
    $dst = Join-Path $stageDir $a.Dst
    if (-not (Test-Path $src)) {
        throw "Arquivo nao encontrado: $src"
    }
    Copy-Item -Path $src -Destination $dst -Force
}

# Nomes de saida
$shortSha = (git rev-parse --short HEAD).Trim()
$dateUtc  = (Get-Date).ToUniversalTime().ToString('yyyyMMdd-HHmmss')
$zipName  = "FeriasAutomacao-$dateUtc-$shortSha.zip"
$zipPath  = Join-Path $distDir $zipName
$zipLatest = Join-Path $distDir 'FeriasAutomacao-latest.zip'

# Remove ZIPs anteriores com mesmo nome
if (Test-Path $zipPath)   { Remove-Item $zipPath -Force }
if (Test-Path $zipLatest) { Remove-Item $zipLatest -Force }

# Compacta a pasta-stage (o ZIP contem a pasta FeriasAutomacao na raiz)
Compress-Archive -Path $stageDir -DestinationPath $zipPath -CompressionLevel Optimal

# Copia com nome estavel
Copy-Item -Path $zipPath -Destination $zipLatest -Force

# Limpa stage (so o ZIP fica)
Remove-Item $stageDir -Recurse -Force

# Resumo
$tamanhoMb = [math]::Round((Get-Item $zipPath).Length / 1MB, 2)
Write-Host ""
Write-Host "============================================================" -ForegroundColor Green
Write-Host (" ZIP gerado: {0} MB" -f $tamanhoMb)                          -ForegroundColor Green
Write-Host "============================================================" -ForegroundColor Green
Write-Host ("  {0}" -f $zipPath)
Write-Host ("  {0}" -f $zipLatest)
Write-Host ""
Write-Host "Conteudo:" -ForegroundColor Cyan
Add-Type -AssemblyName System.IO.Compression.FileSystem
$zipObj = [System.IO.Compression.ZipFile]::OpenRead($zipPath)
$zipObj.Entries | Sort-Object FullName | ForEach-Object {
    $kb = if ($_.Length -gt 0) { ' {0,8} bytes  ' -f $_.Length } else { '              ' }
    Write-Host ("{0}{1}" -f $kb, $_.FullName)
}
$zipObj.Dispose()
