#requires -Version 5.1
<#
.SYNOPSIS
    Compila o instalador MSI do Ferias Automacao a partir de
    installer\Product.wxs usando WiX 4.

.DESCRIPTION
    Pre-requisito: WiX 4 instalado como dotnet global tool:
        dotnet tool install --global wix --version 4.0.6
        wix extension add WixToolset.UI.wixext/4.0.6 --global

    Saida: installer\output\FeriasAutomacao-1.0.0.msi

.EXAMPLE
    .\installer\build.ps1
#>

$ErrorActionPreference = 'Stop'
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8

$repoRoot     = Split-Path -Parent $PSScriptRoot
$installerDir = $PSScriptRoot
$outDir       = Join-Path $installerDir 'output'
$wxs          = Join-Path $installerDir 'Product.wxs'
$msi          = Join-Path $outDir 'FeriasAutomacao-1.0.0.msi'

# Garante a saida limpa
if (Test-Path $outDir) { Remove-Item $outDir -Recurse -Force }
New-Item -ItemType Directory -Path $outDir | Out-Null

# Confere se o WiX esta no PATH
try {
    $wixVersion = (& wix --version) 2>$null
    Write-Host "WiX detectada: $wixVersion" -ForegroundColor DarkGray
} catch {
    throw "WiX nao encontrada no PATH. Instale com:`n  dotnet tool install --global wix --version 4.0.6`n  wix extension add WixToolset.UI.wixext/4.0.6 --global"
}

# Roda no diretorio installer/ pra que os caminhos relativos
# (..\gui.ps1, etc) batam com o que o Product.wxs espera.
Push-Location $installerDir
try {
    & wix build `
        Product.wxs `
        -ext WixToolset.UI.wixext `
        -arch x64 `
        -o $msi
    if ($LASTEXITCODE -ne 0) {
        throw "wix build falhou com exit code $LASTEXITCODE"
    }
} finally {
    Pop-Location
}

# Resumo
$tamanhoMb = [math]::Round((Get-Item $msi).Length / 1MB, 2)
Write-Host ""
Write-Host "============================================================" -ForegroundColor Green
Write-Host (" MSI gerado: {0} MB" -f $tamanhoMb)                          -ForegroundColor Green
Write-Host "============================================================" -ForegroundColor Green
Write-Host ("  {0}" -f $msi)
