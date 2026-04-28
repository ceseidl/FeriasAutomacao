#requires -Version 5.1
<#
.SYNOPSIS
    Compila o instalador MSI do Ferias Automacao a partir de
    installer\Product.wxs usando WiX 4.

.DESCRIPTION
    Pre-requisito: WiX 4 instalado como dotnet global tool:
        dotnet tool install --global wix --version 4.0.6
        wix extension add --global WixToolset.UI.wixext/4.0.6

    Saida: installer\output\FeriasAutomacao-1.0.0.1.msi

.EXAMPLE
    .\installer\build.ps1
#>

$ErrorActionPreference = 'Stop'
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8

$repoRoot     = Split-Path -Parent $PSScriptRoot
$installerDir = $PSScriptRoot
$outDir       = Join-Path $installerDir 'output'
$wxs          = Join-Path $installerDir 'Product.wxs'
$msi          = Join-Path $outDir 'FeriasAutomacao-1.0.0.1.msi'

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

# Copia o MANUAL.docx pra pasta de saida ao lado do MSI. Assim quem for
# distribuir (SharePoint, e-mail, GitHub Release) tem o pacote completo
# numa pasta so: instalador + manual do usuario.
$manualSrc = Join-Path $repoRoot 'docs\MANUAL.docx'
$manualDst = Join-Path $outDir   'MANUAL.docx'
if (Test-Path $manualSrc) {
    Copy-Item -Path $manualSrc -Destination $manualDst -Force
    Write-Host "Manual copiado: $manualDst" -ForegroundColor DarkGray
} else {
    Write-Host "Aviso: $manualSrc nao encontrado, pulando copia do manual." -ForegroundColor Yellow
}

# Resumo
$tamanhoMb = [math]::Round((Get-Item $msi).Length / 1MB, 2)
Write-Host ""
Write-Host "============================================================" -ForegroundColor Green
Write-Host (" MSI gerado: {0} MB" -f $tamanhoMb)                          -ForegroundColor Green
Write-Host "============================================================" -ForegroundColor Green
Write-Host ("  {0}" -f $msi)
if (Test-Path $manualDst) {
    Write-Host ("  {0}" -f $manualDst)
}
