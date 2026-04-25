#requires -Version 5.1
<#
.SYNOPSIS
    Gera assets/icon.ico (multi-resolucao) com o logo AI/R em fundo squircle escuro.

.DESCRIPTION
    Pipeline:
      1. Embute o SVG do logo AI/R (capturado de https://aircompany.ai/).
      2. Monta um HTML com squircle preto + logo branco centralizado.
      3. Usa Microsoft Edge (ou Chrome) em modo headless para rasterizar a 1024x1024.
      4. Redimensiona para 16, 32, 48, 64, 128 e 256.
      5. Empacota tudo num unico .ico no formato PNG-embedded.
#>

Add-Type -AssemblyName System.Drawing

$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$icoPath   = Join-Path $scriptDir 'icon.ico'
$pngPath   = Join-Path $scriptDir 'icon-256.png'

$sizes = @(16, 32, 48, 64, 128, 256)

# ============================================================
# Logo AI/R (https://aircompany.ai/content/dam/sites/logos/logo-header.svg)
# ============================================================
$airSvg = @'
<svg width="99" height="32" viewBox="0 0 99 32" fill="none" xmlns="http://www.w3.org/2000/svg">
<g>
<path d="M70.104 4.87286C68.1387 3.28079 65.1567 3.63458 63.5844 5.60765C62.8932 6.47852 62.5814 7.48547 62.473 8.56045C62.4052 9.28164 62.1748 9.94841 61.8495 10.588C60.8194 12.6699 59.3284 14.3844 57.5392 15.854C55.4248 17.5685 53.0935 18.7524 50.3149 18.8748C49.0814 18.9293 48.0378 19.5144 47.2245 20.4533C44.839 23.1884 46.3028 27.5427 49.8405 28.2503C52.5242 28.7946 55.0724 26.9032 55.4383 24.1681C55.5332 23.4061 55.723 22.6441 56.0212 21.9365C57.3766 18.7115 59.6537 16.303 62.6628 14.6157C64.1266 13.7993 65.6853 13.3366 67.3796 13.3502C68.7757 13.3502 69.9414 12.7379 70.8224 11.6629C72.5438 9.55379 72.2185 6.58739 70.104 4.87286Z" fill="white"/>
<path d="M37.4248 4.36939H43.6733V0.300781H27.1236V4.36939H33.3721V27.8557H27.1236V31.9243H43.6869V27.8557H37.4248V4.36939Z" fill="white"/>
<path d="M6.11457 0.300781L0.299805 7.10447V31.9107H4.35252V19.1606H17.0799V31.9107H21.1326V7.05004L14.7215 0.300781H6.11457ZM17.0935 15.092H4.36607V8.61488L7.98505 4.36939H12.9866L17.0799 8.68292V15.092H17.0935Z" fill="white"/>
<path d="M98.7576 7.05004L92.3465 0.300781H77.9248V31.9107H81.9775V19.1606H91.6146L94.7049 23.8551V31.9243H98.7576V22.6305L95.342 17.4597L98.7576 14.2075V7.05004ZM94.7049 12.4522L91.9399 15.092H81.9775V4.36939H90.6115L94.7049 8.68292V12.4522Z" fill="white"/>
</g>
</svg>
'@

# ============================================================
# Localiza Edge ou Chrome
# ============================================================
function Find-EdgeOrChrome {
    $candidates = @(
        "$env:ProgramFiles\Microsoft\Edge\Application\msedge.exe",
        "${env:ProgramFiles(x86)}\Microsoft\Edge\Application\msedge.exe",
        "$env:LOCALAPPDATA\Microsoft\Edge\Application\msedge.exe",
        "$env:ProgramFiles\Google\Chrome\Application\chrome.exe",
        "${env:ProgramFiles(x86)}\Google\Chrome\Application\chrome.exe",
        "$env:LOCALAPPDATA\Google\Chrome\Application\chrome.exe"
    )
    foreach ($p in $candidates) { if (Test-Path $p) { return $p } }
    return $null
}

# ============================================================
# Bitmap helpers
# ============================================================
function Resize-Bitmap {
    param([System.Drawing.Bitmap]$Source, [int]$Size)
    $out = New-Object System.Drawing.Bitmap($Size, $Size)
    $g   = [System.Drawing.Graphics]::FromImage($out)
    $g.SmoothingMode      = [System.Drawing.Drawing2D.SmoothingMode]::AntiAlias
    $g.InterpolationMode  = [System.Drawing.Drawing2D.InterpolationMode]::HighQualityBicubic
    $g.PixelOffsetMode    = [System.Drawing.Drawing2D.PixelOffsetMode]::HighQuality
    $g.CompositingQuality = [System.Drawing.Drawing2D.CompositingQuality]::HighQuality
    $g.DrawImage($Source, 0, 0, $Size, $Size)
    $g.Dispose()
    return $out
}

function ConvertTo-PngBytes {
    param([System.Drawing.Bitmap]$Bitmap)
    $ms = New-Object System.IO.MemoryStream
    $Bitmap.Save($ms, [System.Drawing.Imaging.ImageFormat]::Png)
    return $ms.ToArray()
}

function Save-Ico {
    param(
        [hashtable]$PngBySize,
        [string]$Path
    )
    $sortedSizes = @($PngBySize.Keys | Sort-Object)
    $count       = $sortedSizes.Count
    $headerSize  = 6
    $entrySize   = 16
    $offset      = $headerSize + ($entrySize * $count)

    $ms = New-Object System.IO.MemoryStream
    $ms.Write([System.BitConverter]::GetBytes([uint16]0), 0, 2)
    $ms.Write([System.BitConverter]::GetBytes([uint16]1), 0, 2)
    $ms.Write([System.BitConverter]::GetBytes([uint16]$count), 0, 2)

    foreach ($size in $sortedSizes) {
        $png = [byte[]]$PngBySize[$size]
        $w = if ($size -ge 256) { 0 } else { [byte]$size }
        $h = $w
        $ms.WriteByte([byte]$w)
        $ms.WriteByte([byte]$h)
        $ms.WriteByte([byte]0)
        $ms.WriteByte([byte]0)
        $ms.Write([System.BitConverter]::GetBytes([uint16]1),  0, 2)
        $ms.Write([System.BitConverter]::GetBytes([uint16]32), 0, 2)
        $ms.Write([System.BitConverter]::GetBytes([uint32]$png.Length), 0, 4)
        $ms.Write([System.BitConverter]::GetBytes([uint32]$offset),     0, 4)
        $offset += $png.Length
    }
    foreach ($size in $sortedSizes) {
        $png = [byte[]]$PngBySize[$size]
        $ms.Write($png, 0, $png.Length)
    }
    [System.IO.File]::WriteAllBytes($Path, $ms.ToArray())
    $ms.Dispose()
}

# ============================================================
# Pipeline principal
# ============================================================
$browser = Find-EdgeOrChrome
if (-not $browser) { throw 'Microsoft Edge ou Google Chrome nao encontrado.' }

$tempBase    = Join-Path $env:TEMP "air-icon-$([guid]::NewGuid().ToString('N').Substring(0,8))"
$tempHtml    = Join-Path $tempBase 'icon.html'
$tempPng     = Join-Path $tempBase 'icon.png'
$tempProfile = Join-Path $tempBase 'profile'
New-Item -ItemType Directory -Path $tempProfile -Force | Out-Null

# HTML: squircle preto com gradiente sutil + logo AI/R branco centralizado
$html = @"
<!DOCTYPE html>
<html><head><style>
  html, body { margin: 0; padding: 0; background: transparent; }
  body { width: 1024px; height: 1024px; }
  .icon {
    width: 1024px; height: 1024px;
    background: linear-gradient(135deg, #1f1f1f 0%, #050505 100%);
    border-radius: 184px;
    display: flex; align-items: center; justify-content: center;
    box-sizing: border-box;
    box-shadow: inset 0 0 0 6px rgba(255,255,255,0.04);
  }
  .icon svg { width: 720px; height: auto; }
</style></head><body>
<div class="icon">
$airSvg
</div>
</body></html>
"@
[System.IO.File]::WriteAllText($tempHtml, $html, [System.Text.Encoding]::UTF8)

$tempHtmlUri = (New-Object System.Uri($tempHtml)).AbsoluteUri
$browserArgs = @(
    '--headless=new',
    '--disable-gpu',
    '--hide-scrollbars',
    '--default-background-color=00000000',
    '--window-size=1024,1024',
    "--user-data-dir=$tempProfile",
    "--screenshot=$tempPng",
    '--virtual-time-budget=2000',
    $tempHtmlUri
)

Write-Host "Renderizando logo AI/R em 1024x1024 via $browser..."
$proc = Start-Process -FilePath $browser -ArgumentList $browserArgs -Wait -PassThru -NoNewWindow
if ($proc.ExitCode -ne 0) { throw "Browser headless retornou exit code $($proc.ExitCode)." }
if (-not (Test-Path $tempPng)) { throw "PNG nao foi criado em $tempPng." }

$baseBmp = [System.Drawing.Bitmap]::FromFile($tempPng)
Write-Host "Renderizado: $($baseBmp.Width)x$($baseBmp.Height)"

$pngBySize = @{}
foreach ($s in $sizes) {
    $bmp = Resize-Bitmap -Source $baseBmp -Size $s
    $pngBySize[$s] = ConvertTo-PngBytes -Bitmap $bmp
    if ($s -eq 256) {
        $bmp.Save($pngPath, [System.Drawing.Imaging.ImageFormat]::Png)
    }
    $bmp.Dispose()
}
$baseBmp.Dispose()

Save-Ico -PngBySize $pngBySize -Path $icoPath

Remove-Item -Path $tempBase -Recurse -Force -ErrorAction SilentlyContinue

Write-Host "Icone AI/R gerado em: $icoPath"
Write-Host "Preview 256px:        $pngPath"
