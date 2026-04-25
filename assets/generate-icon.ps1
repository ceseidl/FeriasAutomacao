#requires -Version 5.1
<#
.SYNOPSIS
    Gera assets/icon.ico (multi-resolucao) com palmeira branca em fundo azul.

.DESCRIPTION
    Cria um icone em varios tamanhos (16, 32, 48, 64, 128, 256) e empacota
    tudo num unico .ico no formato PNG-embedded (suportado por Windows Vista+).
    Tema: quadrado azul arredondado (cor do botao Gerar Relatorio: #4299E1)
    com silhueta de palmeira branca - referencia ao 🌴 usado no template.
#>

Add-Type -AssemblyName System.Drawing

$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$icoPath   = Join-Path $scriptDir 'icon.ico'
$pngPath   = Join-Path $scriptDir 'icon-256.png'

$sizes = @(16, 32, 48, 64, 128, 256)

function New-FeriasIcon {
    param([int]$Size)

    $bmp = New-Object System.Drawing.Bitmap($Size, $Size)
    $g   = [System.Drawing.Graphics]::FromImage($bmp)
    $g.SmoothingMode = [System.Drawing.Drawing2D.SmoothingMode]::AntiAlias
    $g.InterpolationMode = [System.Drawing.Drawing2D.InterpolationMode]::HighQualityBicubic
    $g.PixelOffsetMode = [System.Drawing.Drawing2D.PixelOffsetMode]::HighQuality

    # ---- Fundo: quadrado arredondado em gradiente azul ----
    $blueTop  = [System.Drawing.Color]::FromArgb(99, 179, 237)   # azul claro
    $blueBot  = [System.Drawing.Color]::FromArgb(43, 108, 176)   # azul escuro
    $radius   = [int]($Size * 0.18)
    $rect     = New-Object System.Drawing.Rectangle(0, 0, $Size, $Size)

    $path = New-Object System.Drawing.Drawing2D.GraphicsPath
    $path.AddArc($rect.X,                  $rect.Y,                 $radius * 2, $radius * 2, 180, 90)
    $path.AddArc($rect.Right - $radius*2,  $rect.Y,                 $radius * 2, $radius * 2, 270, 90)
    $path.AddArc($rect.Right - $radius*2,  $rect.Bottom - $radius*2,$radius * 2, $radius * 2, 0,   90)
    $path.AddArc($rect.X,                  $rect.Bottom - $radius*2,$radius * 2, $radius * 2, 90,  90)
    $path.CloseFigure()

    $brush = New-Object System.Drawing.Drawing2D.LinearGradientBrush(
        $rect, $blueTop, $blueBot, [System.Drawing.Drawing2D.LinearGradientMode]::Vertical)
    $g.FillPath($brush, $path)
    $brush.Dispose()

    # ---- Sol amarelo no canto superior direito ----
    $sunSize = [int]($Size * 0.28)
    $sunX    = $Size - $sunSize - [int]($Size * 0.10)
    $sunY    = [int]($Size * 0.10)
    $sunBrush = New-Object System.Drawing.SolidBrush(
        [System.Drawing.Color]::FromArgb(255, 236, 178, 64))
    $g.FillEllipse($sunBrush, $sunX, $sunY, $sunSize, $sunSize)
    $sunBrush.Dispose()

    # ---- Palmeira branca centralizada ----
    $white = New-Object System.Drawing.SolidBrush([System.Drawing.Color]::White)

    # Tronco: curva fina inclinada
    $trunkPen = New-Object System.Drawing.Pen($white, [single]([int]($Size * 0.06)))
    $trunkPen.StartCap = [System.Drawing.Drawing2D.LineCap]::Round
    $trunkPen.EndCap   = [System.Drawing.Drawing2D.LineCap]::Round

    $trunkBottom = New-Object System.Drawing.PointF([single]($Size * 0.52), [single]($Size * 0.85))
    $trunkMid    = New-Object System.Drawing.PointF([single]($Size * 0.45), [single]($Size * 0.60))
    $trunkTop    = New-Object System.Drawing.PointF([single]($Size * 0.50), [single]($Size * 0.38))
    $g.DrawCurve($trunkPen, @($trunkBottom, $trunkMid, $trunkTop))
    $trunkPen.Dispose()

    # Folhas: 5 elipses partindo do topo do tronco em direcoes diferentes
    $leafCenter = $trunkTop
    $leafLen    = [single]($Size * 0.32)
    $leafThick  = [single]($Size * 0.10)

    $leafAngles = @(160, 210, 260, 310, 350) # graus, partindo do leste = 0 horario
    foreach ($angle in $leafAngles) {
        $rad = [Math]::PI * $angle / 180.0
        $tipX = $leafCenter.X + [single]([Math]::Cos($rad) * $leafLen)
        $tipY = $leafCenter.Y + [single]([Math]::Sin($rad) * $leafLen)
        $midX = ($leafCenter.X + $tipX) / 2
        $midY = ($leafCenter.Y + $tipY) / 2

        $leafPath = New-Object System.Drawing.Drawing2D.GraphicsPath
        # Sagitta perpendicular para curvar a folha pra cima
        $perpX = -[Math]::Sin($rad)
        $perpY = [Math]::Cos($rad)
        $bend  = [single]($Size * 0.10)
        $ctrl1 = New-Object System.Drawing.PointF(
            [single]($leafCenter.X + ($tipX - $leafCenter.X) * 0.25 + $perpX * $bend),
            [single]($leafCenter.Y + ($tipY - $leafCenter.Y) * 0.25 + $perpY * $bend))
        $ctrl2 = New-Object System.Drawing.PointF(
            [single]($leafCenter.X + ($tipX - $leafCenter.X) * 0.75 + $perpX * $bend),
            [single]($leafCenter.Y + ($tipY - $leafCenter.Y) * 0.75 + $perpY * $bend))
        $tip   = New-Object System.Drawing.PointF($tipX, $tipY)

        # Borda superior da folha (curva)
        $leafPath.AddBezier($leafCenter, $ctrl1, $ctrl2, $tip)
        # Borda inferior (volta com leve deslocamento)
        $ctrl3 = New-Object System.Drawing.PointF(
            [single]($leafCenter.X + ($tipX - $leafCenter.X) * 0.75 + $perpX * ($bend * 0.4)),
            [single]($leafCenter.Y + ($tipY - $leafCenter.Y) * 0.75 + $perpY * ($bend * 0.4)))
        $ctrl4 = New-Object System.Drawing.PointF(
            [single]($leafCenter.X + ($tipX - $leafCenter.X) * 0.25 + $perpX * ($bend * 0.4)),
            [single]($leafCenter.Y + ($tipY - $leafCenter.Y) * 0.25 + $perpY * ($bend * 0.4)))
        $leafPath.AddBezier($tip, $ctrl3, $ctrl4, $leafCenter)
        $leafPath.CloseFigure()

        $g.FillPath($white, $leafPath)
        $leafPath.Dispose()
    }

    # ---- Cocos pequenos ----
    $cocoSize = [single]($Size * 0.05)
    $cocoBrush = New-Object System.Drawing.SolidBrush(
        [System.Drawing.Color]::FromArgb(255, 120, 80, 40))
    $g.FillEllipse($cocoBrush,
        [single]($leafCenter.X - $cocoSize),
        [single]($leafCenter.Y + $cocoSize * 0.3),
        $cocoSize, $cocoSize)
    $g.FillEllipse($cocoBrush,
        [single]($leafCenter.X + $cocoSize * 0.5),
        [single]($leafCenter.Y + $cocoSize),
        $cocoSize, $cocoSize)
    $cocoBrush.Dispose()

    $white.Dispose()
    $g.Dispose()
    return $bmp
}

function ConvertTo-PngBytes {
    param([System.Drawing.Bitmap]$Bitmap)
    $ms = New-Object System.IO.MemoryStream
    $Bitmap.Save($ms, [System.Drawing.Imaging.ImageFormat]::Png)
    return $ms.ToArray()
}

function Save-Ico {
    param(
        [hashtable]$PngBySize,   # @{16=byte[], 32=byte[], ...}
        [string]$Path
    )

    $sortedSizes = @($PngBySize.Keys | Sort-Object)
    $count       = $sortedSizes.Count

    $headerSize  = 6                       # ICONDIR
    $entrySize   = 16                      # ICONDIRENTRY
    $offset      = $headerSize + ($entrySize * $count)

    $ms = New-Object System.IO.MemoryStream

    # ICONDIR: reserved(2)=0, type(2)=1, count(2)
    $ms.Write([System.BitConverter]::GetBytes([uint16]0), 0, 2)
    $ms.Write([System.BitConverter]::GetBytes([uint16]1), 0, 2)
    $ms.Write([System.BitConverter]::GetBytes([uint16]$count), 0, 2)

    # ICONDIRENTRY x count
    foreach ($size in $sortedSizes) {
        $png = [byte[]]$PngBySize[$size]
        $w = if ($size -ge 256) { 0 } else { [byte]$size }
        $h = $w
        $ms.WriteByte([byte]$w)
        $ms.WriteByte([byte]$h)
        $ms.WriteByte([byte]0)        # color count (0 = no palette)
        $ms.WriteByte([byte]0)        # reserved
        $ms.Write([System.BitConverter]::GetBytes([uint16]1),  0, 2)   # color planes
        $ms.Write([System.BitConverter]::GetBytes([uint16]32), 0, 2)   # bpp
        $ms.Write([System.BitConverter]::GetBytes([uint32]$png.Length), 0, 4)
        $ms.Write([System.BitConverter]::GetBytes([uint32]$offset),     0, 4)
        $offset += $png.Length
    }

    # Image data (PNGs concatenados)
    foreach ($size in $sortedSizes) {
        $png = [byte[]]$PngBySize[$size]
        $ms.Write($png, 0, $png.Length)
    }

    [System.IO.File]::WriteAllBytes($Path, $ms.ToArray())
    $ms.Dispose()
}

# ================== Pipeline ==================
$pngBySize = @{}
foreach ($s in $sizes) {
    $bmp = New-FeriasIcon -Size $s
    $pngBySize[$s] = ConvertTo-PngBytes -Bitmap $bmp
    if ($s -eq 256) {
        $bmp.Save($pngPath, [System.Drawing.Imaging.ImageFormat]::Png)
    }
    $bmp.Dispose()
}

Save-Ico -PngBySize $pngBySize -Path $icoPath

Write-Host "Icone gerado em: $icoPath"
Write-Host "Preview 256px:   $pngPath"
