#requires -Version 5.1
<#
.SYNOPSIS
    Gera docs/splash.png reproduzindo a splash screen do gui.ps1 e capturando
    a regiao da tela ocupada pela janela.

.DESCRIPTION
    Reaproveita o layout da splash (mesmas dimensoes, gradiente, fontes e
    posicionamento). Nao executa nenhum handler do app principal. Mostra a
    janela, deixa o Windows pintar e usa Graphics.CopyFromScreen para salvar
    um PNG em docs/splash.png.
#>

$AUTOR_FIXO = 'Carlos Seidl'

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
[System.Windows.Forms.Application]::EnableVisualStyles()

Add-Type @"
using System;
using System.Runtime.InteropServices;
public class Win32Splash {
    [StructLayout(LayoutKind.Sequential)]
    public struct RECT { public int Left, Top, Right, Bottom; }
    [DllImport("user32.dll")]
    public static extern bool GetWindowRect(IntPtr hWnd, out RECT lpRect);
}
"@

$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$outPath   = Join-Path $scriptDir 'splash.png'
$iconPath  = Join-Path (Split-Path -Parent $scriptDir) 'assets\icon.ico'

$splash = New-Object System.Windows.Forms.Form
$splash.Text             = ''
$splash.Size             = New-Object System.Drawing.Size(528, 324)
$splash.StartPosition    = 'CenterScreen'
$splash.FormBorderStyle  = 'None'
$splash.ShowInTaskbar    = $false
$splash.TopMost          = $true
$splash.BackColor        = [System.Drawing.Color]::FromArgb(15, 15, 15)

$iconBmp = $null
if (Test-Path $iconPath) {
    try {
        $bytes = [System.IO.File]::ReadAllBytes($iconPath)
        if ($bytes.Length -ge 6) {
            $count = [BitConverter]::ToUInt16($bytes, 4)
            $bestSize = 0; $bestOffset = 0; $bestLength = 0
            for ($i = 0; $i -lt $count; $i++) {
                $entryOff = 6 + ($i * 16)
                $w = [int]$bytes[$entryOff]; if ($w -eq 0) { $w = 256 }
                $dataSize = [BitConverter]::ToUInt32($bytes, $entryOff + 8)
                $dataOff  = [BitConverter]::ToUInt32($bytes, $entryOff + 12)
                if ($w -gt $bestSize) {
                    $bestSize = $w; $bestOffset = $dataOff; $bestLength = $dataSize
                }
            }
            if ($bestLength -gt 0 -and $bytes[$bestOffset] -eq 0x89 -and $bytes[$bestOffset + 1] -eq 0x50) {
                $ms = New-Object System.IO.MemoryStream($bytes, $bestOffset, $bestLength)
                $iconBmp = [System.Drawing.Image]::FromStream($ms)
            } else {
                $rawIcon = New-Object System.Drawing.Icon($iconPath, 48, 48)
                $iconBmp = $rawIcon.ToBitmap()
                $rawIcon.Dispose()
            }
        }
    } catch { $iconBmp = $null }
}
$splash.Tag = $iconBmp

$splash.Add_Paint({
    param($sender, $e)
    $g = $e.Graphics
    $g.SmoothingMode     = [System.Drawing.Drawing2D.SmoothingMode]::AntiAlias
    $g.InterpolationMode = [System.Drawing.Drawing2D.InterpolationMode]::HighQualityBicubic

    $w = $sender.ClientSize.Width
    $h = $sender.ClientSize.Height

    $rect = New-Object System.Drawing.Rectangle(0, 0, $w, $h)
    $brush = New-Object System.Drawing.Drawing2D.LinearGradientBrush(
        $rect,
        [System.Drawing.Color]::FromArgb(31, 31, 31),
        [System.Drawing.Color]::FromArgb(5, 5, 5),
        [System.Drawing.Drawing2D.LinearGradientMode]::Vertical)
    $g.FillRectangle($brush, $rect)
    $brush.Dispose()

    $pen = New-Object System.Drawing.Pen([System.Drawing.Color]::FromArgb(80, 255, 255, 255), 1)
    $g.DrawRectangle($pen, 0, 0, $w - 1, $h - 1)
    $pen.Dispose()

    $bmp = $sender.Tag
    if ($bmp) {
        $iconSize = 115
        $iconX = [int](($w - $iconSize) / 2)
        $g.DrawImage($bmp, $iconX, 26, $iconSize, $iconSize)
    }
})

$lblName = New-Object System.Windows.Forms.Label
$lblName.Text       = 'Planejamento de Ferias'
$lblName.Font       = New-Object System.Drawing.Font('Segoe UI', 22, [System.Drawing.FontStyle]::Bold)
$lblName.ForeColor  = [System.Drawing.Color]::White
$lblName.BackColor  = [System.Drawing.Color]::Transparent
$lblName.AutoSize   = $false
$lblName.TextAlign  = 'MiddleCenter'
$lblName.Location   = New-Object System.Drawing.Point(0, 156)
$lblName.Size       = New-Object System.Drawing.Size(528, 46)
$splash.Controls.Add($lblName)

$lblAuthor = New-Object System.Windows.Forms.Label
$lblAuthor.Text       = "Desenvolvido por $AUTOR_FIXO"
$lblAuthor.Font       = New-Object System.Drawing.Font('Segoe UI', 12, [System.Drawing.FontStyle]::Italic)
$lblAuthor.ForeColor  = [System.Drawing.Color]::FromArgb(200, 200, 200)
$lblAuthor.BackColor  = [System.Drawing.Color]::Transparent
$lblAuthor.AutoSize   = $false
$lblAuthor.TextAlign  = 'MiddleCenter'
$lblAuthor.Location   = New-Object System.Drawing.Point(0, 210)
$lblAuthor.Size       = New-Object System.Drawing.Size(528, 26)
$splash.Controls.Add($lblAuthor)

$lblVersion = New-Object System.Windows.Forms.Label
$lblVersion.Text       = 'v1.0.0'
$lblVersion.Font       = New-Object System.Drawing.Font('Segoe UI', 10)
$lblVersion.ForeColor  = [System.Drawing.Color]::FromArgb(140, 140, 140)
$lblVersion.BackColor  = [System.Drawing.Color]::Transparent
$lblVersion.AutoSize   = $false
$lblVersion.TextAlign  = 'MiddleCenter'
$lblVersion.Location   = New-Object System.Drawing.Point(0, 276)
$lblVersion.Size       = New-Object System.Drawing.Size(528, 22)
$splash.Controls.Add($lblVersion)

$splash.Show()
$splash.TopMost = $true
$splash.Activate()
$splash.BringToFront()
[System.Windows.Forms.Application]::DoEvents()

1..40 | ForEach-Object {
    [System.Windows.Forms.Application]::DoEvents()
    Start-Sleep -Milliseconds 50
}

$rect = New-Object Win32Splash+RECT
[void][Win32Splash]::GetWindowRect($splash.Handle, [ref]$rect)
$w = $rect.Right - $rect.Left
$h = $rect.Bottom - $rect.Top

$bmp = New-Object System.Drawing.Bitmap($w, $h)
$gfx = [System.Drawing.Graphics]::FromImage($bmp)
$gfx.CopyFromScreen(
    (New-Object System.Drawing.Point($rect.Left, $rect.Top)),
    [System.Drawing.Point]::Empty,
    (New-Object System.Drawing.Size($w, $h)))
$bmp.Save($outPath, [System.Drawing.Imaging.ImageFormat]::Png)

$gfx.Dispose()
$bmp.Dispose()
$splash.Close()
$splash.Dispose()
if ($iconBmp) { $iconBmp.Dispose() }

Write-Host "Splash salvo em: $outPath"
