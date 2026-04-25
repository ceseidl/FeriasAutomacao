#requires -Version 5.1
<#
.SYNOPSIS
    Gera docs/screenshot.png reproduzindo a janela do gui.ps1 (sem rodar o
    executar.ps1) e capturando a regiao da tela ocupada pelo Form.

.DESCRIPTION
    Reaproveita o layout do gui.ps1 (mesmas dimensoes, fontes, cores e
    posicionamento dos controles) so que sem handlers funcionais. Mostra a
    janela, deixa o Windows pintar e usa Graphics.CopyFromScreen para salvar
    um PNG em docs/screenshot.png.
#>

$AUTOR_FIXO = 'Carlos Seidl'

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
[System.Windows.Forms.Application]::EnableVisualStyles()

# Win32 helpers para capturar a janela exatamente (sem sangramento)
Add-Type @"
using System;
using System.Runtime.InteropServices;
public class Win32 {
    [StructLayout(LayoutKind.Sequential)]
    public struct RECT { public int Left, Top, Right, Bottom; }
    [DllImport("user32.dll")]
    public static extern bool GetWindowRect(IntPtr hWnd, out RECT lpRect);
    [DllImport("user32.dll")]
    public static extern bool PrintWindow(IntPtr hwnd, IntPtr hdcBlt, uint nFlags);
}
"@

$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$outPath   = Join-Path $scriptDir 'screenshot.png'
$xlsxDemo  = Join-Path (Split-Path -Parent $scriptDir) 'ferias-2026.xlsx'

# ================== Form (espelho do gui.ps1) ==================
$form = New-Object System.Windows.Forms.Form
$form.Text = 'Planejamento de Ferias - Gerador'
$form.Size = New-Object System.Drawing.Size(500, 320)
$form.StartPosition = 'CenterScreen'
$form.FormBorderStyle = 'FixedDialog'
$form.MaximizeBox = $false
$form.MinimizeBox = $true
$form.Font = New-Object System.Drawing.Font('Segoe UI', 9)
$form.BackColor = [System.Drawing.Color]::White
$form.TopMost = $true
$form.ShowInTaskbar = $false

$iconPath = Join-Path (Split-Path -Parent $scriptDir) 'assets\icon.ico'
if (Test-Path $iconPath) {
    $form.Icon = New-Object System.Drawing.Icon($iconPath)
}

$lblTitle = New-Object System.Windows.Forms.Label
$lblTitle.Text = 'Planejamento de Ferias'
$lblTitle.Location = New-Object System.Drawing.Point(15, 12)
$lblTitle.Size = New-Object System.Drawing.Size(460, 24)
$lblTitle.Font = New-Object System.Drawing.Font('Segoe UI', 13, [System.Drawing.FontStyle]::Bold)
$lblTitle.ForeColor = [System.Drawing.Color]::FromArgb(45, 55, 72)
$form.Controls.Add($lblTitle)

$lblAno = New-Object System.Windows.Forms.Label
$lblAno.Text = 'Ano:'
$lblAno.Location = New-Object System.Drawing.Point(15, 55)
$lblAno.Size = New-Object System.Drawing.Size(75, 22)
$form.Controls.Add($lblAno)

$nudAno = New-Object System.Windows.Forms.NumericUpDown
$nudAno.Location = New-Object System.Drawing.Point(90, 53)
$nudAno.Size = New-Object System.Drawing.Size(90, 22)
$nudAno.Minimum = 2020
$nudAno.Maximum = 2100
$nudAno.Value = (Get-Date).Year
$nudAno.TextAlign = 'Center'
$form.Controls.Add($nudAno)

$lblXlsx = New-Object System.Windows.Forms.Label
$lblXlsx.Text = 'Planilha:'
$lblXlsx.Location = New-Object System.Drawing.Point(15, 88)
$lblXlsx.Size = New-Object System.Drawing.Size(75, 22)
$form.Controls.Add($lblXlsx)

$txtXlsx = New-Object System.Windows.Forms.TextBox
$txtXlsx.Location = New-Object System.Drawing.Point(90, 86)
$txtXlsx.Size = New-Object System.Drawing.Size(295, 22)
$txtXlsx.Text = $xlsxDemo
$form.Controls.Add($txtXlsx)

$btnBrowse = New-Object System.Windows.Forms.Button
$btnBrowse.Text = 'Procurar...'
$btnBrowse.Location = New-Object System.Drawing.Point(390, 84)
$btnBrowse.Size = New-Object System.Drawing.Size(85, 26)
$form.Controls.Add($btnBrowse)

$chkOpen = New-Object System.Windows.Forms.CheckBox
$chkOpen.Text = 'Abrir HTML apos gerar'
$chkOpen.Location = New-Object System.Drawing.Point(90, 120)
$chkOpen.Size = New-Object System.Drawing.Size(385, 22)
$chkOpen.Checked = $true
$form.Controls.Add($chkOpen)

$lblStatus = New-Object System.Windows.Forms.Label
$lblStatus.Location = New-Object System.Drawing.Point(15, 160)
$lblStatus.Size = New-Object System.Drawing.Size(460, 22)
$lblStatus.Text = 'Pronto. Confira os campos e clique em "Gerar Relatorio".'
$lblStatus.ForeColor = [System.Drawing.Color]::Gray
$form.Controls.Add($lblStatus)

$pb = New-Object System.Windows.Forms.ProgressBar
$pb.Location = New-Object System.Drawing.Point(15, 188)
$pb.Size = New-Object System.Drawing.Size(460, 14)
$pb.Style = [System.Windows.Forms.ProgressBarStyle]::Marquee
$pb.MarqueeAnimationSpeed = 30
$pb.Visible = $false
$form.Controls.Add($pb)

$lblAutoria = New-Object System.Windows.Forms.Label
$lblAutoria.Text = "Criado por $AUTOR_FIXO"
$lblAutoria.Location = New-Object System.Drawing.Point(15, 250)
$lblAutoria.Size = New-Object System.Drawing.Size(220, 20)
$lblAutoria.Font = New-Object System.Drawing.Font('Segoe UI', 8, [System.Drawing.FontStyle]::Italic)
$lblAutoria.ForeColor = [System.Drawing.Color]::Gray
$form.Controls.Add($lblAutoria)

$btnClose = New-Object System.Windows.Forms.Button
$btnClose.Text = 'Fechar'
$btnClose.Location = New-Object System.Drawing.Point(247, 245)
$btnClose.Size = New-Object System.Drawing.Size(85, 32)
$btnClose.Font = New-Object System.Drawing.Font('Segoe UI', 9)
$btnClose.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
$btnClose.FlatAppearance.BorderColor = [System.Drawing.Color]::FromArgb(203, 213, 224)
$btnClose.FlatAppearance.BorderSize = 1
$btnClose.BackColor = [System.Drawing.Color]::White
$btnClose.ForeColor = [System.Drawing.Color]::FromArgb(45, 55, 72)
$form.Controls.Add($btnClose)

$btnGenerate = New-Object System.Windows.Forms.Button
$btnGenerate.Text = 'Gerar Relatorio'
$btnGenerate.Location = New-Object System.Drawing.Point(340, 245)
$btnGenerate.Size = New-Object System.Drawing.Size(135, 32)
$btnGenerate.Font = New-Object System.Drawing.Font('Segoe UI', 9, [System.Drawing.FontStyle]::Bold)
$btnGenerate.BackColor = [System.Drawing.Color]::FromArgb(66, 153, 225)
$btnGenerate.ForeColor = [System.Drawing.Color]::White
$btnGenerate.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
$btnGenerate.FlatAppearance.BorderSize = 0
$form.Controls.Add($btnGenerate)

# ================== Mostra a janela e captura ==================
$form.Show()
$form.TopMost = $true
$form.Activate()
$form.BringToFront()
[System.Windows.Forms.Application]::DoEvents()

# Pinta a janela varias vezes para garantir que esta totalmente desenhada
# e na frente das demais janelas
1..40 | ForEach-Object {
    [System.Windows.Forms.Application]::DoEvents()
    Start-Sleep -Milliseconds 50
}

# GetWindowRect retorna os bounds reais (incluindo qualquer ajuste do Win11)
$rect = New-Object Win32+RECT
[void][Win32]::GetWindowRect($form.Handle, [ref]$rect)
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
$form.Close()
$form.Dispose()

Write-Host "Screenshot salvo em: $outPath"
