#requires -Version 5.1
<#
.SYNOPSIS
    GUI WinForms para gerar o relatorio de Planejamento de Ferias.

.DESCRIPTION
    Janela com Ano (editavel) e Planilha + botao "Gerar Relatorio".
    O Autor e fixo (definido na constante $AUTOR_FIXO abaixo).
    Lancado via "Gerar Relatorio.lnk" ou "Gerar Relatorio.bat".
#>

# ============================================================
# Nome do autor a ser exibido no rodape do HTML, no metadado
# e no rodape da janela. Para trocar de pessoa, edite aqui.
# ============================================================
$AUTOR_FIXO = 'Carlos Seidl'

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
[System.Windows.Forms.Application]::EnableVisualStyles()

$scriptDir   = Split-Path -Parent $MyInvocation.MyCommand.Path
$executar    = Join-Path $scriptDir 'executar.ps1'
$xlsxDefault = Join-Path $scriptDir 'Ferias-template.xlsx'
$resultsDir  = Join-Path $scriptDir 'results'

if (-not (Test-Path $executar)) {
    [System.Windows.Forms.MessageBox]::Show(
        "executar.ps1 nao encontrado em:`n$scriptDir",
        'Erro de configuracao', 'OK', 'Error') | Out-Null
    exit 1
}

# ================== Splash Screen ==================
function Show-SplashScreen {
    param(
        [string]$AppName  = 'Planejamento de Ferias',
        [string]$Author   = 'Carlos Seidl',
        [string]$IconPath,
        [int]$DurationMs  = 1800
    )

    $splash = New-Object System.Windows.Forms.Form
    $splash.Text             = ''
    $splash.Size             = New-Object System.Drawing.Size(528, 324)
    $splash.StartPosition    = 'CenterScreen'
    $splash.FormBorderStyle  = 'None'
    $splash.ShowInTaskbar    = $false
    $splash.TopMost          = $true
    $splash.UseWaitCursor    = $true
    $splash.Cursor           = [System.Windows.Forms.Cursors]::WaitCursor
    $splash.BackColor        = [System.Drawing.Color]::FromArgb(15, 15, 15)

    # Pre-carrega o icone como bitmap em alta resolucao.
    # O ICO tem entradas PNG-encoded para 64+ que System.Drawing.Icon.ToBitmap()
    # nao consegue ler no .NET Framework. Solucao: parse manual do arquivo ICO
    # para extrair a maior entrada (PNG ou BMP) e carregar via Image.FromStream.
    # Guardamos no Tag do form para acesso confiavel dentro do Paint handler.
    $iconBmp = $null
    if ($IconPath -and (Test-Path $IconPath)) {
        try {
            $bytes = [System.IO.File]::ReadAllBytes($IconPath)
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
                # PNG signature: 89 50 4E 47
                if ($bestLength -gt 0 -and $bytes[$bestOffset] -eq 0x89 -and $bytes[$bestOffset + 1] -eq 0x50) {
                    $ms = New-Object System.IO.MemoryStream($bytes, $bestOffset, $bestLength)
                    $iconBmp = [System.Drawing.Image]::FromStream($ms)
                } else {
                    # Fallback: BMP/DIB embutido - usa Icon constructor com tamanho menor
                    $rawIcon = New-Object System.Drawing.Icon($IconPath, 48, 48)
                    $iconBmp = $rawIcon.ToBitmap()
                    $rawIcon.Dispose()
                }
            }
        } catch { $iconBmp = $null }
    }
    $splash.Tag = $iconBmp

    # Desenha tudo no Paint (gradiente + borda + icone) para evitar
    # problemas de PictureBox transparente sobre gradiente
    $splash.Add_Paint({
        param($sender, $e)
        $g = $e.Graphics
        $g.SmoothingMode     = [System.Drawing.Drawing2D.SmoothingMode]::AntiAlias
        $g.InterpolationMode = [System.Drawing.Drawing2D.InterpolationMode]::HighQualityBicubic

        $w = $sender.ClientSize.Width
        $h = $sender.ClientSize.Height

        # Gradiente preto/cinza-escuro (paleta AI/R)
        $rect = New-Object System.Drawing.Rectangle(0, 0, $w, $h)
        $brush = New-Object System.Drawing.Drawing2D.LinearGradientBrush(
            $rect,
            [System.Drawing.Color]::FromArgb(31, 31, 31),
            [System.Drawing.Color]::FromArgb(5, 5, 5),
            [System.Drawing.Drawing2D.LinearGradientMode]::Vertical)
        $g.FillRectangle($brush, $rect)
        $brush.Dispose()

        # Borda sutil
        $pen = New-Object System.Drawing.Pen([System.Drawing.Color]::FromArgb(80, 255, 255, 255), 1)
        $g.DrawRectangle($pen, 0, 0, $w - 1, $h - 1)
        $pen.Dispose()

        # Icone centralizado (recuperado via Tag do form)
        $bmp = $sender.Tag
        if ($bmp) {
            $iconSize = 115
            $iconX = [int](($w - $iconSize) / 2)
            $g.DrawImage($bmp, $iconX, 26, $iconSize, $iconSize)
        }
    })

    # Nome do app
    $lblName = New-Object System.Windows.Forms.Label
    $lblName.Text       = $AppName
    $lblName.Font       = New-Object System.Drawing.Font('Segoe UI', 22, [System.Drawing.FontStyle]::Bold)
    $lblName.ForeColor  = [System.Drawing.Color]::White
    $lblName.BackColor  = [System.Drawing.Color]::Transparent
    $lblName.AutoSize   = $false
    $lblName.TextAlign  = 'MiddleCenter'
    $lblName.Location   = New-Object System.Drawing.Point(0, 156)
    $lblName.Size       = New-Object System.Drawing.Size(528, 46)
    $splash.Controls.Add($lblName)

    # Feito por
    $lblAuthor = New-Object System.Windows.Forms.Label
    $lblAuthor.Text       = "Desenvolvido por $Author"
    $lblAuthor.Font       = New-Object System.Drawing.Font('Segoe UI', 12, [System.Drawing.FontStyle]::Italic)
    $lblAuthor.ForeColor  = [System.Drawing.Color]::FromArgb(200, 200, 200)
    $lblAuthor.BackColor  = [System.Drawing.Color]::Transparent
    $lblAuthor.AutoSize   = $false
    $lblAuthor.TextAlign  = 'MiddleCenter'
    $lblAuthor.Location   = New-Object System.Drawing.Point(0, 210)
    $lblAuthor.Size       = New-Object System.Drawing.Size(528, 26)
    $splash.Controls.Add($lblAuthor)

    # Versao discreta no rodape
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
    $splash.Refresh()
    [System.Windows.Forms.Cursor]::Current = [System.Windows.Forms.Cursors]::WaitCursor
    [System.Windows.Forms.Application]::DoEvents()

    # Mantem visivel pelo tempo configurado, com pump de eventos.
    # Forca o cursor de ampulheta a cada iteracao para que fique
    # sticky mesmo quando o mouse esta fora da splash.
    $endAt = (Get-Date).AddMilliseconds($DurationMs)
    while ((Get-Date) -lt $endAt) {
        [System.Windows.Forms.Cursor]::Current = [System.Windows.Forms.Cursors]::WaitCursor
        [System.Windows.Forms.Application]::DoEvents()
        Start-Sleep -Milliseconds 30
    }

    [System.Windows.Forms.Cursor]::Current = [System.Windows.Forms.Cursors]::Default
    $splash.Close()
    $splash.Dispose()
    if ($iconBmp) { $iconBmp.Dispose() }
}

# ================== Year Picker (popup estilo calendario) ==================
# Usa ToolStripDropDown -> fecha automaticamente em clique fora / Esc / perda de foco
function Show-YearPicker {
    param(
        [Parameter(Mandatory)] [System.Windows.Forms.Control] $AnchorControl,
        [Parameter(Mandatory)] [int] $CurrentYear,
        [int] $MinYear = 2020,
        [int] $MaxYear = 2100,
        [string] $TextFormat = '{0}   {1}'
    )

    $arrowChar = [string][char]0x25BE

    $popupW = 264; $popupH = 226

    # Painel root que vai dentro do dropdown
    $root = New-Object System.Windows.Forms.Panel
    $root.Size = New-Object System.Drawing.Size($popupW, $popupH)
    $root.BackColor = [System.Drawing.Color]::White
    $root.Add_Paint({
        param($s, $e)
        $pen = New-Object System.Drawing.Pen([System.Drawing.Color]::FromArgb(180, 180, 180), 1)
        $e.Graphics.DrawRectangle($pen, 0, 0, $s.Width - 1, $s.Height - 1)
        $pen.Dispose()
    })

    # Estado mutavel (acessivel nos closures aninhados via hashtable)
    $state = @{
        SelectedYear = $CurrentYear
        BaseDecade   = [int]([Math]::Floor($CurrentYear / 10) * 10)
        MinYear      = $MinYear
        MaxYear      = $MaxYear
        Anchor       = $AnchorControl
        TextFormat   = $TextFormat
        ArrowChar    = $arrowChar
        Dropdown     = $null  # preenchido depois que $dropdown for criado
    }

    # Header
    $btnPrev = New-Object System.Windows.Forms.Button
    $btnPrev.Text = [string][char]0x25C0
    $btnPrev.Location = New-Object System.Drawing.Point(4, 4)
    $btnPrev.Size = New-Object System.Drawing.Size(36, 30)
    $btnPrev.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
    $btnPrev.FlatAppearance.BorderSize = 0
    $btnPrev.BackColor = [System.Drawing.Color]::White
    $btnPrev.ForeColor = [System.Drawing.Color]::FromArgb(45, 55, 72)
    $btnPrev.Font = New-Object System.Drawing.Font('Segoe UI', 9)
    $btnPrev.TabStop = $false
    $root.Controls.Add($btnPrev)

    $lblDecade = New-Object System.Windows.Forms.Label
    $lblDecade.TextAlign = 'MiddleCenter'
    $lblDecade.Location = New-Object System.Drawing.Point(40, 4)
    $lblDecade.Size = New-Object System.Drawing.Size(184, 30)
    $lblDecade.Font = New-Object System.Drawing.Font('Segoe UI Semibold', 10)
    $lblDecade.ForeColor = [System.Drawing.Color]::FromArgb(45, 55, 72)
    $root.Controls.Add($lblDecade)

    $btnNext = New-Object System.Windows.Forms.Button
    $btnNext.Text = [string][char]0x25B6
    $btnNext.Location = New-Object System.Drawing.Point(224, 4)
    $btnNext.Size = New-Object System.Drawing.Size(36, 30)
    $btnNext.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
    $btnNext.FlatAppearance.BorderSize = 0
    $btnNext.BackColor = [System.Drawing.Color]::White
    $btnNext.ForeColor = [System.Drawing.Color]::FromArgb(45, 55, 72)
    $btnNext.Font = New-Object System.Drawing.Font('Segoe UI', 9)
    $btnNext.TabStop = $false
    $root.Controls.Add($btnNext)

    # Separador
    $sep = New-Object System.Windows.Forms.Label
    $sep.Location = New-Object System.Drawing.Point(8, 38)
    $sep.Size = New-Object System.Drawing.Size(248, 1)
    $sep.BackColor = [System.Drawing.Color]::FromArgb(220, 220, 220)
    $root.Controls.Add($sep)

    # Grade de anos
    $gridPanel = New-Object System.Windows.Forms.Panel
    $gridPanel.Location = New-Object System.Drawing.Point(4, 42)
    $gridPanel.Size = New-Object System.Drawing.Size(256, 180)
    $gridPanel.BackColor = [System.Drawing.Color]::White
    $root.Controls.Add($gridPanel)

    # Cria o dropdown ja, pois o handler do clique no ano precisa fechar ele
    $dropdown = New-Object System.Windows.Forms.ToolStripDropDown
    $dropdown.AutoClose          = $true
    $dropdown.Padding            = [System.Windows.Forms.Padding]::Empty
    $dropdown.Margin             = [System.Windows.Forms.Padding]::Empty
    $dropdown.DropShadowEnabled  = $true
    $dropdown.BackColor          = [System.Drawing.Color]::White
    $state.Dropdown = $dropdown

    # Handler de clique de ano definido no escopo da funcao (onde $state e LOCAL)
    # Closures aninhados em PowerShell nao capturam variaveis herdadas, por isso
    # criamos o scriptblock aqui e reusamos em todas as celulas via Add_Click().
    $onYearClick = {
        try {
            $picked = [int]$this.Tag
            $state.Anchor.Tag  = $picked
            $state.Anchor.Text = ($state.TextFormat -f $picked, $state.ArrowChar)
        } catch {
            [System.Windows.Forms.MessageBox]::Show("Erro ao selecionar ano: $($_.Exception.Message)", 'Erro', 'OK', 'Error') | Out-Null
        }
        $state.Dropdown.Close()
    }.GetNewClosure()

    $renderGrid = {
        $gridPanel.Controls.Clear()
        $startYear = $state.BaseDecade - 1
        $lblDecade.Text = "$($state.BaseDecade) - $($state.BaseDecade + 9)"

        $cellW = 64; $cellH = 60
        for ($i = 0; $i -lt 12; $i++) {
            $year = $startYear + $i
            $col = $i % 4
            $row = [int]([Math]::Floor($i / 4))

            $cell = New-Object System.Windows.Forms.Button
            $cell.Text = "$year"
            $cell.Size = New-Object System.Drawing.Size($cellW, $cellH)
            $cell.Location = New-Object System.Drawing.Point(($col * $cellW), ($row * $cellH))
            $cell.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
            $cell.FlatAppearance.BorderSize = 0
            $cell.Font = New-Object System.Drawing.Font('Segoe UI', 11)
            $cell.BackColor = [System.Drawing.Color]::White
            $cell.Tag = $year
            $cell.TabStop = $false

            $isOutOfDecade = ($year -lt $state.BaseDecade -or $year -gt ($state.BaseDecade + 9))
            $isSelected    = ($year -eq $state.SelectedYear)
            $isOutOfRange  = ($year -lt $state.MinYear -or $year -gt $state.MaxYear)

            if ($isSelected) {
                $cell.BackColor = [System.Drawing.Color]::FromArgb(66, 153, 225)
                $cell.ForeColor = [System.Drawing.Color]::White
                $cell.Font = New-Object System.Drawing.Font('Segoe UI', 11, [System.Drawing.FontStyle]::Bold)
            } elseif ($isOutOfDecade) {
                $cell.ForeColor = [System.Drawing.Color]::FromArgb(160, 160, 160)
            } else {
                $cell.ForeColor = [System.Drawing.Color]::FromArgb(45, 55, 72)
            }

            if ($isOutOfRange) {
                $cell.Enabled = $false
                $cell.ForeColor = [System.Drawing.Color]::FromArgb(210, 210, 210)
            }

            $cell.Add_Click($onYearClick)

            $gridPanel.Controls.Add($cell)
        }
    }.GetNewClosure()

    & $renderGrid

    $btnPrev.Add_Click({
        $newDecade = $state.BaseDecade - 10
        if (($newDecade + 9) -ge $state.MinYear) {
            $state.BaseDecade = $newDecade
            & $renderGrid
        }
    }.GetNewClosure())

    $btnNext.Add_Click({
        $newDecade = $state.BaseDecade + 10
        if ($newDecade -le $state.MaxYear) {
            $state.BaseDecade = $newDecade
            & $renderGrid
        }
    }.GetNewClosure())

    # Hospeda o painel root no dropdown
    $hostItem = New-Object System.Windows.Forms.ToolStripControlHost($root)
    $hostItem.AutoSize = $false
    $hostItem.Size     = New-Object System.Drawing.Size($popupW, $popupH)
    $hostItem.Margin   = [System.Windows.Forms.Padding]::Empty
    $hostItem.Padding  = [System.Windows.Forms.Padding]::Empty
    [void]$dropdown.Items.Add($hostItem)

    # Mostra logo abaixo do controle ancora
    $dropdown.Show($AnchorControl, 0, $AnchorControl.Height)
}

$splashIconPath = Join-Path $scriptDir 'assets\icon.ico'
Show-SplashScreen -AppName 'Planejamento de Ferias' -Author $AUTOR_FIXO -IconPath $splashIconPath -DurationMs 2500

# ================== Form ==================
$form = New-Object System.Windows.Forms.Form
$form.Text = 'Planejamento de Ferias - Gerador'
$form.Size = New-Object System.Drawing.Size(580, 365)
$form.StartPosition = 'CenterScreen'
$form.FormBorderStyle = 'FixedDialog'
$form.MaximizeBox = $false
$form.MinimizeBox = $true
$form.Font = New-Object System.Drawing.Font('Segoe UI', 9)
$form.BackColor = [System.Drawing.Color]::White

$iconPath = Join-Path $scriptDir 'assets\icon.ico'
if (Test-Path $iconPath) {
    $form.Icon = New-Object System.Drawing.Icon($iconPath)
}

# ---- Titulo ----
$lblTitle = New-Object System.Windows.Forms.Label
$lblTitle.Text = 'Planejamento de Ferias'
$lblTitle.Location = New-Object System.Drawing.Point(15, 12)
$lblTitle.Size = New-Object System.Drawing.Size(540, 24)
$lblTitle.Font = New-Object System.Drawing.Font('Segoe UI', 13, [System.Drawing.FontStyle]::Bold)
$lblTitle.ForeColor = [System.Drawing.Color]::FromArgb(45, 55, 72)
$form.Controls.Add($lblTitle)

# ---- Ano (editavel) ----
$lblAno = New-Object System.Windows.Forms.Label
$lblAno.Text = 'Ano:'
$lblAno.Location = New-Object System.Drawing.Point(15, 55)
$lblAno.Size = New-Object System.Drawing.Size(75, 22)
$form.Controls.Add($lblAno)

$btnAno = New-Object System.Windows.Forms.Button
$btnAno.Location = New-Object System.Drawing.Point(90, 53)
$btnAno.Size = New-Object System.Drawing.Size(100, 24)
$btnAno.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
$btnAno.FlatAppearance.BorderColor = [System.Drawing.Color]::FromArgb(180, 180, 180)
$btnAno.FlatAppearance.BorderSize = 1
$btnAno.BackColor = [System.Drawing.Color]::White
$btnAno.ForeColor = [System.Drawing.Color]::FromArgb(45, 55, 72)
$btnAno.Font = New-Object System.Drawing.Font('Segoe UI', 9)
$btnAno.TextAlign = [System.Drawing.ContentAlignment]::MiddleCenter
$btnAno.Tag = (Get-Date).Year
$btnAno.Text = "$((Get-Date).Year)   $([char]0x25BE)"
$btnAno.Cursor = [System.Windows.Forms.Cursors]::Hand
$form.Controls.Add($btnAno)

$btnAno.Add_Click({
    Show-YearPicker -AnchorControl $btnAno -CurrentYear ([int]$btnAno.Tag) -MinYear 2020 -MaxYear 2100
})

# ---- Planilha ----
$lblXlsx = New-Object System.Windows.Forms.Label
$lblXlsx.Text = 'Planilha:'
$lblXlsx.Location = New-Object System.Drawing.Point(15, 88)
$lblXlsx.Size = New-Object System.Drawing.Size(75, 22)
$form.Controls.Add($lblXlsx)

$txtXlsx = New-Object System.Windows.Forms.TextBox
$txtXlsx.Location = New-Object System.Drawing.Point(90, 86)
$txtXlsx.Size = New-Object System.Drawing.Size(375, 22)
$txtXlsx.Text = $xlsxDefault
$form.Controls.Add($txtXlsx)

$btnBrowse = New-Object System.Windows.Forms.Button
$btnBrowse.Text = 'Procurar...'
$btnBrowse.Location = New-Object System.Drawing.Point(470, 84)
$btnBrowse.Size = New-Object System.Drawing.Size(85, 26)
$btnBrowse.Add_Click({
    $dlg = New-Object System.Windows.Forms.OpenFileDialog
    $dlg.Filter = 'Excel (*.xlsx)|*.xlsx|CSV (*.csv)|*.csv|Todos (*.*)|*.*'
    $dlg.InitialDirectory = $scriptDir
    if ($dlg.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        $txtXlsx.Text = $dlg.FileName
    }
})
$form.Controls.Add($btnBrowse)

# ---- Checkbox abrir apos gerar ----
$chkOpen = New-Object System.Windows.Forms.CheckBox
$chkOpen.Text = 'Abrir apos gerar'
$chkOpen.Location = New-Object System.Drawing.Point(90, 120)
$chkOpen.Size = New-Object System.Drawing.Size(465, 22)
$chkOpen.Checked = $true
$form.Controls.Add($chkOpen)

# ---- Checkbox gerar PDF (compativel com SharePoint) ----
$chkPdf = New-Object System.Windows.Forms.CheckBox
$chkPdf.Text = 'Gerar PDF tambem (para upload no SharePoint)'
$chkPdf.Location = New-Object System.Drawing.Point(90, 145)
$chkPdf.Size = New-Object System.Drawing.Size(465, 22)
$chkPdf.Checked = $false
$form.Controls.Add($chkPdf)

# ---- Status (multilinha: mensagens longas como "Sucesso! Arquivos gerados: ..." podem precisar de 2 linhas) ----
$lblStatus = New-Object System.Windows.Forms.Label
$lblStatus.Location = New-Object System.Drawing.Point(15, 178)
$lblStatus.Size = New-Object System.Drawing.Size(540, 36)
$lblStatus.Text = 'Pronto. Confira os campos e clique em "Gerar Relatorio".'
$lblStatus.ForeColor = [System.Drawing.Color]::Gray
$form.Controls.Add($lblStatus)

# ---- Progress bar (marquee, visivel so durante o processamento) ----
$pb = New-Object System.Windows.Forms.ProgressBar
$pb.Location = New-Object System.Drawing.Point(15, 220)
$pb.Size = New-Object System.Drawing.Size(540, 14)
$pb.Style = [System.Windows.Forms.ProgressBarStyle]::Marquee
$pb.MarqueeAnimationSpeed = 30
$pb.Visible = $false
$form.Controls.Add($pb)

# ---- Rodape: "Criado por Carlos Seidl" ----
$lblAutoria = New-Object System.Windows.Forms.Label
$lblAutoria.Text = "Desenvolvido por $AUTOR_FIXO"
$lblAutoria.Location = New-Object System.Drawing.Point(15, 295)
$lblAutoria.Size = New-Object System.Drawing.Size(220, 20)
$lblAutoria.Font = New-Object System.Drawing.Font('Segoe UI', 8, [System.Drawing.FontStyle]::Italic)
$lblAutoria.ForeColor = [System.Drawing.Color]::Gray
$form.Controls.Add($lblAutoria)

# ---- Botao Fechar ----
$btnClose = New-Object System.Windows.Forms.Button
$btnClose.Text = 'Fechar'
$btnClose.Location = New-Object System.Drawing.Point(327, 290)
$btnClose.Size = New-Object System.Drawing.Size(85, 32)
$btnClose.Font = New-Object System.Drawing.Font('Segoe UI', 9)
$btnClose.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
$btnClose.FlatAppearance.BorderColor = [System.Drawing.Color]::FromArgb(203, 213, 224)
$btnClose.FlatAppearance.BorderSize = 1
$btnClose.BackColor = [System.Drawing.Color]::White
$btnClose.ForeColor = [System.Drawing.Color]::FromArgb(45, 55, 72)
$btnClose.Add_Click({ $form.Close() })
$form.Controls.Add($btnClose)

# ---- Botao Gerar ----
$btnGenerate = New-Object System.Windows.Forms.Button
$btnGenerate.Text = 'Gerar Relatorio'
$btnGenerate.Location = New-Object System.Drawing.Point(420, 290)
$btnGenerate.Size = New-Object System.Drawing.Size(135, 32)
$btnGenerate.Font = New-Object System.Drawing.Font('Segoe UI', 9, [System.Drawing.FontStyle]::Bold)
$btnGenerate.BackColor = [System.Drawing.Color]::FromArgb(66, 153, 225)
$btnGenerate.ForeColor = [System.Drawing.Color]::White
$btnGenerate.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
$btnGenerate.FlatAppearance.BorderSize = 0

$btnGenerate.Add_Click({
    $autor = $AUTOR_FIXO
    $ano   = [int]$btnAno.Tag
    $xlsx  = $txtXlsx.Text.Trim()

    if (-not (Test-Path $xlsx)) {
        [System.Windows.Forms.MessageBox]::Show("Planilha nao encontrada:`n$xlsx",
            'Erro', 'OK', 'Error') | Out-Null
        return
    }

    $btnGenerate.Enabled = $false
    $btnClose.Enabled    = $false
    $btnBrowse.Enabled   = $false
    $btnAno.Enabled      = $false
    $txtXlsx.Enabled     = $false
    $chkPdf.Enabled      = $false
    $chkOpen.Enabled     = $false
    $form.UseWaitCursor  = $true
    $pb.Visible          = $true
    $pb.MarqueeAnimationSpeed = 30
    $msgBase = if ($chkPdf.Checked) { 'Gerando relatorio HTML + PDF...' } else { 'Gerando relatorio...' }
    $lblStatus.Text = "$msgBase Aguarde (primeira execucao pode demorar)."
    $lblStatus.ForeColor = [System.Drawing.Color]::FromArgb(43, 108, 176)
    $form.Refresh()
    [System.Windows.Forms.Application]::DoEvents()

    try {
        & $executar -XlsxPath $xlsx -Autor $autor -Ano $ano -Pdf:$chkPdf.Checked *>&1 | Out-Null

        # Procura o ultimo HTML gerado; se PDF marcado, prefere abrir o PDF.
        # Cada execucao cria results\Ferias-{timestamp}\ com seus 3-4 arquivos
        # dentro, entao buscamos recursivamente (e -Recurse tambem cobre
        # arquivos antigos que ficaram soltos no formato anterior).
        $latestHtml = Get-ChildItem -Path $resultsDir -Filter 'Ferias-*.html' -Recurse -File -ErrorAction SilentlyContinue |
                      Sort-Object LastWriteTime -Descending | Select-Object -First 1
        $latestPdf  = Get-ChildItem -Path $resultsDir -Filter 'Ferias-*.pdf'  -Recurse -File -ErrorAction SilentlyContinue |
                      Sort-Object LastWriteTime -Descending | Select-Object -First 1

        if (-not $latestHtml) {
            throw "Script rodou mas nenhum HTML foi encontrado em:`n$resultsDir"
        }

        $generatedNames = @($latestHtml.Name)
        if ($chkPdf.Checked -and $latestPdf) { $generatedNames += $latestPdf.Name }
        $lblStatus.Text = "Sucesso! Arquivos gerados: " + ($generatedNames -join ', ')
        $lblStatus.ForeColor = [System.Drawing.Color]::FromArgb(47, 133, 90)

        if ($chkOpen.Checked) {
            $toOpen = if ($chkPdf.Checked -and $latestPdf) { $latestPdf.FullName } else { $latestHtml.FullName }
            Start-Process $toOpen
        }
    }
    catch {
        $lblStatus.Text = 'Erro ao gerar relatorio. Veja detalhes na janela.'
        $lblStatus.ForeColor = [System.Drawing.Color]::FromArgb(197, 48, 48)
        [System.Windows.Forms.MessageBox]::Show($_.Exception.Message,
            'Erro ao gerar relatorio', 'OK', 'Error') | Out-Null
    }
    finally {
        $btnGenerate.Enabled = $true
        $btnClose.Enabled    = $true
        $btnBrowse.Enabled   = $true
        $btnAno.Enabled      = $true
        $txtXlsx.Enabled     = $true
        $chkPdf.Enabled      = $true
        $chkOpen.Enabled     = $true
        $pb.Visible          = $false
        $pb.MarqueeAnimationSpeed = 0
        $form.UseWaitCursor  = $false
    }
})
$form.Controls.Add($btnGenerate)

$form.AcceptButton = $btnGenerate
$form.CancelButton = $btnClose
[void]$form.ShowDialog()
$form.Dispose()
