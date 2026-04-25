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
$xlsxDefault = Join-Path $scriptDir 'ferias-2026.xlsx'
$resultsDir  = Join-Path $scriptDir 'results'

if (-not (Test-Path $executar)) {
    [System.Windows.Forms.MessageBox]::Show(
        "executar.ps1 nao encontrado em:`n$scriptDir",
        'Erro de configuracao', 'OK', 'Error') | Out-Null
    exit 1
}

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

$nudAno = New-Object System.Windows.Forms.NumericUpDown
$nudAno.Location = New-Object System.Drawing.Point(90, 53)
$nudAno.Size = New-Object System.Drawing.Size(90, 22)
$nudAno.Minimum = 2020
$nudAno.Maximum = 2100
$nudAno.Value = (Get-Date).Year
$nudAno.TextAlign = 'Center'
$form.Controls.Add($nudAno)

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
$lblAutoria.Text = "Criado por $AUTOR_FIXO"
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
    $ano   = [int]$nudAno.Value
    $xlsx  = $txtXlsx.Text.Trim()

    if (-not (Test-Path $xlsx)) {
        [System.Windows.Forms.MessageBox]::Show("Planilha nao encontrada:`n$xlsx",
            'Erro', 'OK', 'Error') | Out-Null
        return
    }

    $btnGenerate.Enabled = $false
    $btnClose.Enabled    = $false
    $btnBrowse.Enabled   = $false
    $nudAno.Enabled      = $false
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

        # Procura o ultimo HTML gerado; se PDF marcado, prefere abrir o PDF
        $latestHtml = Get-ChildItem -Path $resultsDir -Filter 'Ferias-*.html' -ErrorAction SilentlyContinue |
                      Sort-Object LastWriteTime -Descending | Select-Object -First 1
        $latestPdf  = Get-ChildItem -Path $resultsDir -Filter 'Ferias-*.pdf'  -ErrorAction SilentlyContinue |
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
        $nudAno.Enabled      = $true
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
