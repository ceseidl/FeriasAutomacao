#requires -Version 5.1
<#
.SYNOPSIS
    Ghost test end-to-end: roda executar.ps1 contra a planilha ghost
    (tests/ghost-data/ferias-ghost-mixed.xlsx) que tem 6 pessoas
    imaginarias em todos os anos 2023-2029.

.DESCRIPTION
    Diferente de Test-YearFilter-EndToEnd.ps1 (que usa a planilha real
    com so 2026), aqui esperamos SUCESSO em todos os 7 anos. Cada
    geracao deve filtrar 42 -> 6 linhas (so as do ano selecionado).

    Validacoes por ano:
      1. executar.ps1 termina com sucesso
      2. HTML foi gerado em results/
      3. O HTML contem o ano correto no titulo
      4. O HTML contem os 6 colaboradores ghost (Ana, Bruno, Camila,
         Diego, Eduarda, Fernando)
      5. O HTML NAO contem nenhuma data de outro ano (ex: rodando 2025
         nao deve aparecer "/2024" ou "/2026" em nenhuma linha de tabela)

    Pre-requisito: rodar tests/ghost-data/Generate-GhostXlsx.ps1 primeiro.
#>

$ErrorActionPreference = 'Stop'
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8

$repoRoot   = Split-Path -Parent $PSScriptRoot
$executar   = Join-Path $repoRoot 'executar.ps1'
$ghostXlsx  = Join-Path $PSScriptRoot 'ghost-data\ferias-ghost-mixed.xlsx'
$resultsDir = Join-Path $repoRoot 'results'

if (-not (Test-Path $ghostXlsx)) {
    Write-Host "Planilha ghost nao encontrada. Gerando agora..." -ForegroundColor Yellow
    & (Join-Path $PSScriptRoot 'ghost-data\Generate-GhostXlsx.ps1')
}

$colaboradoresEsperados = @(
    'Ana Beatriz Souza',
    'Bruno Ferreira Lima',
    'Camila Oliveira Costa',
    'Diego Rodrigues Santos',
    'Eduarda Pereira Mendes',
    'Fernando Almeida Rocha'
)

$pass = 0
$fail = 0

Write-Host ""
Write-Host "==> Planilha ghost: 42 linhas (6 pessoas x 7 anos 2023-2029)"
Write-Host "==> Esperado: SUCESSO em todos os 7 anos, cada um com 6 linhas filtradas"
Write-Host ""

foreach ($ano in 2023..2029) {
    Write-Host ("---- Ano {0} ----" -f $ano)

    $erro = $null
    $sucesso = $false
    try {
        & $executar -XlsxPath $ghostXlsx -Ano $ano -Autor 'Ghost Runner' *>&1 | Out-Null
        $sucesso = $true
    } catch {
        $erro = $_.Exception.Message
    }

    if (-not $sucesso) {
        Write-Host ("    FAIL  Esperava sucesso mas deu erro: {0}" -f $erro) -ForegroundColor Red
        $fail++
        continue
    }

    # Pega o HTML mais recente
    $latest = Get-ChildItem -Path $resultsDir -Filter 'Ferias-*.html' -ErrorAction SilentlyContinue |
              Sort-Object LastWriteTime -Descending | Select-Object -First 1
    if (-not $latest) {
        Write-Host "    FAIL  Sucesso reportado mas nenhum HTML em results/" -ForegroundColor Red
        $fail++
        continue
    }

    $html = Get-Content -Path $latest.FullName -Raw -Encoding UTF8

    # Valida: ano correto no titulo
    if ($html -notmatch "Planejamento de F[eé]rias.*$ano") {
        Write-Host ("    FAIL  HTML nao contem 'Planejamento de Ferias $ano' no titulo" -f $ano) -ForegroundColor Red
        $fail++
        continue
    }

    # Valida: todos os 6 colaboradores aparecem
    $faltando = $colaboradoresEsperados | Where-Object { $html -notmatch [regex]::Escape($_) }
    if ($faltando) {
        Write-Host ("    FAIL  Colaboradores ausentes no HTML: {0}" -f ($faltando -join ', ')) -ForegroundColor Red
        $fail++
        continue
    }

    # Valida: nenhuma data dd/MM/YYYY de outro ano aparece em <td> da tabela.
    # Olha apenas celulas de tabela (cronograma). Ignora o rodape com a data
    # de geracao (que esta em <em>...</em>, nao em <td>).
    $tdDates = [regex]::Matches($html, '<td[^>]*>\s*(\d{2}/\d{2}/(\d{4}))\s*</td>')
    $vazamentos = @{}
    foreach ($m in $tdDates) {
        $yr = [int]$m.Groups[2].Value
        if ($yr -ne $ano) {
            if (-not $vazamentos.ContainsKey($yr)) { $vazamentos[$yr] = 0 }
            $vazamentos[$yr]++
        }
    }
    if ($vazamentos.Keys.Count -gt 0) {
        $detalhe = ($vazamentos.GetEnumerator() | ForEach-Object { "$($_.Key)x$($_.Value)" }) -join ', '
        Write-Host ("    FAIL  HTML tem datas de outros anos em celulas de tabela: {0}" -f $detalhe) -ForegroundColor Red
        $fail++
        continue
    }

    # Conta tambem datas ISO no Gantt (YYYY-MM-DD) pra cobrir o bloco mermaid
    $isoDates = [regex]::Matches($html, '\b(\d{4})-\d{2}-\d{2}\b')
    $vazIso = @{}
    foreach ($m in $isoDates) {
        $yr = [int]$m.Groups[1].Value
        if ($yr -lt 2020 -or $yr -gt 2099) { continue }  # ignora anos fora do range esperado
        if ($yr -ne $ano) {
            if (-not $vazIso.ContainsKey($yr)) { $vazIso[$yr] = 0 }
            $vazIso[$yr]++
        }
    }
    if ($vazIso.Keys.Count -gt 0) {
        $detalhe = ($vazIso.GetEnumerator() | ForEach-Object { "$($_.Key)x$($_.Value)" }) -join ', '
        Write-Host ("    FAIL  HTML tem datas ISO de outros anos no Gantt: {0}" -f $detalhe) -ForegroundColor Red
        $fail++
        continue
    }

    Write-Host ("    PASS  HTML: {0}" -f $latest.Name) -ForegroundColor Green
    Write-Host ("          - 6 colaboradores presentes" )                                       -ForegroundColor DarkGray
    Write-Host ("          - Nenhuma data de outro ano vazou" )                                 -ForegroundColor DarkGray
    $pass++
}

Write-Host ""
Write-Host "============================================================"
Write-Host ("RESUMO Ghost E2E: {0} PASS / {1} FAIL" -f $pass, $fail)
Write-Host "============================================================"

if ($fail -gt 0) { exit 1 } else { exit 0 }
