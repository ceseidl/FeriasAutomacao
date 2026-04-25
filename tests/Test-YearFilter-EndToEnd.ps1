#requires -Version 5.1
<#
.SYNOPSIS
    Smoke test end-to-end: roda executar.ps1 para 7 anos contra a planilha
    real (que so tem 2026) e valida o comportamento esperado.

.DESCRIPTION
    Esperado:
      - 2023, 2024, 2025: throw "Nenhuma linha ... para o ano <X>"
      - 2026:             sucesso (HTML gerado em results/)
      - 2027, 2028, 2029: throw "Nenhuma linha ... para o ano <X>"

    O teste captura stderr/stdout, conta sucessos e falhas, e imprime
    PASS/FAIL pra cada ano. Nao gera PDF (so HTML, mais rapido).
#>

$ErrorActionPreference = 'Stop'
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8

$repoRoot = Split-Path -Parent $PSScriptRoot
$executar = Join-Path $repoRoot 'executar.ps1'

$cases = @(
    @{ Ano = 2023; Esperado = 'erro' },
    @{ Ano = 2024; Esperado = 'erro' },
    @{ Ano = 2025; Esperado = 'erro' },
    @{ Ano = 2026; Esperado = 'sucesso' },
    @{ Ano = 2027; Esperado = 'erro' },
    @{ Ano = 2028; Esperado = 'erro' },
    @{ Ano = 2029; Esperado = 'erro' }
)

$pass = 0
$fail = 0

Write-Host ""
Write-Host "==> Planilha: ferias-2026.xlsx (so contem 2026)"
Write-Host "==> Rodando 7 cenarios (3 pra tras, ano atual, 3 pra frente)"
Write-Host ""

foreach ($c in $cases) {
    $ano = $c.Ano
    $esp = $c.Esperado
    Write-Host ("---- Ano {0} (esperado: {1}) ----" -f $ano, $esp)

    $erro = $null
    $sucesso = $false
    try {
        & $executar -Ano $ano -Autor 'Test Runner' *>&1 | Out-Null
        $sucesso = $true
    } catch {
        $erro = $_.Exception.Message
    }

    $okErro = ($esp -eq 'erro' -and -not $sucesso -and $erro -match "Nenhuma linha de ferias para o ano $ano")
    $okSucesso = ($esp -eq 'sucesso' -and $sucesso)

    if ($okErro) {
        Write-Host ("    PASS  Erro esperado capturado:") -ForegroundColor Green
        Write-Host ("          {0}" -f ($erro.Trim() -replace "`r?`n", ' ')) -ForegroundColor DarkGray
        $pass++
    } elseif ($okSucesso) {
        $latest = Get-ChildItem -Path (Join-Path $repoRoot 'results') -Filter 'Ferias-*.html' -ErrorAction SilentlyContinue |
                  Sort-Object LastWriteTime -Descending | Select-Object -First 1
        Write-Host ("    PASS  HTML gerado: {0}" -f ($latest.Name)) -ForegroundColor Green
        $pass++
    } elseif ($esp -eq 'erro' -and $sucesso) {
        Write-Host ("    FAIL  Esperava erro pra ano {0} mas executou com sucesso" -f $ano) -ForegroundColor Red
        $fail++
    } elseif ($esp -eq 'sucesso' -and -not $sucesso) {
        Write-Host ("    FAIL  Esperava sucesso pra ano {0} mas falhou: {1}" -f $ano, $erro) -ForegroundColor Red
        $fail++
    } else {
        Write-Host ("    FAIL  Ano {0}: erro inesperado: {1}" -f $ano, $erro) -ForegroundColor Red
        $fail++
    }

    Write-Host ""
}

Write-Host "============================================================"
Write-Host ("RESUMO E2E: {0} PASS / {1} FAIL" -f $pass, $fail)
Write-Host "============================================================"

if ($fail -gt 0) { exit 1 } else { exit 0 }
