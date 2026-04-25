#requires -Version 5.1
<#
.SYNOPSIS
    Ghost test do filtro por ano (regra: relatorio so traz ferias do ano
    selecionado no picker, independente do que a planilha tem).

.DESCRIPTION
    Cria linhas sinteticas com Inicios em 2023, 2024, 2025, 2026, 2027,
    2028 e 2029 (3 anos pra tras + ano corrente + 3 anos pra frente),
    aplica a mesma logica de filtro do executar.ps1 e verifica:

      1. Apenas as linhas com Inicio no ano selecionado sao mantidas.
      2. A contagem bate com o esperado.
      3. Anos sem dados retornam zero linhas (sem misturar com vizinhos).
      4. Mistura de anos na mesma planilha nao "contamina" o resultado.

    Nao chama o executar.ps1 inteiro (Pandoc, Edge, etc.); exercita apenas
    a logica pura do filtro. Rapido e isolado.

.EXAMPLE
    .\tests\Test-YearFilter.ps1
#>

$ErrorActionPreference = 'Stop'
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8

# Funcao identica a do executar.ps1 (mantida em sincronia manual).
function ConvertTo-IsoDate([string]$br) {
    $dt = [datetime]::MinValue
    if ([datetime]::TryParseExact($br, 'dd/MM/yyyy', [System.Globalization.CultureInfo]::InvariantCulture, [System.Globalization.DateTimeStyles]::None, [ref]$dt)) {
        return $dt.ToString('yyyy-MM-dd')
    }
    return $br
}

# Replica do filtro do executar.ps1 (qualquer mudanca la deve refletir aqui).
function Filter-RowsByYear {
    param(
        [object[]] $Rows,
        [int]      $Ano
    )
    return @($Rows | Where-Object {
        $iso = ConvertTo-IsoDate $_.Inicio
        $iso -and $iso.StartsWith("$Ano-")
    })
}

# ============================================================
# Massa de teste: 2 linhas por ano de 2023 a 2029
# ============================================================
$mockRows = @()
foreach ($yr in 2023..2029) {
    $mockRows += [pscustomobject]@{
        Mes = 'Janeiro'; Colaborador = "Test A $yr"; Squad = 'Cross'
        Inicio = "10/01/$yr"; Fim = "24/01/$yr"; Dias = 15; Status = 'Aprovada'
    }
    $mockRows += [pscustomobject]@{
        Mes = 'Julho'; Colaborador = "Test B $yr"; Squad = 'Frontend & CMS'
        Inicio = "05/07/$yr"; Fim = "19/07/$yr"; Dias = 15; Status = 'Planejada'
    }
}

Write-Host "==> Massa de teste: $($mockRows.Count) linhas (2 por ano de 2023 a 2029)"
Write-Host ""

$pass = 0
$fail = 0

# ============================================================
# Caso 1: para cada ano de 2023 a 2029, esperar 2 linhas, todas
# com Inicio naquele ano.
# ============================================================
Write-Host "[1] Filtro por ano: 3 anos pra tras + ano corrente + 3 anos pra frente"
foreach ($Ano in 2023..2029) {
    $filtered = Filter-RowsByYear -Rows $mockRows -Ano $Ano
    $expected = 2

    $allMatch = $true
    foreach ($r in $filtered) {
        $isoYr = (ConvertTo-IsoDate $r.Inicio).Substring(0, 4)
        if ($isoYr -ne "$Ano") { $allMatch = $false; break }
    }

    if ($filtered.Count -eq $expected -and $allMatch) {
        Write-Host ("    PASS  Ano {0}: {1} linhas, todas em {0}" -f $Ano, $filtered.Count) -ForegroundColor Green
        $pass++
    } else {
        Write-Host ("    FAIL  Ano {0}: esperado {1} linhas em {0}, recebido {2} (allMatch={3})" -f $Ano, $expected, $filtered.Count, $allMatch) -ForegroundColor Red
        $fail++
    }
}

Write-Host ""

# ============================================================
# Caso 2: ano que nao existe na planilha -> zero linhas
# ============================================================
Write-Host "[2] Ano sem dados na planilha -> zero linhas"
$filtered = Filter-RowsByYear -Rows $mockRows -Ano 2099
if ($filtered.Count -eq 0) {
    Write-Host "    PASS  Ano 2099: 0 linhas (correto)" -ForegroundColor Green
    $pass++
} else {
    Write-Host ("    FAIL  Ano 2099: esperado 0, recebido {0}" -f $filtered.Count) -ForegroundColor Red
    $fail++
}

Write-Host ""

# ============================================================
# Caso 3: planilha com varios anos misturados -> filtro nao
# vaza linhas de outros anos (regra principal pedida)
# ============================================================
Write-Host "[3] Mistura de anos na mesma planilha nao contamina o resultado"
foreach ($Ano in 2023..2029) {
    $filtered = Filter-RowsByYear -Rows $mockRows -Ano $Ano
    $bleedThrough = $filtered | Where-Object {
        $isoYr = (ConvertTo-IsoDate $_.Inicio).Substring(0, 4)
        $isoYr -ne "$Ano"
    }
    if (-not $bleedThrough -or @($bleedThrough).Count -eq 0) {
        Write-Host ("    PASS  Ano {0}: nenhuma linha de outro ano vazou" -f $Ano) -ForegroundColor Green
        $pass++
    } else {
        Write-Host ("    FAIL  Ano {0}: {1} linhas de outro ano vazaram" -f $Ano, @($bleedThrough).Count) -ForegroundColor Red
        $fail++
    }
}

Write-Host ""

# ============================================================
# Caso 4: linha com data invalida nao deve quebrar nem ser
# incluida em nenhum ano.
# ============================================================
Write-Host "[4] Linhas com data invalida sao ignoradas (nao incluidas em nenhum ano)"
$mockComLixo = $mockRows + [pscustomobject]@{
    Mes = 'Janeiro'; Colaborador = 'Linha Lixo'; Squad = 'X'
    Inicio = 'data-invalida'; Fim = ''; Dias = 0; Status = 'Aprovada'
}
$totalApos = 0
foreach ($Ano in 2023..2029) {
    $totalApos += (Filter-RowsByYear -Rows $mockComLixo -Ano $Ano).Count
}
# Esperado: 14 (2 linhas x 7 anos), a linha lixo nao bate com nenhum
if ($totalApos -eq 14) {
    Write-Host "    PASS  Linha invalida foi ignorada (14 linhas validas mantidas)" -ForegroundColor Green
    $pass++
} else {
    Write-Host ("    FAIL  Esperado 14 linhas validas, recebido {0}" -f $totalApos) -ForegroundColor Red
    $fail++
}

Write-Host ""
Write-Host "============================================================"
Write-Host ("RESUMO: {0} PASS / {1} FAIL" -f $pass, $fail)
Write-Host "============================================================"

if ($fail -gt 0) { exit 1 } else { exit 0 }
