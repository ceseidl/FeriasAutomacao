#requires -Version 5.1
<#
.SYNOPSIS
    Gera tests/ghost-data/ferias-ghost-mixed.xlsx com 6 pessoas imaginarias
    e ferias cadastradas em todos os anos de 2023 a 2029 (42 linhas no total).

.DESCRIPTION
    Usado pelo ghost test end-to-end pra validar que o filtro por ano em
    executar.ps1 funciona quando a planilha tem dados misturados de varios
    anos. Cada pessoa tira ferias uma vez por ano, sempre no mesmo mes.

    Cada execucao recria o arquivo do zero.
#>

$ErrorActionPreference = 'Stop'
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8

if (-not (Get-Module -ListAvailable -Name ImportExcel)) {
    Write-Host "Instalando modulo ImportExcel no escopo do usuario..." -ForegroundColor Yellow
    Install-Module ImportExcel -Scope CurrentUser -Force -ErrorAction Stop
}
Import-Module ImportExcel -ErrorAction Stop | Out-Null

$outDir = $PSScriptRoot
$out    = Join-Path $outDir 'ferias-ghost-mixed.xlsx'

# Pessoas imaginarias com squad e mes "preferido" de ferias
$pessoas = @(
    @{ Nome = 'Ana Beatriz Souza';      Squad = 'Cross';                    Mes = 'Janeiro';   StartDay = 10; StartMonth = 1  },
    @{ Nome = 'Bruno Ferreira Lima';    Squad = 'Frontend & CMS';           Mes = 'Marco';     StartDay = 6;  StartMonth = 3  },
    @{ Nome = 'Camila Oliveira Costa';  Squad = 'Jornada de Compras';       Mes = 'Maio';      StartDay = 2;  StartMonth = 5  },
    @{ Nome = 'Diego Rodrigues Santos'; Squad = 'Pedido/OMS - Integracao';  Mes = 'Julho';     StartDay = 3;  StartMonth = 7  },
    @{ Nome = 'Eduarda Pereira Mendes'; Squad = 'Backend';                  Mes = 'Setembro';  StartDay = 4;  StartMonth = 9  },
    @{ Nome = 'Fernando Almeida Rocha'; Squad = 'Mobile';                   Mes = 'Novembro';  StartDay = 6;  StartMonth = 11 }
)

$statusCiclo = @('Aprovada', 'Solicitada', 'Planejada')

$rows = New-Object System.Collections.Generic.List[object]
$idx  = 0
foreach ($yr in 2023..2029) {
    foreach ($p in $pessoas) {
        $start = [datetime]::new($yr, $p.StartMonth, $p.StartDay)
        $dias  = 15
        $end   = $start.AddDays($dias - 1)
        $rows.Add([pscustomobject]@{
            Mes         = $p.Mes
            Colaborador = $p.Nome
            Squad       = $p.Squad
            Inicio      = $start.ToString('dd/MM/yyyy')
            Fim         = $end.ToString('dd/MM/yyyy')
            Dias        = $dias
            Status      = $statusCiclo[$idx % 3]
        }) | Out-Null
        $idx++
    }
}

if (Test-Path $out) { Remove-Item $out -Force }

$rows | Export-Excel -Path $out -WorksheetName 'Ferias' -AutoSize -FreezeTopRow -BoldTopRow

Write-Host ("Gerado: {0} ({1} linhas, {2} pessoas, anos 2023-2029)" -f $out, $rows.Count, $pessoas.Count)
