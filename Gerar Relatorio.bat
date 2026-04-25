@echo off
REM Lancador do gerador de Relatorio de Ferias 2026.
REM Duplo-clique abre a janela. Esconde o console pra ficar limpo.
start "" /b powershell.exe -NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -File "%~dp0gui.ps1"
