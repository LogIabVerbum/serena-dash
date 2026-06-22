@echo off
chcp 65001 > nul
title Atualiza Dashboard - Logistica Serena

echo ================================================
echo   ATUALIZA DASHBOARD - LOGISTICA SERENA
echo   %date% %time%
echo ================================================
echo.

cd /d "%~dp0"

python "%~dp0atualizar_dashboard.py" --so-dashboard

if %errorlevel% neq 0 (
    echo     ERRO! Verifique a pasta logs.
) else (
    echo     Dashboard atualizado com sucesso!
)

echo.
echo ================================================
echo   Concluido! Abra o dashboard_logistica.html
echo   no navegador para ver as atualizacoes.
echo ================================================
echo.
echo Pressione qualquer tecla para fechar...
pause > nul
