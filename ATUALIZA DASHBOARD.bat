@echo off
chcp 65001 > nul
title Atualiza Dashboard - Logistica Serena

echo ================================================
echo   ATUALIZA DASHBOARD - LOGISTICA SERENA
echo   %date% %time%
echo ================================================
echo.

cd /d "%~dp0"

if not exist "%~dp0logs" mkdir "%~dp0logs"
set LOG=%~dp0logs\dashboard_%date:~6,4%%date:~3,2%%date:~0,2%.txt

echo [%date% %time%] Iniciando atualizacao do Dashboard >> "%LOG%"

python "%~dp0atualizar_dashboard.py" --so-dashboard >> "%LOG%" 2>&1

if %errorlevel% neq 0 (
    echo     ERRO! Veja o log em: %LOG%
    echo [%time%] ERRO %errorlevel% >> "%LOG%"
) else (
    echo     Dashboard atualizado com sucesso!
    echo [%time%] OK >> "%LOG%"
)

echo.
echo ================================================
echo   Concluido! Abra o dashboard_logistica.html
echo   no navegador para ver as atualizacoes.
echo ================================================
echo.
echo Pressione qualquer tecla para fechar...
pause > nul
