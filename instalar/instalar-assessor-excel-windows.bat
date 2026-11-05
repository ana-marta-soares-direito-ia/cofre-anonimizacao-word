@echo off
chcp 65001 >nul
title Assessor para Excel
echo.
echo   Assessor para Excel
echo   A instalar...
echo.
taskkill /IM EXCEL.EXE >nul 2>&1
set "DIR=%LOCALAPPDATA%\AssessorExcel"
if not exist "%DIR%" mkdir "%DIR%"
powershell -NoProfile -ExecutionPolicy Bypass -Command "[Net.ServicePointManager]::SecurityProtocol=[Net.SecurityProtocolType]::Tls12; Invoke-WebRequest -UseBasicParsing -Uri 'https://ana-marta-soares-direito-ia.github.io/cofre-anonimizacao-word/excel/manifest.xml' -OutFile '%DIR%\manifest.xml'"
if not exist "%DIR%\manifest.xml" goto erro
reg add "HKCU\Software\Microsoft\Office\16.0\WEF\Developer" /v "9b4b6e2c-31f7-4f0e-a1c5-6d3e0b7a2f51" /t REG_SZ /d "%DIR%\manifest.xml" /f >nul
if errorlevel 1 goto erro
echo   Instalação concluída.
echo.
echo   Abra o Excel. No separador Base, seleccione Suplementos
echo   e escolha Assessor para Excel. Depois, cole o seu código de licença.
echo.
pause
exit /b 0
:erro
echo   Não foi possível concluir a instalação. Verifique a ligação à Internet e tente novamente.
echo.
pause
exit /b 1
