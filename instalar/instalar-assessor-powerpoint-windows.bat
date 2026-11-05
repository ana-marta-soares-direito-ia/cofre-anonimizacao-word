@echo off
chcp 65001 >nul
title Assessor para PowerPoint
echo.
echo   Assessor para PowerPoint
echo   A instalar...
echo.
taskkill /IM POWERPNT.EXE >nul 2>&1
set "DIR=%LOCALAPPDATA%\AssessorPowerPoint"
if not exist "%DIR%" mkdir "%DIR%"
powershell -NoProfile -ExecutionPolicy Bypass -Command "[Net.ServicePointManager]::SecurityProtocol=[Net.SecurityProtocolType]::Tls12; Invoke-WebRequest -UseBasicParsing -Uri 'https://ana-marta-soares-direito-ia.github.io/cofre-anonimizacao-word/powerpoint/manifest.xml' -OutFile '%DIR%\manifest.xml'"
if not exist "%DIR%\manifest.xml" goto erro
reg add "HKCU\Software\Microsoft\Office\16.0\WEF\Developer" /v "4d0f6a71-8c2e-4b9d-9e35-2f1a7c6b8e04" /t REG_SZ /d "%DIR%\manifest.xml" /f >nul
if errorlevel 1 goto erro
echo   Instalação concluída.
echo.
echo   Abra o PowerPoint. No separador Base, seleccione Suplementos
echo   e escolha Assessor para PowerPoint. Depois, cole o seu código de licença.
echo.
pause
exit /b 0
:erro
echo   Não foi possível concluir a instalação. Verifique a ligação à Internet e tente novamente.
echo.
pause
exit /b 1
