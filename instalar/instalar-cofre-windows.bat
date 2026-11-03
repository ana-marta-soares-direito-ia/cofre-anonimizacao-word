@echo off
chcp 65001 >nul
title Cofre de Anonimização para Word
echo.
echo   Cofre de Anonimização para Word
echo   A instalar...
echo.
taskkill /IM WINWORD.EXE >nul 2>&1
set "DIR=%LOCALAPPDATA%\CofreAnonimizacao"
if not exist "%DIR%" mkdir "%DIR%"
powershell -NoProfile -ExecutionPolicy Bypass -Command "[Net.ServicePointManager]::SecurityProtocol=[Net.SecurityProtocolType]::Tls12; Invoke-WebRequest -UseBasicParsing -Uri 'https://ana-marta-soares-direito-ia.github.io/cofre-anonimizacao-word/manifest.xml' -OutFile '%DIR%\manifest.xml'"
if not exist "%DIR%\manifest.xml" goto erro
reg add "HKCU\Software\Microsoft\Office\16.0\WEF\Developer" /v "df844483-b33b-4fb4-a28a-196c614ce8fd" /t REG_SZ /d "%DIR%\manifest.xml" /f >nul
if errorlevel 1 goto erro
echo   Instalação concluída.
echo.
echo   Abra o Word e um documento. No separador Base, seleccione Suplementos
echo   e escolha Cofre de Anonimização. Depois, cole o seu código de licença.
echo.
pause
exit /b 0
:erro
echo   Não foi possível concluir a instalação. Verifique a ligação à Internet e tente novamente.
echo.
pause
exit /b 1
