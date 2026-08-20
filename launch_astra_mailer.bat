@echo off
title Astra Mailer
cd /d "%~dp0"
echo.
echo  Starting Astra Mailer...
echo.
py -3 "%~dp0astra_mailer.py"
if %ERRORLEVEL% neq 0 (
    echo.
    echo  -----------------------------------------------
    echo  Something went wrong. Trying alternate Python...
    echo  -----------------------------------------------
    echo.
    python "%~dp0astra_mailer.py"
)
if %ERRORLEVEL% neq 0 (
    echo.
    echo  Astra Mailer could not start.
    echo  Try running INSTALL.bat again.
    echo.
    pause
)
