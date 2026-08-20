@echo off
setlocal EnableDelayedExpansion

:: ============================================================================
::  Astra Mailer Installer for Windows
::  Double-click this file and follow the prompts.
:: ============================================================================

title Astra Mailer — Installer
color 1F

:: Get the folder this script is in (remove trailing backslash for safety)
set "APP_DIR=%~dp0"
if "!APP_DIR:~-1!"=="\" set "APP_DIR=!APP_DIR:~0,-1!"

echo.
echo  ============================================
echo     Welcome to Astra Mailer Installer
echo  ============================================
echo.
echo  App folder: !APP_DIR!
echo.
echo  This will set up everything you need.
echo  It should take about 2-5 minutes.
echo.
echo  DO NOT close this window until it says "complete".
echo.
pause

:: ── Step 1: Find Python ────────────────────────────────────────────────────

echo.
echo  ============================================
echo  [Step 1/4] Checking for Python...
echo  ============================================
echo.

set "PYTHON_CMD="

:: --- Try "py" launcher first (most reliable on Windows) ---
py --version >nul 2>&1
if !ERRORLEVEL! equ 0 (
    for /f "tokens=*" %%v in ('py --version 2^>^&1') do echo  Found: %%v
    set "PYTHON_CMD=py"
    goto :python_found
)

:: --- Try "python" but verify it's real ---
python --version > "%TEMP%\pycheck.txt" 2>&1
if !ERRORLEVEL! equ 0 (
    findstr /i "Python 3" "%TEMP%\pycheck.txt" >nul 2>&1
    if !ERRORLEVEL! equ 0 (
        for /f "tokens=*" %%v in ('type "%TEMP%\pycheck.txt"') do echo  Found: %%v
        set "PYTHON_CMD=python"
        del "%TEMP%\pycheck.txt" >nul 2>&1
        goto :python_found
    )
)
del "%TEMP%\pycheck.txt" >nul 2>&1

:: --- Try python3 ---
python3 --version >nul 2>&1
if !ERRORLEVEL! equ 0 (
    for /f "tokens=*" %%v in ('python3 --version 2^>^&1') do echo  Found: %%v
    set "PYTHON_CMD=python3"
    goto :python_found
)

:: --- Try common install locations ---
for %%P in (
    "%LOCALAPPDATA%\Programs\Python\Python313\python.exe"
    "%LOCALAPPDATA%\Programs\Python\Python312\python.exe"
    "%LOCALAPPDATA%\Programs\Python\Python311\python.exe"
    "%LOCALAPPDATA%\Programs\Python\Python310\python.exe"
    "C:\Python313\python.exe"
    "C:\Python312\python.exe"
    "C:\Python311\python.exe"
    "C:\Python310\python.exe"
) do (
    if exist %%P (
        for /f "tokens=*" %%v in ('%%P --version 2^>^&1') do echo  Found: %%v at %%P
        set "PYTHON_CMD=%%~P"
        goto :python_found
    )
)

:: --- Python not found ---
echo  Python is NOT installed on this computer.
echo.
echo  When the Python installer opens:
echo    1. CHECK the box "Add python.exe to PATH" at the BOTTOM
echo    2. Click "Install Now"
echo    3. Wait for it to finish, then click "Close"
echo.
pause

set "PYTHON_URL=https://www.python.org/ftp/python/3.12.7/python-3.12.7-amd64.exe"
set "PYTHON_INSTALLER=%TEMP%\python-installer.exe"

echo.
echo  Downloading Python...
powershell -Command "& { [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12; Invoke-WebRequest -Uri '%PYTHON_URL%' -OutFile '%PYTHON_INSTALLER%' }" 2>&1

if not exist "%PYTHON_INSTALLER%" (
    echo.
    echo  *** DOWNLOAD FAILED ***
    echo  Please install Python manually from https://www.python.org/downloads/
    echo  Then run this installer again.
    echo.
    pause
    exit /b 1
)

start /wait "" "%PYTHON_INSTALLER%"
del "%PYTHON_INSTALLER%" >nul 2>&1

:: Refresh PATH
for /f "tokens=2*" %%a in ('reg query "HKCU\Environment" /v Path 2^>nul') do set "USER_PATH=%%b"
for /f "tokens=2*" %%a in ('reg query "HKLM\SYSTEM\CurrentControlSet\Control\Session Manager\Environment" /v Path 2^>nul') do set "SYS_PATH=%%b"
set "PATH=!USER_PATH!;!SYS_PATH!;%LOCALAPPDATA%\Programs\Python\Python312;%LOCALAPPDATA%\Programs\Python\Python312\Scripts"

py --version >nul 2>&1
if !ERRORLEVEL! equ 0 (
    set "PYTHON_CMD=py"
    goto :python_found
)
python --version >nul 2>&1
if !ERRORLEVEL! equ 0 (
    set "PYTHON_CMD=python"
    goto :python_found
)

echo.
echo  Python installed but not found. Please RESTART your computer
echo  and run INSTALL.bat again.
echo.
pause
exit /b 1

:python_found
echo.
echo  Python is ready. Using: !PYTHON_CMD!
echo.

:: ── Step 2: Install Packages ───────────────────────────────────────────────

echo  ============================================
echo  [Step 2/4] Installing required packages...
echo  ============================================
echo.
echo  This is the longest step. Please be patient.
echo.

echo  --- Upgrading pip ---
"!PYTHON_CMD!" -m pip install --upgrade pip 2>&1
echo.
echo  --- Installing PySide6 (may take 1-2 min) ---
"!PYTHON_CMD!" -m pip install PySide6 2>&1
echo.
echo  --- Installing openpyxl ---
"!PYTHON_CMD!" -m pip install openpyxl 2>&1
echo.
echo  --- Installing anthropic ---
"!PYTHON_CMD!" -m pip install anthropic 2>&1
echo.
echo  --- Installing python-dotenv ---
"!PYTHON_CMD!" -m pip install python-dotenv 2>&1
echo.
echo  --- Installing pywin32 ---
"!PYTHON_CMD!" -m pip install pywin32 2>&1
echo.

echo  --- Verifying ---
"!PYTHON_CMD!" -c "import PySide6; print('  PySide6: OK')" 2>&1 || echo   PySide6: FAILED
"!PYTHON_CMD!" -c "import openpyxl; print('  openpyxl: OK')" 2>&1 || echo   openpyxl: FAILED
"!PYTHON_CMD!" -c "import anthropic; print('  anthropic: OK')" 2>&1 || echo   anthropic: FAILED
"!PYTHON_CMD!" -c "import dotenv; print('  dotenv: OK')" 2>&1 || echo   dotenv: FAILED
echo.

:: ── Step 3: API Key ────────────────────────────────────────────────────────

echo  ============================================
echo  [Step 3/4] Setting up your API key...
echo  ============================================
echo.

set "ENV_FILE=!APP_DIR!\.env"

if exist "!ENV_FILE!" (
    findstr /C:"sk-ant-" "!ENV_FILE!" >nul 2>&1
    if !ERRORLEVEL! equ 0 (
        echo  API key already configured.
        goto :key_done
    )
)

echo  You should have received an API key like:
echo    sk-ant-api03-xxxxxxxxxxxxxxxxxxxx
echo.
echo  RIGHT-CLICK in this window to paste, then press Enter.
echo  Or just press Enter to skip (you can add it in the app later).
echo.

set /p "API_KEY=  Paste your API Key here: "

if "!API_KEY!"=="" (
    echo.
    echo  No key entered. You can add it in the app later.
    echo ANTHROPIC_API_KEY=> "!ENV_FILE!"
) else (
    echo ANTHROPIC_API_KEY=!API_KEY!> "!ENV_FILE!"
    echo.
    echo  API key saved.
)

:key_done
echo.

:: ── Step 4: Create Desktop Shortcut ─────────────────────────────────────────

echo  ============================================
echo  [Step 4/4] Creating desktop shortcut...
echo  ============================================
echo.

:: The launcher bat is already included in the zip — don't overwrite it.
:: Just create a shortcut pointing to it with Astra's icon.

set "LAUNCHER=!APP_DIR!\launch_astra_mailer.bat"
set "ICO=!APP_DIR!\astra_icon.ico"
set "DESKTOP=%USERPROFILE%\Desktop"
set "VBS=%TEMP%\_mkshortcut.vbs"

> "!VBS!" (
    echo Set oShell = CreateObject("WScript.Shell"^)
    echo Set oLink = oShell.CreateShortcut("!DESKTOP!\Astra Mailer.lnk"^)
    echo oLink.TargetPath = "!LAUNCHER!"
    echo oLink.WorkingDirectory = "!APP_DIR!"
    echo oLink.Description = "Astra Mailer"
    echo oLink.WindowStyle = 1
    echo oLink.IconLocation = "!ICO!"
    echo oLink.Save
)

cscript //nologo "!VBS!" >nul 2>&1
del "!VBS!" >nul 2>&1

if exist "!DESKTOP!\Astra Mailer.lnk" (
    echo  Desktop shortcut created with Astra's icon!
) else (
    echo  Could not create shortcut. Use launch_astra_mailer.bat instead.
)

echo.

:: ── Done ───────────────────────────────────────────────────────────────────

echo  ============================================
echo.
echo     INSTALLATION COMPLETE!
echo.
echo  ============================================
echo.
echo  To run: double-click "Astra Mailer" on your Desktop
echo.

set /p "LAUNCH=  Launch Astra Mailer now? (Y/N): "
if /i "!LAUNCH!"=="Y" (
    echo.
    echo  Starting Astra Mailer...
    echo.
    cd /d "!APP_DIR!"
    !PYTHON_CMD! "!APP_DIR!\astra_mailer.py"
)

echo.
pause
exit /b 0
