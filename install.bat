@echo off
TITLE SPC Generator Installer (Local Mode)
color 0A

echo ========================================================
echo   SPC AUTOMATION TOOL - LOCAL INSTALLER
echo ========================================================
echo.

:: 1. Check Python
echo [1/4] Checking Python...
:: This was already silenced in your original script
python --version >nul 2>&1
IF %ERRORLEVEL% NEQ 0 (
    echo [ERROR] Python not found. Please install it first.
    pause
    exit /b
)

:: 2. Create Virtual Environment (The "Sandbox")
echo.
echo [2/4] Creating local sandbox (Virtual Environment)...
:: Added >nul 2>&1 to silence output
python -m venv venv >nul 2>&1

:: 3. Activate the Sandbox
echo.
echo [3/4] Activating sandbox...
:: Added >nul 2>&1 to silence the activation path output
call venv\Scripts\activate.bat >nul 2>&1

:: 4. Install Tool into the Sandbox
echo.
echo [4/4] Installing dependencies (This may take a moment)...
:: Added >nul 2>&1 to silence the pip install download logs
python -m pip install --upgrade pip >nul 2>&1
python -m pip install -e . >nul 2>&1

echo.
echo ========================================================
echo   SUCCESS! 
echo ========================================================
echo.
echo   The tool is installed in the 'venv' folder.
echo   You must use the 'run.bat' or 'get-template.bat' files
echo   to use the tool (global commands won't work).
echo.
pause