@echo off
TITLE SPC Template Generator
color 0F

:: 1. Check for install
IF NOT EXIST "venv\Scripts\python.exe" (
    color 0C
    echo [ERROR] Please run 'install.bat' first.
    pause
    exit /b
)

:: 2. Run the template function directly
venv\Scripts\python.exe -c "from spc_generator.template import main; main()"

echo.
pause