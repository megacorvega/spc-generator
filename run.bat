@echo off
TITLE SPC Generator
color 0F

:: ==========================================
:: 1. CHECK ENVIRONMENT
:: ==========================================
IF NOT EXIST "venv\Scripts\python.exe" GOTO ERROR_NOVENV

:: ==========================================
:: 2. CHECK FOR DATA FILES
:: ==========================================
:: This checks if any file matches the pattern.
:: If NO file is found, errorlevel becomes 1, and we jump to the NO_DATA label.
dir /b "SPC-DATA_*.xlsx" >nul 2>&1
IF %ERRORLEVEL% NEQ 0 GOTO NO_DATA

:: ==========================================
:: 3. RUN THE TOOL
:: ==========================================
echo [INFO] Found data files. Starting generator...
venv\Scripts\python.exe -c "from spc_generator.generator import main; main()"

:: Check if Python crashed
IF %ERRORLEVEL% NEQ 0 GOTO ERROR_RUNTIME

:: Success - Keep window open
echo.
pause
exit /b

:: ==========================================
:: LABELS (The logic jumps here)
:: ==========================================

:NO_DATA
color 0E
echo.
echo [NOTICE] No data files found matching 'SPC-DATA_*.xlsx'.

:: Check if template exists using simple logic
IF EXIST "SPC-DATA_Input_Template.xlsx" GOTO SHOW_INSTRUCTIONS

:: If we are here, template is missing. Create it.
echo [ACTION] Generating blank template file...
venv\Scripts\python.exe -c "from spc_generator.template import main; main()"

:SHOW_INSTRUCTIONS
echo.
echo -------------------------------------------------------------
echo   INSTRUCTIONS:
echo   1. A file named 'SPC-DATA_Input_Template.xlsx' is ready.
echo   2. Open it and fill in your measurement data.
echo   3. Save it (keep the SPC-DATA_ prefix).
echo   4. Run this script again to generate reports.
echo -------------------------------------------------------------
echo.
pause
exit /b

:ERROR_NOVENV
color 0C
echo.
echo [ERROR] Sandbox folder (venv) not found!
echo Please run 'install.bat' first to set up the tool.
echo.
pause
exit /b

:ERROR_RUNTIME
color 0C
echo.
echo [ERROR] The Python script encountered an issue and closed.
echo.
pause
exit /b