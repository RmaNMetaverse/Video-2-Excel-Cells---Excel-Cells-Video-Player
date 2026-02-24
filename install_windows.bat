@echo off
REM Installer for Windows: creates venv and installs dependencies for excel_puppet.py
SETLOCAL

REM IMPORTANT: This installer and Python must be run with administrative privileges.
REM Open an elevated Command Prompt (right-click -> Run as administrator) before running.
NET SESSION >nul 2>&1
IF %ERRORLEVEL% NEQ 0 (
  echo This installer must be run with administrative privileges.
  echo Right-click this script and choose "Run as administrator", or open an elevated Command Prompt.
  pause
  exit /b 1
)

REM Prefer python, fall back to py launcher
where python >nul 2>&1 || where py >nul 2>&1
IF ERRORLEVEL 1 (
  echo Python not found in PATH. Please install Python 3.8+ and re-run this script.
  exit /b 1
)

IF NOT EXIST venv (
  echo Creating virtual environment in venv\
  python -m venv venv 2>nul || py -3 -m venv venv
) ELSE (
  echo Using existing venv
)

call venv\Scripts\activate
python -m pip install --upgrade pip
pip install -r requirements.txt

REM Run pywin32 postinstall if present
IF EXIST "%VIRTUAL_ENV%\Scripts\pywin32_postinstall.py" (
  echo Running pywin32 post-install step...
  python "%VIRTUAL_ENV%\Scripts\pywin32_postinstall.py" -install
)

echo
echo Installation complete. Activate the venv with:
echo    call venv\Scripts\activate
echo Then run: python excel_puppet.py
ENDLOCAL