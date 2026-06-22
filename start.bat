@echo off
setlocal
cd /d "%~dp0"

set "VENV_PYTHON=.venv\Scripts\python.exe"

if not exist "%VENV_PYTHON%" (
  where python >nul 2>nul
  if errorlevel 1 (
    echo Python was not found in PATH.
    exit /b 1
  )

  echo Creating local Python environment...
  python -m venv .venv
  if errorlevel 1 (
    echo Failed to create .venv.
    exit /b 1
  )
)

"%VENV_PYTHON%" -c "import openpyxl" >nul 2>nul
if errorlevel 1 (
  echo Installing required Python packages...
  "%VENV_PYTHON%" -m pip install -q --upgrade pip
  "%VENV_PYTHON%" -m pip install -r requirements.txt
  if errorlevel 1 (
    echo Failed to install Python packages.
    exit /b 1
  )
)

"%VENV_PYTHON%" "src\pipeline.py" %*
exit /b %ERRORLEVEL%
