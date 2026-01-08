@echo off
REM Helper to build the LIS-ANALYSIS GUI executable on Windows via PyInstaller.
setlocal enabledelayedexpansion

set "REPO_ROOT=%~dp0.."
cd /d "%REPO_ROOT%"

set "VENV_DIR=.venv-pyinstaller"
if not exist "%VENV_DIR%" (
    python -m venv "%VENV_DIR%"
)
call "%VENV_DIR%\Scripts\activate.bat"
python -m pip install --upgrade pip
pip install -r requirements.txt pyinstaller

pyinstaller --clean --noconfirm LIS_ANALYSIS.spec

echo Build artifacts are available under dist\LIS-ANALYSIS\
endlocal
