#!/usr/bin/env bash
set -euo pipefail

# Helper to build the LIS-ANALYSIS GUI executable on Linux via PyInstaller.
REPO_ROOT="$(cd "$(dirname "${BASH_SOURCE[0]}")/.." && pwd)"
cd "$REPO_ROOT"

VENV_DIR=${VENV_DIR:-.venv-pyinstaller}
PYTHON_BIN=${PYTHON_BIN:-python3}

if [[ ! -d "$VENV_DIR" ]]; then
    "$PYTHON_BIN" -m venv "$VENV_DIR"
fi

# shellcheck disable=SC1090
source "$VENV_DIR/bin/activate"
python -m pip install --upgrade pip
pip install -r requirements.txt pyinstaller

pyinstaller --clean --noconfirm LIS_ANALYSIS.spec

echo "Build artifacts are available under dist/LIS-ANALYSIS/"
