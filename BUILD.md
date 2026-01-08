# LIS-ANALYSIS Build Guide

This document explains how to generate standalone executables for Linux and Windows using PyInstaller.

## Requirements
- Python 3.9+ installed on the build machine.
- System packages required by your GUI toolkit (for example, `python3-tk` on Linux if the GUI uses Tkinter).
- The folders `ACP/` and `Listas ATP/` must stay alongside the source code so PyInstaller can bundle them.

## Local builds
### Linux (WSL or native)
```bash
./scripts/build_linux.sh
```
- Creates/uses `.venv-pyinstaller` for isolation.
- Installs `pyinstaller` plus runtime dependencies.
- Produces artifacts under `dist/LIS-ANALYSIS/` (binary is `dist/LIS-ANALYSIS/LIS-ANALYSIS`).
- Run `dist/LIS-ANALYSIS/LIS-ANALYSIS --gui` to test the GUI.

### Windows
Run inside PowerShell or Command Prompt from the repository root:
```bat
scripts\build_windows.bat
```
- Uses `.venv-pyinstaller` (Windows format) and runs `pyinstaller --clean --noconfirm LIS_ANALYSIS.spec`.
- Result: `dist\LIS-ANALYSIS\LIS-ANALYSIS.exe`.
- Double-click or run `dist\LIS-ANALYSIS\LIS-ANALYSIS.exe --gui` for validation.

### One-file builds
To emit a single executable instead of a directory, add `--onefile` when invoking PyInstaller:
```bash
pyinstaller --clean --noconfirm --onefile LIS_ANALYSIS.spec
```
> One-file mode extracts resources to a temp folder at runtime, so startup can be slower. Use only after confirming the `--onedir` build works.

## Continuous Integration (GitHub Actions)
- Workflow: `.github/workflows/build.yml`.
- Triggers: pushes, pull requests, manual `workflow_dispatch`.
- Jobs: `build-linux` (Ubuntu) and `build-windows` (Windows Server).
- Outputs: workflow artifacts named `lis-analysis-linux` and `lis-analysis-windows` containing the `dist/LIS-ANALYSIS/` directories.
- To download: open the workflow run → **Artifacts** → download the desired zip, then test on the respective OS.

## Troubleshooting
- **Missing GUI backend**: ensure Tkinter/Qt libs are installed on the machine running the executable.
- **Data not found at runtime**: confirm new resource folders are added to `datas` inside `LIS_ANALYSIS.spec`.
- **ImportError**: add the module to `hiddenimports` in `LIS_ANALYSIS.spec` or pass `--hidden-import <module>` when running PyInstaller.
- **Antivirus flags on Windows**: sign the executable or add it to the allowlist; PyInstaller binaries are often flagged until signed.
