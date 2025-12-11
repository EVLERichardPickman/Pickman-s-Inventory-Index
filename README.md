# Pickman's Inventory Index — Build & Packaging Guide

This guide explains how to build the PyInstaller executable, embed the app icon, and optionally reduce size using UPX.

## Prerequisites
- Windows 10/11 (PowerShell)
- Python 3.11 (Conda env or system Python)
- Packages installed in the active environment:
  - PyQt (PyQt5)
  - requests
  - openpyxl
  - pyinstaller

Example (Conda):
```
conda create -n pickman python=3.11
conda activate pickman
conda install -c conda-forge pyqt
conda install requests openpyxl
pip install pyinstaller
```

## Optional: UPX for smaller exe
UPX compresses binaries that PyInstaller bundles. Download and add it to PATH.

- Download (Win64):
  https://github.com/upx/upx/releases/download/v5.0.2/upx-5.0.2-win64.zip
- Extract `upx.exe` to a folder, e.g. `C:\Tools\upx`
- Add to PATH for this session:
```
$env:PATH = "C:\Tools\upx;$env:PATH"
Get-Command upx
```
(If `Get-Command upx` shows the path, PyInstaller can use it.)

## Files of interest
- `PII.py` — main application script
- `PII.spec` — PyInstaller spec (onefile style), embeds icon and bundles data
- `favicon.ico` — app icon (embedded and bundled)

## Clean build commands
Run from the project folder:
```
# Clean previous outputs
Remove-Item -Recurse -Force .\build
Remove-Item -Recurse -Force .\dist

# Build using the spec (onefile exe)
pyinstaller --clean ".\PII.spec"
```

## Where the exe ends up
- Onefile exe is placed in `dist/` with the name:
  - `Pickman's Inventory Index.exe`

Launch:
```
.\dist\Pickman's Inventory Index.exe
```

## Notes about the icon
- The spec sets `icon=['favicon.ico']` to embed a PE icon for Explorer/taskbar.
- The app also bundles `favicon.ico` as data and loads it at runtime from `sys._MEIPASS` when frozen, ensuring the window icon shows even if PE extraction fails.

## Troubleshooting
- If build fails or exe is missing required libs:
  - Ensure the environment contains `PyQt5`, `requests`, `openpyxl`, and `pyinstaller`.
  - Rebuild with `--clean` to clear PyInstaller cache.
  - If UPX causes plugin issues, remove it from PATH or set `upx_exclude` in the spec for specific DLLs.
- If the icon doesn’t appear:
  - Confirm `favicon.ico` exists and is a multi-size Windows ICO (16/32/48/256px).
  - Rebuild with `--clean`.

## Git workflow (safe)
- Keep changes on a feature branch (e.g., `feature/icon-spec-fixes`).
- Push the branch to origin and open a PR to `main`.

Example:
```
git checkout -b feature/icon-spec-fixes
git add PII.py PII.spec favicon.ico README.md
git commit -m "Packaging: spec fixes + icon + README"
git push -u origin feature/icon-spec-fixes
```
Open PR via GitHub UI.
