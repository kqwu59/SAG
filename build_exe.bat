@echo off
REM Build script for Windows to create a single EXE.
REM Requirements: Python 3, pip.

echo Installing dependencies (pandas, openpyxl, tkinterdnd2, pyinstaller)...
py -m pip install --upgrade pip
py -m pip install pandas openpyxl tkinterdnd2 pyinstaller

echo Building EXE (onefile, windowed)...
py -m PyInstaller --noconsole --onefile --name NettoieXLSX-V13 NettoieXLSX_GUI-V13.py

echo.
echo Build finished. The EXE is located in the "dist" folder as NettoieXLSX-V13.exe
pause
