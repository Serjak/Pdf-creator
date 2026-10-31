@echo off
REM Para que el .exe funcione en Windows 7, compilalo con Python 3.8 (32 bits).
REM Genera dist\PDFCreator.exe (Windows)
python -m pip install -r requirements.txt "pyinstaller==5.13.2"
python -m PyInstaller --noconfirm --onefile --windowed --name PDFCreator --hidden-import win32timezone run.py
echo.
echo Ejecutable generado en dist\PDFCreator.exe
pause
