@echo off
REM Genera dist\PDFCreator.exe (Windows)
python -m pip install -r requirements.txt pyinstaller
python -m PyInstaller --noconfirm --onefile --windowed --name PDFCreator --hidden-import win32timezone run.py
echo.
echo Ejecutable generado en dist\PDFCreator.exe
pause
