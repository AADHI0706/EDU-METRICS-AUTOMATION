@echo off
REM Builds AcademicResultPortal.exe  (needs Python 3.9+ installed, run once)
cd /d "%~dp0"
python -m venv venv
call venv\Scripts\activate
pip install -r requirements.txt
pyinstaller --noconfirm --onefile --name AcademicResultPortal ^
  --add-data "templates;templates" ^
  --collect-submodules flask_sqlalchemy ^
  --collect-submodules sqlalchemy.dialects.sqlite ^
  --hidden-import openpyxl ^
  app.py
echo.
echo Done. Your program is: dist\AcademicResultPortal.exe
pause
