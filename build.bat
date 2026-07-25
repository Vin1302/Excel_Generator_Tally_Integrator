@echo off
REM ---- One-click build of BankStatementTool.exe (run on Windows) ----
setlocal

echo Creating virtual environment...
python -m venv .venv
call .venv\Scripts\activate

echo Installing dependencies...
pip install --upgrade pip
pip install -r requirements.txt
pip install pyinstaller

echo Building the .exe ...
pyinstaller --noconfirm --onefile --windowed ^
  --name "BankStatementTool" ^
  --collect-all pdfplumber ^
  --collect-all pdfminer ^
  --add-data "config;config" ^
  main.py

echo.
echo Done. Your app is at:  dist\BankStatementTool.exe
echo Keep the "config" folder next to the exe so ledger_map.csv can be edited.
pause
