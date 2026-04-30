@echo off
setlocal
cd /d "%~dp0"

if not exist ".venv\Scripts\python.exe" (
  echo Creating local Python environment...
  py -3 -m venv .venv
  if errorlevel 1 python -m venv .venv
)

call ".venv\Scripts\activate.bat"

python -c "import flask, openpyxl, waitress" >nul 2>nul
if errorlevel 1 (
  echo Installing app requirements...
  python -m pip install -r requirements.txt
)

echo Starting ChitraGuPT...
python serve.py --host 127.0.0.1 --port 5055 --open
pause
