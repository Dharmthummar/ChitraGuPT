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

echo Starting share mode for phones on this network...
echo If Windows Firewall asks, allow private network access.
python serve.py --host 0.0.0.0 --port 5055 --open
pause
