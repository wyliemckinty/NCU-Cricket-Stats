@echo off
echo Terminating any existing Secretary Portal instances...

:: Kills only Python processes running secretary_app.py
powershell -Command "Get-CimInstance Win32_Process | Where-Object { $_.CommandLine -like '*secretary_app.py*' } | Stop-Process -Force" 2>nul

echo Starting NCU Cricket Secretary Portal...
python -m streamlit run secretary_app.py --server.port 8503
