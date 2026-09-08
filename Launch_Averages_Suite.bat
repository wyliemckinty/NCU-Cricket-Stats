@echo off
echo Terminating any existing Stats App instances...

:: Kills only Python processes running stats_app.py
powershell -Command "Get-CimInstance Win32_Process | Where-Object { $_.CommandLine -like '*stats_app.py*' } | Stop-Process -Force" 2>nul

echo Starting NCU Cricket Stats App...
python -m streamlit run stats_app.py --server.port 8502