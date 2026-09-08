
@echo off
echo Terminating any existing Main Hub instances...

:: Kills only Python processes running app.py
powershell -Command "Get-CimInstance Win32_Process | Where-Object { $_.CommandLine -like '*app.py*' } | Stop-Process -Force" 2>nul

echo Starting NCU Cricket Administration Suite...
python -m streamlit run app.py --server.port 8501

