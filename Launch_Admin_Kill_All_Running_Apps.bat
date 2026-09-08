@echo off
echo Stopping existing NCU Cricket Administration Apps...

:: Forcefully close any running Python processes
taskkill /F /IM python.exe >nul 2>&1

:: Adding a brief timeout gives Windows a second to clear the processes
timeout /t 2 /nobreak >nul

echo All Apps stopped
