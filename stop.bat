@echo off
echo ==========================
echo Zaustavljanje svih Python procesa...
taskkill /F /IM python.exe >nul 2>&1
taskkill /F /IM pythonw.exe >nul 2>&1
echo Gotovo.
echo ==========================
pause
