
@echo off
echo =============================
echo Gasi sve Python procese...
taskkill /F /IM python.exe
timeout /t 2 > nul

echo =============================
echo Pokrece Flask aplikaciju...
cd /d "C:\Users\Napravi Dokument\Desktop\PDF Generator radnih svezaka\Generator"
python app.py

pause
