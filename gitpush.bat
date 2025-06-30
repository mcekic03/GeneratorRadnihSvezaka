@echo off
cd /d "C:\Users\Napravi Dokument\Desktop\PDF Generator radnih svezaka\Generator"
echo Unesi poruku za commit:
set /p msg=
git add .
git commit -m "%msg%"
git push origin main
pause