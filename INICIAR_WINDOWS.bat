@echo off
cd /d "%~dp0"
echo Iniciando vista previa en http://localhost:8080/VISTA_PREVIA.html
start "" "http://localhost:8080/VISTA_PREVIA.html"
py -3 -m http.server 8080
if errorlevel 1 python -m http.server 8080
pause
