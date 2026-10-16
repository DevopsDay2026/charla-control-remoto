@echo off
chcp 65001 >nul
cd /d "%~dp0"
echo Control remoto de la charla. Abre antes la presentacion en PowerPoint.
python server.py %*
pause
