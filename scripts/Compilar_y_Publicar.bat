@echo off
setlocal
cd /d "%~dp0\.."
powershell -NoProfile -ExecutionPolicy Bypass -File ".\scripts\build_installer.ps1" -Publish actions
echo.
echo Listo. Se empujo el tag y GitHub Actions esta publicando la release.
echo Revisa la pestaña "Actions" del repo o "Releases".
pause
