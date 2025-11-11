@echo off
set "SCRIPT=%~dp0build_installer.ps1"
REM Usa PowerShell 7 si está, si no cae a Windows PowerShell
where pwsh >nul 2>nul && (
  pwsh -NoProfile -ExecutionPolicy Bypass -File "%SCRIPT%"
) || (
  powershell -NoProfile -ExecutionPolicy Bypass -File "%SCRIPT%"
)
