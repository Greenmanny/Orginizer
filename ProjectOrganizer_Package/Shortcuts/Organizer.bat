@echo off
set "SCRIPT=%USERPROFILE%\PROJECTS\_Organizer\Organizer.ps1"
where pwsh >nul 2>&1 && (
  pwsh -STA -ExecutionPolicy Bypass -File "%SCRIPT%"
) || (
  powershell -STA -ExecutionPolicy Bypass -File "%SCRIPT%"
)
