@echo off
SET ThisScriptsDirectory=%~dp0
SET PSScriptPath=%ThisScriptsDirectory%GetOutlookCalendar_v2.ps1
echo "Script Execution Started:"
Powershell.exe -NoProfile -ExecutionPolicy Unrestricted -Command "& '%PSScriptPath%'";
echo "Script Execution Completed.."
pause
