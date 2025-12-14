@echo off
set "psScript=%~dp0CodeSiging_Files_Folder.ps1"

:: Check if running as admin
net session >nul 2>&1
if %errorlevel% neq 0 (
    echo Requesting admin privileges...
    powershell -Command "Start-Process 'cmd.exe' -ArgumentList '/c ""%~fs0""' -Verb RunAs"
    exit /b
)

:: Run the PowerShell script
powershell -executionpolicy bypass -File "%psScript%"
