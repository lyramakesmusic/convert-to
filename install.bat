@echo off
setlocal EnableDelayedExpansion

:: Create program directory
set "INSTALL_DIR=%LOCALAPPDATA%\MediaConverter"
if not exist "%INSTALL_DIR%" mkdir "%INSTALL_DIR%"

:: Copy files to install directory
echo Installing files to %INSTALL_DIR%...
copy /Y "%~dp0convert.ps1" "%INSTALL_DIR%\convert.ps1"
copy /Y "%~dp0SetupContextMenu.ps1" "%INSTALL_DIR%\SetupContextMenu.ps1"

:: Update the script file path in SetupContextMenu.ps1
echo Configuring scripts...
powershell -Command ^
    "$content = Get-Content '%INSTALL_DIR%\SetupContextMenu.ps1'; ^
     $content = $content -replace '\$PSScriptRoot\\convert\.ps1', '%INSTALL_DIR:\=\\%\\convert.ps1' -replace '\$PSScriptRoot\\silent_wrapper\.vbs', '%INSTALL_DIR:\=\\%\\silent_wrapper.vbs'; ^
     $content | Set-Content '%INSTALL_DIR%\SetupContextMenu.ps1'"

:: Set execution policy and run setup script
echo Setting PowerShell execution policy...
powershell -Command "Set-ExecutionPolicy Bypass -Scope CurrentUser -Force"

echo Installing context menu entries...
powershell -NoProfile -ExecutionPolicy Bypass -File "%INSTALL_DIR%\SetupContextMenu.ps1"

echo Installation complete!
echo Files installed to: %INSTALL_DIR%
timeout /t 5