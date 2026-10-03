@echo off
rem Launches the shipped build. The Debug build has its own separate
rem settings and will look like demo mode - always launch dist.
if not exist "%~dp0dist\BTA_OSG_DocumentTracking.exe" (
    echo Run publish.bat first - dist\BTA_OSG_DocumentTracking.exe is missing.
    pause
    exit /b 1
)
start "" "%~dp0dist\BTA_OSG_DocumentTracking.exe"
