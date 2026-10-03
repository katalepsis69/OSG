@echo off
REM Publish settings (single-file, self-contained, compression, en-only
REM satellites) live in the vbproj under the Release configuration.
REM Cleaning first prevents stale files from old publishes piling up in dist
REM and stops runtime leftovers (e.g. bta_osg_db.xml) from being swept in.
echo Cleaning previous build output...
if exist dist rmdir /s /q dist
if exist src\BTA_OSG_DocumentTracking\bin\Release rmdir /s /q src\BTA_OSG_DocumentTracking\bin\Release
echo Publishing single-file portable executable to dist\...
dotnet publish src/BTA_OSG_DocumentTracking -c Release -o ./dist
if errorlevel 1 (
    echo PUBLISH FAILED.
    exit /b 1
)
echo Build complete. Ship dist\BTA_OSG_DocumentTracking.exe together with the dist\Resources folder.
