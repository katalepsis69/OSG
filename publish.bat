@echo off
echo Publishing single-file portable executable to dist\...
dotnet publish src/BTA_OSG_DocumentTracking -c Release -r win-x64 --self-contained true -p:PublishSingleFile=true -o ./dist
echo Build complete. Executable ready in dist\
