@echo off
rem Builds the app and produces installer\FileConverterSetup.msi
rem Requires the WiX v6 tool:  dotnet tool install --global wix --version 6.0.2
rem (Do NOT upgrade to WiX v7 - it requires accepting the Open Source Maintenance Fee EULA.)

setlocal
cd /d "%~dp0"

echo [1/2] Publishing FileConverter.exe...
dotnet publish -c Release -r win-x64 --self-contained true ^
  -p:PublishSingleFile=true ^
  -p:IncludeNativeLibrariesForSelfExtract=true ^
  -p:EnableCompressionInSingleFile=true ^
  -o publish
if errorlevel 1 ( echo Publish FAILED. & exit /b 1 )

echo [2/2] Building installer MSI...
wix build installer\setup.wxs -o installer\FileConverterSetup.msi
if errorlevel 1 ( echo MSI build FAILED. & exit /b 1 )

echo.
echo Done: installer\FileConverterSetup.msi
echo Install silently:   msiexec /i installer\FileConverterSetup.msi /qn
echo Install with UI:    double-click the MSI
pause
