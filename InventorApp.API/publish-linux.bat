@echo off
echo Building and publishing InventorApp.API for Linux...

REM Clean previous builds
if exist "bin\Release\net8.0-windows\publish" rmdir /s /q "bin\Release\net8.0-windows\publish"
if exist "bin\Release\net8.0-windows\linux-x64" rmdir /s /q "bin\Release\net8.0-windows\linux-x64"

REM Publish as self-contained executable for Linux x64
dotnet publish -c Release -r linux-x64 --self-contained true -p:PublishSingleFile=true -p:IncludeNativeLibrariesForSelfExtract=true -o "bin\Release\net8.0-windows\linux-x64"

if %ERRORLEVEL% EQU 0 (
    echo.
    echo ✅ Build successful!
    echo 📁 Executable location: bin\Release\net8.0-windows\linux-x64\InventorApp.API
    echo.
    echo To run the application on Linux:
    echo   chmod +x bin\Release\net8.0-windows\linux-x64\InventorApp.API
    echo   ./bin\Release\net8.0-windows\linux-x64\InventorApp.API
    echo.
) else (
    echo.
    echo ❌ Build failed!
    echo.
)

pause
