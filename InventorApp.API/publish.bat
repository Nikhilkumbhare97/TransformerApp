@echo off
echo Building and publishing InventorApp.API...

REM Clean previous builds
if exist "bin\Release\net8.0-windows\publish" rmdir /s /q "bin\Release\net8.0-windows\publish"
if exist "bin\Release\net8.0-windows\win-x64" rmdir /s /q "bin\Release\net8.0-windows\win-x64"

REM Publish as self-contained executable for Windows x64
dotnet publish -c Release -r win-x64 --self-contained true -p:PublishSingleFile=true -p:IncludeNativeLibrariesForSelfExtract=true -o "bin\Release\net8.0-windows\win-x64"

if %ERRORLEVEL% EQU 0 (
    echo.
    echo ✅ Build successful!
    echo 📁 Executable location: bin\Release\net8.0-windows\win-x64\InventorApp.API.exe
    echo.
    echo To run the application:
    echo   cd bin\Release\net8.0-windows\win-x64
    echo   InventorApp.API.exe
    echo.
) else (
    echo.
    echo ❌ Build failed!
    echo.
)

pause
