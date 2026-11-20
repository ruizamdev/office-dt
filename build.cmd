@echo off
rem Compila InstaladorOffice.exe con el SDK de .NET (https://dotnet.microsoft.com/download). El exe queda en dist\.
setlocal
cd /d "%~dp0"
dotnet build src\InstaladorOffice\InstaladorOffice.csproj -c Release -nologo -o dist
if errorlevel 1 (
    echo La compilacion fallo.
    pause
    exit /b 1
)
del /q dist\*.config 2>nul
echo Listo: %CD%\dist\InstaladorOffice.exe
pause
