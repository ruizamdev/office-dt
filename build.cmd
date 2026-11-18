@echo off
rem Compila InstaladorOffice.exe con el compilador de C# incluido en .NET Framework 4.x
setlocal
cd /d "%~dp0"

set "CSC=%WINDIR%\Microsoft.NET\Framework64\v4.0.30319\csc.exe"
if not exist "%CSC%" set "CSC=%WINDIR%\Microsoft.NET\Framework\v4.0.30319\csc.exe"
if not exist "%CSC%" (
    echo No se encontro csc.exe de .NET Framework 4.x.
    echo Activa .NET Framework 4.8 en "Caracteristicas de Windows" e intenta de nuevo.
    pause
    exit /b 1
)

"%CSC%" /nologo /target:winexe /optimize+ /codepage:65001 ^
    /win32manifest:gui\app.manifest ^
    /reference:System.dll /reference:System.Drawing.dll /reference:System.Windows.Forms.dll ^
    /reference:System.Xml.dll /reference:System.Xml.Linq.dll /reference:System.Core.dll ^
    /out:InstaladorOffice.exe gui\InstaladorOffice.cs
if errorlevel 1 (
    echo.
    echo La compilacion fallo.
    pause
    exit /b 1
)

echo InstaladorOffice.exe generado en %CD%
pause
