@echo off
rem ==========================================================================
rem  Instalador interactivo de Office 2019 / 2021 / 2024 con Office Deployment Tool
rem  - Paso 1: elegir version/edicion (Retail o Volume)
rem  - Paso 2: elegir los programas a instalar
rem  - Paso 3: revisar arquitectura e idioma, generar XML e instalar
rem  Requiere setup.exe (ODT) en la misma carpeta que este script.
rem ==========================================================================
setlocal EnableExtensions EnableDelayedExpansion
title Instalador de Office - Office Deployment Tool

set "ROOT=%~dp0"
set "OUTDIR=%ROOT%Personalizado"
set "ODT_URL=https://www.microsoft.com/en-us/download/details.aspx?id=49117"

rem --- Permisos de administrador (se relanza elevado si hace falta) ---
net session >nul 2>&1
if errorlevel 1 (
    echo Solicitando permisos de administrador...
    powershell -NoProfile -Command "Start-Process -FilePath '%~f0' -Verb RunAs" >nul 2>&1
    if errorlevel 1 (
        echo No se pudieron obtener permisos de administrador.
        pause
    )
    exit /b
)

rem --- Catalogo de productos: ID, nombre, licencia, anio, Outlook por defecto ---
set "PCOUNT=0"
call :addprod ProPlus2024Retail      "Office Profesional Plus 2024"      Retail 2024 1
call :addprod Professional2024Retail "Office Profesional 2024"           Retail 2024 1
call :addprod HomeBusiness2024Retail "Office Hogar y Empresas 2024"      Retail 2024 1
call :addprod Home2024Retail         "Office Hogar 2024"                 Retail 2024 0
call :addprod ProPlus2021Retail      "Office Profesional Plus 2021"      Retail 2021 1
call :addprod Professional2021Retail "Office Profesional 2021"           Retail 2021 1
call :addprod HomeBusiness2021Retail "Office Hogar y Empresas 2021"      Retail 2021 1
call :addprod HomeStudent2021Retail  "Office Hogar y Estudiantes 2021"   Retail 2021 0
call :addprod Personal2021Retail     "Office Personal 2021"              Retail 2021 0
call :addprod ProPlus2019Retail      "Office Profesional Plus 2019"      Retail 2019 1
call :addprod Professional2019Retail "Office Profesional 2019"           Retail 2019 1
call :addprod HomeBusiness2019Retail "Office Hogar y Empresas 2019"      Retail 2019 1
call :addprod HomeStudent2019Retail  "Office Hogar y Estudiantes 2019"   Retail 2019 0
call :addprod Personal2019Retail     "Office Personal 2019"              Retail 2019 0
call :addprod ProPlus2024Volume      "Office LTSC Profesional Plus 2024" Volume 2024 1
call :addprod Standard2024Volume     "Office LTSC Standard 2024"         Volume 2024 1
call :addprod ProPlus2021Volume      "Office LTSC Profesional Plus 2021" Volume 2021 1
call :addprod Standard2021Volume     "Office LTSC Standard 2021"         Volume 2021 1
call :addprod ProPlus2019Volume      "Office Profesional Plus 2019"      Volume 2019 1
call :addprod Standard2019Volume     "Office Standard 2019"              Volume 2019 1

rem --- Catalogo de programas (ID de ExcludeApp y etiqueta) ---
set "ACOUNT=0"
call :addapp Word       "Word"
call :addapp Excel      "Excel"
call :addapp PowerPoint "PowerPoint"
call :addapp Outlook    "Outlook"
call :addapp Access     "Access"
call :addapp Publisher  "Publisher"
call :addapp OneNote    "OneNote"
call :addapp Lync       "Skype Empresarial"
call :addapp Groove     "OneDrive para la Empresa - Groove"
call :addapp OneDrive   "OneDrive"

rem --- Idiomas disponibles ---
set "LCOUNT=4"
set "L[1]=es-mx"
set "L[2]=es-es"
set "L[3]=en-us"
set "L[4]=MatchOS"

rem --- Valores por defecto ---
set "ARCH=64"
set "LIDX=1"
set "OLANG=!L[1]!"

rem --- Descargar Office Deployment Tool si falta setup.exe ---
if not exist "%ROOT%setup.exe" (
    cls
    call :header
    call :get_odt
    if errorlevel 1 pause
)

rem ==========================================================================
:MENU_VERSION
cls
call :header
echo   Paso 1 de 3: elige la version de Office
set "last="
for /l %%i in (1,1,!PCOUNT!) do (
    if not "!P_LIC[%%i]!"=="!last!" (
        echo.
        if "!P_LIC[%%i]!"=="Retail" (
            echo   --- RETAIL: clave de producto individual ---
        ) else (
            echo   --- VOLUMEN: KMS / MAK ---
        )
        set "last=!P_LIC[%%i]!"
    )
    set "num=  %%i"
    set "num=!num:~-2!"
    set "nm=!P_NAME[%%i]!                                    "
    set "nm=!nm:~0,35!"
    set "note="
    if !P_YEAR[%%i]! lss 2024 set "note= *"
    echo    [!num!] !nm! !P_ID[%%i]!!note!
)
echo.
echo   * Office 2019 y 2021 estan fuera de soporte desde el 14/10/2025.
echo.
echo   [Q] Salir
echo.
set "sel="
set /p "sel=  Opcion: "
if not defined sel goto MENU_VERSION
if /i "!sel!"=="Q" goto END
call :parsenum "!sel!" !PCOUNT! || goto MENU_VERSION
set "P=!NUM!"

set "PID=!P_ID[%P%]!"
set "PNAME=!P_NAME[%P%]!"
set "PYEAR=!P_YEAR[%P%]!"
set "PCHAN="
if "!P_LIC[%P%]!"=="Volume" set "PCHAN=PerpetualVL!PYEAR!"

rem Programas por defecto, igual que los XML de Retail/ y Volume/
for /l %%i in (1,1,!ACOUNT!) do set "A_ON[%%i]=0"
call :setapp Word 1
call :setapp Excel 1
call :setapp PowerPoint 1
call :setapp OneDrive 1
call :setapp Outlook !P_OUT[%P%]!

rem ==========================================================================
:MENU_APPS
cls
call :header
echo   Paso 2 de 3: programas a instalar
echo   Version: !PNAME!  [!PID!]
echo.
for /l %%i in (1,1,!ACOUNT!) do (
    set "mark= "
    if "!A_ON[%%i]!"=="1" set "mark=X"
    set "num=  %%i"
    set "num=!num:~-2!"
    set "nm=!A_LBL[%%i]!                                    "
    set "nm=!nm:~0,34!"
    set "note="
    if "!A_ID[%%i]!"=="Publisher" if "!PYEAR!"=="2024" set "note=no incluido en 2024"
    echo    !num!. [!mark!] !nm!!note!
)
echo.
echo   Nota: solo se instalan los programas que incluye la edicion elegida.
echo.
echo   Escribe uno o varios numeros separados por espacio para marcar/desmarcar.
echo   [T] Todos   [N] Ninguno   [C] Continuar   [V] Volver   [Q] Salir
echo.
set "in="
set /p "in=  Opcion: "
if not defined in goto MENU_APPS
if /i "!in!"=="Q" goto END
if /i "!in!"=="V" goto MENU_VERSION
if /i "!in!"=="T" (
    for /l %%i in (1,1,!ACOUNT!) do set "A_ON[%%i]=1"
    goto MENU_APPS
)
if /i "!in!"=="N" (
    for /l %%i in (1,1,!ACOUNT!) do set "A_ON[%%i]=0"
    goto MENU_APPS
)
if /i "!in!"=="C" goto CHECK_APPS
rem Solo se aceptan digitos y espacios
for /f "delims=0123456789 " %%a in ("!in!") do goto MENU_APPS
for %%n in (!in!) do call :toggle "%%n"
goto MENU_APPS

:CHECK_APPS
set "APPLIST="
for /l %%i in (1,1,!ACOUNT!) do (
    if "!A_ON[%%i]!"=="1" (
        if defined APPLIST (set "APPLIST=!APPLIST!, !A_ID[%%i]!") else set "APPLIST=!A_ID[%%i]!"
    )
)
if not defined APPLIST (
    echo.
    echo   Selecciona al menos un programa.
    pause
    goto MENU_APPS
)

rem ==========================================================================
:MENU_OPTS
cls
call :header
echo   Paso 3 de 3: resumen
echo.
echo   Version      : !PNAME!
echo   Producto ID  : !PID!
if defined PCHAN (
    echo   Canal        : !PCHAN!
) else (
    echo   Canal        : predeterminado para Retail
)
echo   Arquitectura : !ARCH! bits
echo   Idioma       : !OLANG!
echo   Programas    : !APPLIST!
echo.
echo   [I] Instalar ahora
echo   [G] Solo guardar el XML
echo   [A] Cambiar arquitectura 64/32 bits
echo   [L] Cambiar idioma
echo   [P] Volver a programas
echo   [V] Volver a version
echo   [Q] Salir
echo.
choice /c IGALPVQ /n /m "  Opcion: "
set "k=!errorlevel!"
if "!k!"=="1" goto DO_INSTALL
if "!k!"=="2" goto DO_SAVE
if "!k!"=="3" (
    if "!ARCH!"=="64" (set "ARCH=32") else set "ARCH=64"
    goto MENU_OPTS
)
if "!k!"=="4" (
    set /a "LIDX=LIDX %% LCOUNT + 1"
    for %%l in (!LIDX!) do set "OLANG=!L[%%l]!"
    goto MENU_OPTS
)
if "!k!"=="5" goto MENU_APPS
if "!k!"=="6" goto MENU_VERSION
goto END

:DO_SAVE
call :write_xml || (pause & goto MENU_OPTS)
echo.
echo   XML guardado en:
echo   !XML!
echo.
echo   Para instalarlo despues:  setup.exe /configure "!XML!"
echo.
pause
goto MENU_OPTS

:DO_INSTALL
if not exist "%ROOT%setup.exe" (
    echo.
    echo   No se encontro setup.exe en:
    echo   %ROOT%
    echo.
    echo   [D] Descargar Office Deployment Tool automaticamente
    echo   [P] Abrir la pagina de descarga manual
    echo   [V] Volver
    echo.
    choice /c DPV /n /m "  Opcion: "
    if !errorlevel!==1 (
        call :get_odt
        pause
    )
    if !errorlevel!==2 start "" "%ODT_URL%"
    goto MENU_OPTS
)
echo.
echo   Se instalara !PNAME! - !ARCH! bits, !OLANG!
echo   Programas: !APPLIST!
echo   Office se descarga desde la CDN de Microsoft; puede tardar varios minutos.
echo.
choice /c SN /n /m "  Confirmar instalacion? [S/N]: "
if errorlevel 2 goto MENU_OPTS
call :write_xml || (pause & goto MENU_OPTS)
echo.
echo   Configuracion: !XML!
echo   Ejecutando setup.exe /configure ...
pushd "%ROOT%"
setup.exe /configure "!XML!"
set "RC=!errorlevel!"
popd
echo.
if "!RC!"=="0" (
    echo   Instalacion terminada.
) else (
    echo   setup.exe termino con codigo !RC!.
    echo   Revisa los registros de ODT en !TEMP!
)
echo.
pause
goto END

rem ==========================================================================
rem  Subrutinas
rem ==========================================================================

:header
echo ===================================================================
echo    Instalador de Office 2019 - 2024  ^|  Office Deployment Tool
echo ===================================================================
echo.
exit /b

rem addprod ID "Nombre" Retail/Volume Anio OutlookPorDefecto
:addprod
set /a PCOUNT+=1
set "P_ID[%PCOUNT%]=%~1"
set "P_NAME[%PCOUNT%]=%~2"
set "P_LIC[%PCOUNT%]=%~3"
set "P_YEAR[%PCOUNT%]=%~4"
set "P_OUT[%PCOUNT%]=%~5"
exit /b

rem addapp ID "Etiqueta"
:addapp
set /a ACOUNT+=1
set "A_ID[%ACOUNT%]=%~1"
set "A_LBL[%ACOUNT%]=%~2"
exit /b

rem setapp ID 0/1
:setapp
for /l %%i in (1,1,%ACOUNT%) do if /i "!A_ID[%%i]!"=="%~1" set "A_ON[%%i]=%~2"
exit /b

rem get_odt  ->  descarga el ODT de Microsoft, verifica la firma y extrae setup.exe
rem El codigo de PowerShell no debe contener comillas dobles, porcentajes ni signos de exclamacion.
:get_odt
echo   Descargando Office Deployment Tool desde Microsoft...
set "ODT_DEST=%ROOT%setup.exe"
set "PS=$ErrorActionPreference='Stop'; $ProgressPreference='SilentlyContinue';"
set "PS=%PS% [Net.ServicePointManager]::SecurityProtocol=[Net.ServicePointManager]::SecurityProtocol -bor [Net.SecurityProtocolType]::Tls12;"
set "PS=%PS% $ua='Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/130.0.0.0 Safari/537.36';"
set "PS=%PS% $re='https://download\.microsoft\.com/download/[\w\-/\.]*officedeploymenttool[\w\-\.]*\.exe'; $u=$null;"
set "PS=%PS% foreach($p in 'details','confirmation'){ try{ $c=(Invoke-WebRequest -UseBasicParsing -UserAgent $ua -Headers @{'Accept-Language'='en-US,en;q=0.9'} -Uri ('https://www.microsoft.com/en-us/download/'+$p+'.aspx?id=49117')).Content;"
set "PS=%PS% $m=[regex]::Match($c,$re,'IgnoreCase'); if($m.Success){$u=$m.Value; break} }catch{} };"
set "PS=%PS% if(-not $u){throw 'No se encontro el enlace de descarga en la pagina de Microsoft.'};"
set "PS=%PS% function Test-Ms($x){ $s=Get-AuthenticodeSignature -FilePath $x; if($s.Status -ne 'Valid' -or $s.SignerCertificate.Subject -notmatch 'O=Microsoft Corporation'){throw ('Firma digital no valida: '+(Split-Path $x -Leaf))} };"
set "PS=%PS% $t=Join-Path $env:TEMP ('odt-'+[guid]::NewGuid().ToString('N')); $null=New-Item -ItemType Directory -Path $t;"
set "PS=%PS% try{ $f=Join-Path $t (Split-Path $u -Leaf); Write-Host ('  '+$u); Invoke-WebRequest -UseBasicParsing -UserAgent $ua -Uri $u -OutFile $f;"
set "PS=%PS% Write-Host '  Verificando firma digital...'; Test-Ms $f; $e=Join-Path $t 'x';"
set "PS=%PS% Write-Host '  Extrayendo setup.exe...'; $pr=Start-Process -FilePath $f -ArgumentList ('/quiet /extract:'+[char]34+$e+[char]34) -Wait -PassThru;"
set "PS=%PS% if($pr.ExitCode -ne 0){throw ('La extraccion termino con codigo '+$pr.ExitCode)};"
set "PS=%PS% $s2=Join-Path $e 'setup.exe'; if(-not (Test-Path $s2)){throw 'La extraccion no genero setup.exe.'};"
set "PS=%PS% Test-Ms $s2; Copy-Item -LiteralPath $s2 -Destination $env:ODT_DEST -Force }"
set "PS=%PS% finally{ Remove-Item -LiteralPath $t -Recurse -Force -ErrorAction SilentlyContinue }"
powershell -NoProfile -ExecutionPolicy Bypass -Command "%PS%"
if errorlevel 1 (
    echo.
    echo   No se pudo descargar Office Deployment Tool.
    echo   Puedes descargarlo manualmente desde: %ODT_URL%
    exit /b 1
)
if not exist "%ROOT%setup.exe" exit /b 1
echo   Office Deployment Tool listo: %ROOT%setup.exe
exit /b 0

rem toggle "numero"  ->  marca/desmarca el programa
:toggle
call :parsenum %1 %ACOUNT% || exit /b 0
if "!A_ON[%NUM%]!"=="1" (set "A_ON[%NUM%]=0") else set "A_ON[%NUM%]=1"
exit /b 0

rem parsenum "texto" max  ->  NUM con el valor; errorlevel 1 si no es 1..max
:parsenum
set "NUM=%~1"
set "NUM=!NUM: =!"
if not defined NUM exit /b 1
for /f "delims=0123456789" %%a in ("!NUM!") do exit /b 1
for /f "tokens=* delims=0" %%a in ("!NUM!") do set "NUM=%%a"
if not "!NUM:~3!"=="" exit /b 1
if !NUM! lss 1 exit /b 1
if !NUM! gtr %~2 exit /b 1
exit /b 0

rem write_xml  ->  genera la configuracion en %OUTDIR% y deja la ruta en XML
:write_xml
if not exist "%OUTDIR%" mkdir "%OUTDIR%" 2>nul
set "XML=%OUTDIR%\!PID!-x!ARCH!-!OLANG!.xml"
set "CHATTR="
if defined PCHAN set "CHATTR= Channel="!PCHAN!""
> "!XML!" (
    echo ^<Configuration^>
    echo   ^<Add OfficeClientEdition="!ARCH!"!CHATTR!^>
    echo     ^<Product ID="!PID!"^>
    echo       ^<Language ID="!OLANG!" /^>
    for /l %%i in (1,1,!ACOUNT!) do (
        if "!A_ON[%%i]!"=="0" echo       ^<ExcludeApp ID="!A_ID[%%i]!" /^>
    )
    echo     ^</Product^>
    echo   ^</Add^>
    echo   ^<Updates Enabled="TRUE" /^>
    echo   ^<Display Level="Full" AcceptEULA="TRUE" /^>
    echo ^</Configuration^>
)
if errorlevel 1 (
    echo   No se pudo escribir !XML!
    exit /b 1
)
exit /b 0

:END
endlocal
exit /b 0
