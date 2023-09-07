@echo off
color 0F

echo Seleccione la ruta de su office:

echo Ejecutando en modo administrador...
echo.

:: Comprobar si se está ejecutando como administrador
net session >nul 2>&1
if %errorlevel% neq 0 (
    color 04
    echo No se esta ejecutando como administrador.
    echo Por favor, abre el archivo en modo administrador.
    pause
    exit
)

:menu 
echo 1. Windows activador y office
echo 2. Programas
echo.

choice /c 123 /n /m "Seleccione una opcion: "
if errorlevel 2 goto menu2
if errorlevel 1 goto menu1

:menu2
start powershell -Command "iwr -useb https://christitus.com/win | iex"'
goto fin

:menu1 
echo 1. Activar Windows
echo 2. Office
echo. 

choice /c 123 /n /m "Seleccione una opcion: "
if errorlevel 2 goto seleccion
if errorlevel 1 goto actWin

:actWin
cmd slmgr /ipk VK7JG-NPHTM-C97JM-9MPGT-3V66T
cmd slmgr /skms kms.digiboy.ir
cmd kms.msguides.com
cmd slmgr /ato
echo Windows activado.
goto fin

:seleccion
echo 1. Activar office
echo 2. Instalar Office.
echo.


choice /c 123 /n /m "Seleccione una opcion: "

if errorlevel 2 goto opcion3
if errorlevel 1 goto opcion1

:opcion1

color 02
echo Ha seleccionado la opcion para activar office.
goto fin

:opcion3 

color 02
echo Ha seleccionado la opcion para instalar Office Professional Plus 2019.
if exist "%ProgramFiles%" (
    set "ruta_descarga=%ProgramFiles%"
    pushd "%~dp0"
    bin.exe /configure "configuration/configuration-x64.xml"
) else (
    set "ruta_descarga=%ProgramFiles(x86)%"
    pushd "%~dp0"
    bin.exe /configure "configuration/configuration-x32.xml"
)

if %errorlevel% equ 0 (
    echo Descarga completada correctamente.
) else (
    echo Se produjo un error durante la descarga.
)

goto :seleccion

:fin

cmd.exe /c for /f %x in ('dir /b ..\root\Licenses16\ProPlus2019VL*.xrm-ms') do cscript ospp.vbs /inslic:"..\root\Licenses16\%x"

cmd.exe /c cscript ospp.vbs /setprt:1688
cmd.exe /c cscript ospp.vbs /unpkey:6MWKP >nul
cls
echo Seleccione el paquete de office:
color 0F
:seleccion2
echo 1. Office Professional
echo.

set /p "opcion=Seleccione una opcion para activar: "

:opcionA
IF "%opcion%"=="1" goto opcionA1

color 0F
echo opcion inválida. Por favor, seleccione una opcion válida.
goto seleccion2


:opcionA1
color 02
echo Ha seleccionado la opcion para activar Office Professional Plus 2019.
if exist "C:\Program Files\Microsoft Office\Office16\ospp.vbs" cd /d "C:\Program Files\Microsoft Office\Office16"
if exist "C:\Program Files (x86)\Microsoft Office\Office16\ospp.vbs" cd /d "C:\Program Files (x86)\Microsoft Office\Office16"
for /f %%x in ('dir /b ..\root\Licenses16\ProPlus2021VL_KMS*.xrm-ms') do cscript ospp.vbs /inslic:"..\root\Licenses16\%%x"
cscript ospp.vbs /inpkey:FXYTK-NJJ8C-GB6DW-3DYQT-6F7TH
cscript ospp.vbs /sethst:kms.msgang.com
cscript ospp.vbs /act
cls
goto :listo

:listo

cmd.exe /c cscript ospp.vbs /sethst:e8.us.to
cmd.exe /c cscript ospp.vbs /act
color 02
cls
echo Reinicie el ordenador para aplicar los cambios
echo ===================================
echo ==         by tomigatica         ==
echo ===================================
pause