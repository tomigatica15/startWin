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
echo 1. Activar office 64 bits
echo 2. Activar office 32 bits
echo 3. Instalar Office Professional Plus 2019.
echo.


choice /c 123 /n /m "Seleccione una opcion: "

if errorlevel 3 goto opcion3
if errorlevel 2 goto opcion2
if errorlevel 1 goto opcion1

:opcion1

color 02
echo Ha seleccionado la opcion de 64 bits.
cd /d %ProgramFiles(x86)%\Microsoft Office\Office16
if not exist "%ProgramFiles(x86)%\Microsoft Office\Office16" (
    echo La ruta seleccionada no existe.
    goto seleccion
)
goto fin

:opcion2

color 02
echo Ha seleccionado la opcion de 32 bits.
cd /d %ProgramFiles%\Microsoft Office\Office16
if not exist "%ProgramFiles%\Microsoft Office\Office16" (
    echo La ruta seleccionada no existe.
    goto seleccion
)
goto fin

:opcion3 

color 02
echo Ha seleccionado la opcion para instalar Office Professional Plus 2019.
if exist "%ProgramFiles(x86)%" (
    set "ruta_descarga=%ProgramFiles(x86)%"
) else (
    set "ruta_descarga=%ProgramFiles%"
)
cmd -command "& {Invoke-WebRequest -Uri '' -OutFile '%ruta_descarga%'}"

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
echo 1. Office Professional Plus 2019
echo 2. Office Standard 2019
echo 3. Project Professional 2019
echo 4. Project Standard 2019
echo 5. Visio Professional 2019
echo 6. Visio Standard 2019
echo 7. Word 2019
echo 8. Excel 2019
echo 9. PowerPoint 2019
echo 10. Access 2019
echo 11. Outlook 2019
echo 12. Publisher 2019
echo 13. Skype for Business 2019
echo.

set /p "opcion=Seleccione una opcion para activar: "

:opcionA
IF "%opcion%"=="1" goto opcionA1
IF "%opcion%"=="2" goto opcionA2
IF "%opcion%"=="3" goto opcionA3
IF "%opcion%"=="4" goto opcionA4
IF "%opcion%"=="5" goto opcionA5
IF "%opcion%"=="6" goto opcionA6
IF "%opcion%"=="7" goto opcionA7
IF "%opcion%"=="8" goto opcionA8
IF "%opcion%"=="9" goto opcionA9
IF "%opcion%"=="10" goto opcionA10
IF "%opcion%"=="11" goto opcionA11
IF "%opcion%"=="12" goto opcionA12
IF "%opcion%"=="13" goto opcionA13

color 0F
echo opcion inválida. Por favor, seleccione una opcion válida.
goto seleccion2


:opcionA1
color 02
echo Ha seleccionado la opcion para activar Office Professional Plus 2019.
cmd.exe /c cscript ospp.vbs /inpkey:NMMKJ-6RK4F-KMJVX-8D9MJ-6MWKP
cls
goto :listo
:opcionA2
color 02
echo Ha seleccionado la opcion para activar Office Standard 2019.
cmd.exe /c cscript ospp.vbs /inpkey:6NWWJ-YQWMR-QKGCB-6TMB3-9D9HK 
cls
goto :listo
:opcionA3
color 02
echo Ha seleccionado la opcion para activar Project Professional 2019.
cmd.exe /c cscript ospp.vbs /inpkey:B4NPR-3FKK7-T2MBV-FRQ4W-PKD2B
cls
goto :listo
:opcionA4
color 02
echo Ha seleccionado la opcion para activar Project Standard 2019.
cmd.exe /c cscript ospp.vbs /inpkey:C4F7P-NCP8C-6CQPT-MQHV9-JXD2M
cls
goto :listo
:opcionA5
color 02
echo Ha seleccionado la opcion para activar Visio Professional 2019.
cmd.exe /c cscript ospp.vbs /inpkey:9BGNQ-K37YR-RQHF2-38RQ3-7VCBB
cls
goto :listo
:opcionA6
color 02
echo Ha seleccionado la opcion para activar Visio Standard 2019.
cmd.exe /c cscript ospp.vbs /inpkey:7TQNQ-K3YQQ-3PFH7-CCPPM-X4VQ2
cls
goto :listo
:opcionA7
color 02
echo Ha seleccionado la opcion para activar Word 2019.
cmd.exe /c cscript ospp.vbs /inpkey:PBX3G-NWMT6-Q7XBW-PYJGG-WXD33
cls
goto :listo
:opcionA8
color 02
echo Ha seleccionado la opcion para activar Excel 2019.
cmd.exe /c cscript ospp.vbs /inpkey:TMJWT-YYNMB-3BKTF-644FC-RVXBD
cls
goto :listo
:opcionA9
color 02
echo Ha seleccionado la opcion para activar PowerPoint 2019.
cmd.exe /c cscript ospp.vbs /inpkey:RRNCX-C64HY-W2MM7-MCH9G-TJHMQ
cls
goto :listo
:opcionA10
color 02
echo Ha seleccionado la opcion para activar Access 2019.
cmd.exe /c cscript ospp.vbs /inpkey:9N9PT-27V4Y-VJ2PD-YXFMF-YTFQT
cls
goto :listo
:opcionA11
color 02
echo Ha seleccionado la opcion para activar Outlook 2019.
cmd.exe /c cscript ospp.vbs /inpkey:7HD7K-N4PVK-BHBCQ-YWQRW-XW4VK
cls
goto :listo
:opcionA12
color 02
echo Ha seleccionado la opcion para activar Publisher 2019.
cmd.exe /c cscript ospp.vbs /inpkey:G2KWX-3NW6P-PY93R-JXK2T-C9Y9V
cls
goto :listo

:opcionA13
color 02
echo Ha seleccionado la opcion para activar Skype for Business 2019.
cmd.exe /c cscript ospp.vbs /inpkey:NCJ33-JHBBY-HTK98-MYCV8-HMKHJ
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