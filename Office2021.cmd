@echo off
title Activate Microsoft Office 2021 (ALL versions) for FREE - MSGuides.com
cls

echo =====================================================================================
echo # Project: Activating Microsoft software products for FREE without additional software
echo =====================================================================================
echo.
echo # Supported products:
echo - Microsoft Office Standard 2021
echo - Microsoft Office Professional Plus 2021
echo.

:: Change directory to the Office installation folder if ospp.vbs exists
if exist "%ProgramFiles%\Microsoft Office\Office16\ospp.vbs" (
    cd /d "%ProgramFiles%\Microsoft Office\Office16"
) else if exist "%ProgramFiles(x86)%\Microsoft Office\Office16\ospp.vbs" (
    cd /d "%ProgramFiles(x86)%\Microsoft Office\Office16"
)

:: Install the KMS licenses
for /f %%x in ('dir /b ..\root\Licenses16\ProPlus2021VL_KMS*.xrm-ms') do (
    cscript ospp.vbs /inslic:"..\root\Licenses16\%%x" >nul
)

echo.
echo =====================================================================================
echo Activating your product...

:: Clear existing KMS settings and set the port
cscript //nologo slmgr.vbs /ckms >nul
cscript //nologo ospp.vbs /setprt:1688 >nul
cscript //nologo ospp.vbs /unpkey:6F7TH >nul

set i=1

:: Attempt to input a product key
cscript //nologo ospp.vbs /inpkey:FXYTK-NJJ8C-GB6DW-3DYQT-6F7TH >nul || goto notsupported

:skms
if %i% GTR 10 goto busy

:: Set KMS server based on the attempt count
if %i% EQU 1 set KMS=kms7.MSGuides.com
if %i% EQU 2 set KMS=107.175.77.7
if %i% GTR 2 goto ato

cscript //nologo ospp.vbs /sethst:%KMS% >nul

:ato
echo =====================================================================================
echo.

:: Attempt activation and check for success
if cscript //nologo ospp.vbs /act | find /i "successful" >nul (
    echo Activation successful!
    echo # Visit: MSGuides.com for more info.
    echo # How it works: bit.ly/kms-server
    echo # Please feel free to contact me at msguides.com@gmail.com if you have any questions or concerns.
    echo # Please consider supporting this project: donate.msguides.com
    echo # Your support is helping me keep my servers running 24/7!
    echo.
    echo =====================================================================================
    choice /n /c YN /m "Would you like to visit my blog [Y,N]?" 
    if errorlevel 2 exit
) else (
    echo The connection to my KMS server failed! Trying to connect to another one...
    echo Please wait...
    echo.
    set /a i+=1
    goto skms
)

explorer "http://MSGuides.com"
goto halt

:notsupported
echo =====================================================================================
echo.
echo Sorry, your version is not supported.
echo.
goto halt

:busy
echo =====================================================================================
echo.
echo Sorry, the server is busy and can't respond to your request. Please try again.
echo.

:halt
pause >nul
