@echo off
title Activate Office 365 ProPlus for FREE - MSGuides.com
cls

echo =====================================================================================
echo # Project: Activating Microsoft software products for FREE without additional software
echo =====================================================================================
echo.
echo # Supported products: Office 365 ProPlus (x86-x64)
echo.

:: Change directory to the Office installation folder if ospp.vbs exists
if exist "%ProgramFiles%\Microsoft Office\Office16\ospp.vbs" (
    cd /d "%ProgramFiles%\Microsoft Office\Office16"
) else if exist "%ProgramFiles(x86)%\Microsoft Office\Office16\ospp.vbs" (
    cd /d "%ProgramFiles(x86)%\Microsoft Office\Office16"
)

:: Install the KMS licenses
for /f %%x in ('dir /b ..\root\Licenses16\proplusvl_kms*.xrm-ms') do (
    cscript ospp.vbs /inslic:"..\root\Licenses16\%%x" >nul
)
for /f %%x in ('dir /b ..\root\Licenses16\proplusvl_mak*.xrm-ms') do (
    cscript ospp.vbs /inslic:"..\root\Licenses16\%%x" >nul
)

echo.
echo ============================================================================

echo Activating your Office...
cscript //nologo slmgr.vbs /ckms >nul
cscript //nologo ospp.vbs /setprt:1688 >nul
cscript //nologo ospp.vbs /unpkey:WFG99 >nul
cscript //nologo ospp.vbs /unpkey:DRTFM >nul
cscript //nologo ospp.vbs /unpkey:BTDRB >nul

set i=1

:: Attempt to input a product key
cscript //nologo ospp.vbs /inpkey:XQNVK-8JYDB-WJ9W3-YJ8YR-WFG99 >nul || ^
cscript //nologo ospp.vbs /inpkey:NMMKJ-6RK4F-KMJVX-8D9MJ-6MWKP >nul || ^
goto notsupported

:skms
if %i% GTR 10 goto busy

:: Set KMS server based on the attempt count
if %i% EQU 1 set KMS=kms7.MSGuides.com
if %i% EQU 2 set KMS=107.175.77.7
if %i% GTR 2 goto ato

cscript //nologo ospp.vbs /sethst:%KMS% >nul

:ato
echo ============================================================================

:: Attempt activation and check for success
if cscript //nologo ospp.vbs /act | find /i "successful" >nul (
    echo Activation successful!
    echo # Visit: MSGuides.com for more info.
    goto halt
) else (
    echo The connection to my KMS server failed! Trying to connect to another one...
    echo Please wait...
    set /a i+=1
    goto skms
)

explorer "http://MSGuides.com"
goto halt

:notsupported
echo ============================================================================ 
echo Sorry, your version is not supported.
goto halt

:busy
echo ============================================================================ 
echo Sorry, the server is busy and can't respond to your request. Please try again.
echo.

:halt
pause >nul
