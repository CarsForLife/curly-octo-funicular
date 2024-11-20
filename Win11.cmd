@echo off
title Activate Windows 11 (ALL versions) 
cls

echo =====================================================================================
echo # Project: Activating Microsoft software products for FREE without additional software
echo =====================================================================================
echo.
echo # Supported products:
echo - Windows 11 Home
echo - Windows 11 Professional
echo - Windows 11 Education
echo - Windows 11 Enterprise
echo.
echo ============================================================================

echo Activating your Windows...

:: Clear KMS settings and uninstall previous keys
cscript //nologo slmgr.vbs /ckms >nul
cscript //nologo slmgr.vbs /upk >nul
cscript //nologo slmgr.vbs /cpky >nul

set i=1

:: Check the OS version and set the appropriate product key
wmic os | findstr /I "enterprise" >nul
if %errorlevel% EQU 0 (
    cscript //nologo slmgr.vbs /ipk NPPR9-FWDCX-D2C8J-H872K-2YT43 >nul || ^
    cscript //nologo slmgr.vbs /ipk DPH2V-TTNVB-4X9Q3-TJR4H-KHJW4 >nul || ^
    cscript //nologo slmgr.vbs /ipk YYVX9-NTFWV-6MDM3-9PT4T-4M68B >nul || ^
    cscript //nologo slmgr.vbs /ipk 44RPN-FTY23-9VTTB-MP9BX-T84FV >nul || ^
    cscript //nologo slmgr.vbs /ipk WNMTR-4C88C-JK8YV-HQ7T2-76DF9 >nul || ^
    cscript //nologo slmgr.vbs /ipk 2F77B-TNFGY-69QQF-B8YKP-D69TJ >nul || ^
    cscript //nologo slmgr.vbs /ipk DCPHK-NFMTC-H88MJ-PFHPY-QJ4BJ >nul || ^
    cscript //nologo slmgr.vbs /ipk QFFDN-GRT3P-VKWWX-X7T3R-8B639 >nul || ^
    cscript //nologo slmgr.vbs /ipk M7XTQ-FN8P6-TTKYV-9D4CC-J462D >nul || ^
    cscript //nologo slmgr.vbs /ipk 92NFX-8DJQP-P6BBQ-THF9C-7CG2H >nul
    goto skms
) else (
    wmic os | findstr /I "home" >nul
    if %errorlevel% EQU 0 (
        cscript //nologo slmgr.vbs /ipk TX9XD-98N7V-6WMQ6-BX7FG-H8Q99 >nul || ^
        cscript //nologo slmgr.vbs /ipk 3KHY7-WNT83-DGQKR-F7HPR-844BM >nul || ^
        cscript //nologo slmgr.vbs /ipk 7HNRX-D7KGG-3K4RQ-4WPJ4-YTDFH >nul || ^
        cscript //nologo slmgr.vbs /ipk PVMJN-6DFY6-9CCP6-7BKTT-D3WVR >nul
        goto skms
    ) else (
        wmic os | findstr /I "education" >nul
        if %errorlevel% EQU 0 (
            cscript //nologo slmgr.vbs /ipk NW6C2-QMPVW-D7KKK-3GKT6-VCFB2 >nul || ^
            cscript //nologo slmgr.vbs /ipk 2WH4N-8QGBV-H22JP-CT43Q-MDWWJ >nul
            goto skms
        ) else (
            wmic os | findstr /I "11 pro" >nul
            if %errorlevel% EQU 0 (
                cscript //nologo slmgr.vbs /ipk W269N-WFGWX-YVC9B-4J6C9-T83GX >nul || ^
                cscript //nologo slmgr.vbs /ipk MH37W-N47XK-V7XM9-C7227-GCQG9 >nul || ^
                cscript //nologo slmgr.vbs /ipk NRG8B-V 7F9C-4F3C9-8F4F9-9F4F9 >nul
                goto skms
            )
        )
    )
)

echo Sorry, your version is not supported.
exit /b

:skms
if %i% GTR 10 goto busy
set KMS=kms%[i].MSGuides.com
cscript //nologo slmgr.vbs /skms %KMS%:1688 >nul

:: Attempt activation
:ato
echo ============================================================================

if cscript //nologo slmgr.vbs /ato | find /i "successfully" >nul
(
    echo Activation successful!
    echo 
    exit /b
) else (
    echo Activation failed! Trying another server...
    set /a i+=1
    goto skms
)

:busy
echo ============================================================================ 
echo Sorry, the server is busy. Please try again later.
exit /b
