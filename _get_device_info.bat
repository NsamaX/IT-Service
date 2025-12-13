@echo off
title Asset Inventory Report
chcp 65001 >nul
setlocal enabledelayedexpansion

:: Get current date and time
for /f "tokens=2 delims==" %%a in ('wmic os get localdatetime /value') do set dt=%%a
set "date_str=%dt:~0,4%/%dt:~4,2%/%dt:~6,2%"
set "time_str=%dt:~8,2%:%dt:~10,2%:%dt:~12,2%.%dt:~15,2%"

:: Find SANDISK Flash Drive automatically
set "output_path="
for %%d in (D E F G H I J K L M N O P Q R S T U V W X Y Z) do (
    if exist "%%d:\" (
        for /f "tokens=*" %%v in ('vol %%d: 2^>nul ^| findstr /i "SANDISK"') do (
            set "output_path=%%d:\"
        )
    )
)

:: If SANDISK not found, use Desktop
if not defined output_path (
    set "output_path=%USERPROFILE%\Desktop\"
    echo Warning: SANDISK drive not found. Saving to Desktop instead.
    timeout /t 3 >nul
)

set "filename=%computername%_%dt:~0,4%-%dt:~4,2%-%dt:~6,2%.txt"
set "fullpath=%output_path%%filename%"

:: Create report header
(
echo ====================================================
echo --- IT DEVICE INVENTORY DATA ---
echo Date/Time: %date_str%  %time_str%
echo ====================================================
echo.
) > "%fullpath%"

:: [1] Brand
for /f "tokens=2 delims==" %%a in ('wmic csproduct get Vendor /value ^| find "="') do (
    echo [1] Brand: %%a>> "%fullpath%"
)
echo.>> "%fullpath%"

:: [2] Model
for /f "tokens=2 delims==" %%a in ('wmic csproduct get Name /value ^| find "="') do (
    echo [2] Model: %%a>> "%fullpath%"
)
echo.>> "%fullpath%"

:: [3] Spec
echo [3] Spec:>> "%fullpath%"

:: CPU
for /f "tokens=2 delims==" %%a in ('wmic cpu get Name /value ^| find "="') do (
    echo     CPU: %%a>> "%fullpath%"
)

:: RAM
for /f %%a in ('powershell -command "[math]::Round((Get-WmiObject Win32_ComputerSystem).TotalPhysicalMemory/1GB)"') do (
    echo     RAM: %%a GB>> "%fullpath%"
)

:: HDD/SSD (excluding USB drives)
for /f "skip=1 tokens=1,2,3 delims=," %%a in ('wmic diskdrive get Caption^,Size^,InterfaceType /format:csv') do (
    if not "%%b"=="" (
        echo %%c | findstr /i "USB" >nul
        if errorlevel 1 (
            echo     HDD/SSD: %%b %%c>> "%fullpath%"
        )
    )
)
echo.>> "%fullpath%"

:: [4] Operating System
for /f "tokens=2 delims==" %%a in ('wmic os get Caption /value ^| find "="') do (
    set "os_name=%%a"
)
for /f "tokens=2 delims==" %%a in ('wmic os get OSArchitecture /value ^| find "="') do (
    set "os_arch=%%a"
)
echo [4] Operating System: !os_name! !os_arch!>> "%fullpath%"
echo.>> "%fullpath%"

:: [5] Device ID
for /f %%a in ('powershell -command "(Get-WmiObject Win32_ComputerSystemProduct).UUID"') do (
    echo [5] Device ID: %%a>> "%fullpath%"
)
echo.>> "%fullpath%"

:: [6] Product ID
for /f "skip=2 tokens=2*" %%a in ('reg query "HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion" /v ProductId 2^>nul') do (
    echo [6] Product ID: %%b>> "%fullpath%"
)
echo.>> "%fullpath%"

:: [7] Serial No
for /f "tokens=2 delims==" %%a in ('wmic bios get SerialNumber /value ^| find "="') do (
    echo [7] Serial No: %%a>> "%fullpath%"
)
echo.>> "%fullpath%"

:: [8] Device Name
for /f "tokens=2 delims==" %%a in ('wmic computersystem get Name /value ^| find "="') do (
    echo [8] Device Name: %%a>> "%fullpath%"
)
echo.>> "%fullpath%"

:: [9] User Email
echo [9] User Email: %username%>> "%fullpath%"
echo.>> "%fullpath%"

:: [10] Microsoft Office Information
echo [10] Microsoft Office:>> "%fullpath%"
set "office_found=0"

:: Check Office version from Registry (Office 2016, 2019, 2021, 365)
for %%v in (16.0 15.0 14.0) do (
    for /f "skip=2 tokens=2*" %%a in ('reg query "HKLM\SOFTWARE\Microsoft\Office\%%v\Registration" /s /v ProductName 2^>nul ^| findstr "ProductName"') do (
        echo     Version: %%b>> "%fullpath%"
        set "office_found=1"
        goto :office_version_found
    )
)
:office_version_found

:: Get Product ID
for %%v in (16.0 15.0 14.0) do (
    for /f "skip=2 tokens=2*" %%a in ('reg query "HKLM\SOFTWARE\Microsoft\Office\%%v\Registration" /s /v ProductID 2^>nul ^| findstr "ProductID"') do (
        echo     Product ID: %%b>> "%fullpath%"
        goto :office_pid_found
    )
)
:office_pid_found

:: Check License Status
powershell -command "try { $null = Get-WmiObject -query 'select * from SoftwareLicensingProduct where ApplicationID=\"55c92734-d682-4d71-983e-d6ec3f16059f\" and LicenseStatus=1' -ErrorAction Stop | Where-Object { $_.Name -like '*Office*' } | Select-Object -First 1 -ExpandProperty Name; if ($?) { Write-Output 'Licensed' } else { Write-Output 'Not found' } } catch { Write-Output 'Not found' }" > "%temp%\office_lic.txt" 2>nul
for /f "delims=" %%a in (%temp%\office_lic.txt) do (
    if not "%%a"=="Not found" (
        echo     License Status: %%a>> "%fullpath%"
    )
)
del "%temp%\office_lic.txt" 2>nul

:: Check Click-to-Run installation (Office 365)
for /f "skip=2 tokens=2*" %%a in ('reg query "HKLM\SOFTWARE\Microsoft\Office\ClickToRun\Configuration" /v ProductReleaseIds 2^>nul') do (
    echo     Edition: %%b>> "%fullpath%"
    set "office_found=1"
)

if !office_found!==0 (
    echo     Status: Not installed or not detected>> "%fullpath%"
)
echo.>> "%fullpath%"

echo ====================================================
echo SUCCESS! Report saved to:
echo %fullpath%
echo ====================================================
echo.
echo Press any key to close this window...
pause > nul
