@echo off
cls
:: ==========================================================================
::  MT_PropertyExchange - ROLLBACK to the old 32-bit installation
::
::  Use this if the new 64-bit installation causes problems:
::    1. removes the new toolkit COMPLETELY (files + all registrations,
::       32-bit and 64-bit - via cleanup.bat /auto)
::    2. starts the OLD 32-bit installer again:
::         a) if an installer exe is placed in the "old-installer" folder
::            next to this script, it is started automatically
::         b) otherwise the official Primetals share folder is opened and
::            you run "Install NEW DLLs for batchimport" from there
::
::  Close ALL SAP and Office windows before running this script.
:: ==========================================================================

:: BatchGotAdmin
>nul 2>&1 "%SYSTEMROOT%\system32\cacls.exe" "%SYSTEMROOT%\system32\config\system"
if "%errorlevel%"=="2" goto gotAdmin
if NOT "%errorlevel%"=="0" goto UACPrompt
goto gotAdmin
:UACPrompt
    echo Requesting administrative privileges...
    echo Set UAC = CreateObject^("Shell.Application"^) > "%temp%\getadmin.vbs"
    echo UAC.ShellExecute "%~s0", "", "", "runas", 1 >> "%temp%\getadmin.vbs"
    "%temp%\getadmin.vbs"
    exit /B
:gotAdmin
    if exist "%temp%\getadmin.vbs" del "%temp%\getadmin.vbs"
    CD /D "%~dp0"

:: Escape WOW64 so registry/tools are not redirected
if not defined PROCESSOR_ARCHITEW6432 goto bitness_ok
if exist "%WINDIR%\sysnative\cmd.exe" "%WINDIR%\sysnative\cmd.exe" /c ""%~f0"" & exit /b
:bitness_ok

color 8F
set SHARE=\\pmt.primetals.net\ptdfs\at\apps\Software-DWP\Public\SAP

echo.
echo === ROLLBACK: remove 64-bit toolkit, restore old 32-bit installation ===
echo.
echo Close ALL SAP and Office windows now, otherwise files stay locked.
pause

if exist "%~dp0cleanup.bat" goto havecleanup
color CF
echo ERROR: cleanup.bat not found next to this script.
echo Keep rollback.bat and cleanup.bat in the same folder.
pause
exit /b 1
:havecleanup

echo.
echo [1/2] Removing the new installation completely ...
call "%~dp0cleanup.bat" /auto
if "%errorlevel%"=="0" goto cleanok
color CF
echo.
echo ERROR: cleanup could not remove everything - files are still locked.
echo Close SAP/Office or REBOOT, then run rollback.bat again.
pause
exit /b 1
:cleanok

echo.
echo [2/2] Restoring the old 32-bit installation ...

rem --- a) local old installer provided? ------------------------------------
set OLDSETUP=
for %%f in ("%~dp0old-installer\*.exe") do set OLDSETUP=%%~ff
if not defined OLDSETUP goto tryshare
echo Starting the old installer: "%OLDSETUP%"
echo Follow its prompts; this window stays open until it finishes.
start /wait "" "%OLDSETUP%"
goto done

:tryshare
rem --- b) open the official share ------------------------------------------
if not exist "%SHARE%" goto noshare
echo Opening the official package folder:
echo    %SHARE%
echo Run "Install NEW DLLs for batchimport" from there.
start "" explorer "%SHARE%"
goto done

:noshare
color CF
echo WARNING: the share %SHARE%
echo is not reachable from this PC. Get the old installer package from IT
echo or place its exe into the folder "old-installer" next to this script
echo and run rollback.bat again.
pause
exit /b 1

:done
color 2F
echo.
echo Rollback prepared: the new toolkit is removed. After the old installer
echo has finished, test with transaction ZBATIMP.
echo NOTE: the old package is 32-bit only - the 64-bit SAP Business Client
echo will NOT work with it; use the 32-bit SAP GUI until the new toolkit is
echo installed again.
echo.
pause
exit /b 0
