@echo off
cls
:: ==========================================================================
::  MT_PropertyExchange - complete cleanup
::
::  Removes ALL traces of the toolkit from this machine:
::    - COM registrations (32-bit and 64-bit views)
::    - both install folders:  C:\Program Files\Siemens\MT_PropertyExchange
::                             C:\Program Files (x86)\Siemens\MT_PropertyExchange
::      (the x86 path may be a compatibility junction - handled safely)
::    - leftover registry keys (ProgID, CLSID, TypeLib, DllSurrogate hack)
::
::  After this, NO property exchange works until a package is installed
::  again. Close all SAP and Office windows before running.
:: ==========================================================================

:: BatchGotAdmin
>nul 2>&1 "%SYSTEMROOT%\system32\cacls.exe" "%SYSTEMROOT%\system32\config\system"
if "%errorlevel%"=="2" goto gotAdmin
if NOT "%errorlevel%"=="0" goto UACPrompt
goto gotAdmin
:UACPrompt
    echo Requesting administrative privileges...
    echo Set UAC = CreateObject^("Shell.Application"^) > "%temp%\getadmin.vbs"
    echo UAC.ShellExecute "%~s0", "%1", "", "runas", 1 >> "%temp%\getadmin.vbs"
    "%temp%\getadmin.vbs"
    exit /B
:gotAdmin
    if exist "%temp%\getadmin.vbs" del "%temp%\getadmin.vbs"
    CD /D "%~dp0"

:: Escape WOW64 so registry/tools are not redirected
if not defined PROCESSOR_ARCHITEW6432 goto bitness_ok
if exist "%WINDIR%\sysnative\cmd.exe" "%WINDIR%\sysnative\cmd.exe" /c ""%~f0" %1" & exit /b
:bitness_ok

color 8F
set DIR64=%ProgramFiles%\Siemens\MT_PropertyExchange
set DIR86=
if defined ProgramFiles(x86) set DIR86=%ProgramFiles(x86)%\Siemens\MT_PropertyExchange
set REGASM32=%WINDIR%\Microsoft.NET\Framework\v4.0.30319\regasm.exe
set REGASM64=%WINDIR%\Microsoft.NET\Framework64\v4.0.30319\regasm.exe
set FAIL=0

echo.
echo === MT_PropertyExchange complete cleanup ===
echo.
echo Close ALL SAP and Office windows now, otherwise files stay locked.
rem /auto = no pauses, propagate result via exit code (used by rollback.bat)
if not "%1"=="/auto" pause

echo.
echo [1/4] Unregistering COM components ...
if exist "%DIR64%\MT_PropertyExchangeDLL.dll" if exist "%REGASM64%" "%REGASM64%" "%DIR64%\MT_PropertyExchangeDLL.dll" /nologo /unregister >NUL 2>NUL
if exist "%DIR64%\MT_PropertyExchangeDLL.dll" if exist "%REGASM32%" "%REGASM32%" "%DIR64%\MT_PropertyExchangeDLL.dll" /nologo /unregister >NUL 2>NUL
if exist "%DIR64%\x64\dsofile.dll" "%WINDIR%\System32\regsvr32.exe" /s /u "%DIR64%\x64\dsofile.dll" 2>NUL
if exist "%DIR64%\x86\dsofile.dll" "%WINDIR%\SysWOW64\regsvr32.exe" /s /u "%DIR64%\x86\dsofile.dll" 2>NUL
if exist "%DIR64%\dsofile.dll" "%WINDIR%\System32\regsvr32.exe" /s /u "%DIR64%\dsofile.dll" 2>NUL
if exist "%DIR64%\dsofile.dll" "%WINDIR%\SysWOW64\regsvr32.exe" /s /u "%DIR64%\dsofile.dll" 2>NUL

echo [2/4] Removing the x86 folder or compatibility junction ...
if not defined DIR86 goto dir86done
rem Plain rmdir removes ONLY a junction or an empty folder - never contents.
rmdir "%DIR86%" 2>NUL
if not exist "%DIR86%" goto dir86done
rem Still there: it is a real folder. Unregister its DLLs, then remove it.
if exist "%DIR86%\MT_PropertyExchangeDLL.dll" if exist "%REGASM32%" "%REGASM32%" "%DIR86%\MT_PropertyExchangeDLL.dll" /nologo /unregister >NUL 2>NUL
if exist "%DIR86%\MT_PropertyExchangeDLL.dll" if exist "%REGASM64%" "%REGASM64%" "%DIR86%\MT_PropertyExchangeDLL.dll" /nologo /unregister >NUL 2>NUL
if exist "%DIR86%\dsofile.dll" "%WINDIR%\SysWOW64\regsvr32.exe" /s /u "%DIR86%\dsofile.dll" 2>NUL
if exist "%DIR86%\x86\dsofile.dll" "%WINDIR%\SysWOW64\regsvr32.exe" /s /u "%DIR86%\x86\dsofile.dll" 2>NUL
if exist "%DIR86%\x64\dsofile.dll" "%WINDIR%\System32\regsvr32.exe" /s /u "%DIR86%\x64\dsofile.dll" 2>NUL
rmdir /s /q "%DIR86%" 2>NUL
if exist "%DIR86%" set FAIL=1
if exist "%DIR86%" echo    WARNING: could not remove "%DIR86%" completely - files locked?
:dir86done

echo [3/4] Removing the main install folder ...
if not exist "%DIR64%" goto dir64done
rmdir /s /q "%DIR64%" 2>NUL
if exist "%DIR64%" set FAIL=1
if exist "%DIR64%" echo    WARNING: could not remove "%DIR64%" completely - files locked?
:dir64done

echo [4/4] Removing leftover registry entries (both views) ...
rem MT_PropertyExchange.Globals
reg delete "HKCR\MT_PropertyExchange.Globals" /f >NUL 2>NUL
reg delete "HKCR\CLSID\{41A13AC0-103B-40EC-9A52-AEE2C6C846C6}" /f >NUL 2>NUL
reg delete "HKCR\Wow6432Node\CLSID\{41A13AC0-103B-40EC-9A52-AEE2C6C846C6}" /f >NUL 2>NUL
reg delete "HKCR\TypeLib\{E13AAF33-8EC6-4C2D-8EA4-489560C6EB12}" /f >NUL 2>NUL
reg delete "HKCR\Wow6432Node\TypeLib\{E13AAF33-8EC6-4C2D-8EA4-489560C6EB12}" /f >NUL 2>NUL
rem DllSurrogate hack leftovers
reg delete "HKCR\Wow6432Node\AppID\{41A13AC0-103B-40EC-9A52-AEE2C6C846C6}" /f >NUL 2>NUL
reg delete "HKLM\Software\Classes\AppID\{41A13AC0-103B-40EC-9A52-AEE2C6C846C6}" /f >NUL 2>NUL
rem DSOFile (only our registrations; a new package will re-register its own)
reg delete "HKCR\DSOFile.OleDocumentProperties" /f >NUL 2>NUL
reg delete "HKCR\CLSID\{58968145-CF05-4341-995F-2EE093F6ABA3}" /f >NUL 2>NUL
reg delete "HKCR\Wow6432Node\CLSID\{58968145-CF05-4341-995F-2EE093F6ABA3}" /f >NUL 2>NUL
reg delete "HKCR\TypeLib\{58968145-CF00-4341-995F-2EE093F6ABA3}" /f >NUL 2>NUL
reg delete "HKCR\Wow6432Node\TypeLib\{58968145-CF00-4341-995F-2EE093F6ABA3}" /f >NUL 2>NUL

echo.
if "%FAIL%"=="0" goto ok
color CF
echo RESULT: some folders could not be removed - files are locked.
echo Close SAP/Office or reboot, then run this script once more.
echo These processes are holding toolkit files:
tasklist /M MT_PropertyExchange* /FO TABLE 2>NUL
tasklist /M dsofile* /FO TABLE 2>NUL
goto ende
:ok
color 2F
echo RESULT: cleanup complete. The machine is clean - install ONE package
echo (the official PMT one or ours) before using SAP property exchange again.
:ende
echo.
if not "%1"=="/auto" pause
exit /b %FAIL%
