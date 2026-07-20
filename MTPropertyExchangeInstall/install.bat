@echo off
cls

:: ==========================================================================
::  MT_PropertyExchange toolkit - installer
::
::  Installs the toolkit and registers the COM components in BOTH registry
::  views on 64-bit Windows, so that 32-bit callers (legacy SAP GUI, SEAL
::  scripts) and 64-bit callers (SAP Business Client 64-bit) both work.
::
::  Expected payload layout (created by stage.bat from the build output):
::     install.bat, uninstall.bat
::     bin\   PropertyExchange.exe, MT_PropertyExchangeDLL.dll, CommandLine.dll,
::            Interop.Dsofile.dll, log4net.dll, MT_PropertyExchangeLog.config, ...
::     x86\   dsofile.dll   (32-bit build)
::     x64\   dsofile.dll   (64-bit build)
::     test\  smoke tests and sample documents
::
::  Usage:
::     install.bat               ... install + register + quick test
::     install.bat /u            ... unregister + remove installation
::     install.bat /unregister   ... unregister only
::     install.bat /register     ... (re-)register only
:: ==========================================================================

:: BatchGotAdmin
:-------------------------------------
REM  --> Check for permissions
>nul 2>&1 "%SYSTEMROOT%\system32\cacls.exe" "%SYSTEMROOT%\system32\config\system"
REM --> If error flag set, we do not have admin.

if "%errorlevel%"=="2" goto gotAdmin
if NOT "%errorlevel%"=="0" goto UACPrompt
goto gotAdmin

:UACPrompt
    echo Requesting administrative privileges...
    echo Set UAC = CreateObject^("Shell.Application"^) > "%temp%\getadmin.vbs"
    if "%1" == "" echo UAC.ShellExecute "%~s0", "", "", "runas", 1 >> "%temp%\getadmin.vbs"
    if "%1" == "/u" echo UAC.ShellExecute "%~s0", "/u", "", "runas", 1 >> "%temp%\getadmin.vbs"
    if "%1" == "/unregister" echo UAC.ShellExecute "%~s0", "/unregister", "", "runas", 1 >> "%temp%\getadmin.vbs"
    if "%1" == "/register" echo UAC.ShellExecute "%~s0", "/register", "", "runas", 1 >> "%temp%\getadmin.vbs"
    "%temp%\getadmin.vbs"
    exit /B

:gotAdmin
    if exist "%temp%\getadmin.vbs" ( del "%temp%\getadmin.vbs" )
    pushd "%CD%"
    CD /D "%~dp0"

:--------------------------------------
:start
   pushd "%~dp0"
   set INSTALL_ROOT=%CD%
   mode con lines=50 cols=100
   color 8F

   echo.
   if "%1" == "" echo Installation of MT_PropertyExchange toolkit
   if "%1" == "/u" echo Uninstall MT_PropertyExchange toolkit
   if "%1" == "/unregister" echo Unregistering COM components
   if "%1" == "/register" echo Registering COM components
   echo.

   set TARGET_DIR=%ProgramFiles%\Siemens\MT_PropertyExchange
   set REGASM32=%WINDIR%\Microsoft.NET\Framework\v4.0.30319\regasm.exe
   set REGASM64=%WINDIR%\Microsoft.NET\Framework64\v4.0.30319\regasm.exe
   set REGSVR_NATIVE=%WINDIR%\System32\regsvr32.exe
   set REGSVR_WOW=%WINDIR%\SysWOW64\regsvr32.exe

   set IS64=NO
   if defined ProgramFiles(x86) set IS64=YES

   if not exist "%REGASM32%" echo WARNING: 32-bit .NET Framework 4.x regasm.exe not found.
   if "%IS64%"=="YES" if not exist "%REGASM64%" echo WARNING: 64-bit .NET Framework 4.x regasm.exe not found.

   if "%1"=="/u" goto unreg
   if "%1"=="/unregister" goto unreg
   if "%1"=="/register" goto reg

REM ==========================================================================
REM  Full installation: unregister old state, copy files, register, test
REM ==========================================================================
:install
   if exist "%INSTALL_ROOT%\bin\PropertyExchange.exe" goto payloadok
   color C0
   echo.
   echo ERROR: installer payload is incomplete ("bin\PropertyExchange.exe" not found).
   echo Run MTPropertyExchangeInstall\stage.bat after building, then retry.
   echo.
   pause
   goto ende
:payloadok
   call :dounreg

:copyfiles
   echo.
   echo Copying toolkit to "%TARGET_DIR%" ...
   mkdir "%TARGET_DIR%" 2>NUL
   xcopy /Y /C /Q "%INSTALL_ROOT%\bin\*.*" "%TARGET_DIR%\"
   xcopy /Y /C /Q /I "%INSTALL_ROOT%\x86\*.*" "%TARGET_DIR%\x86\"
   if "%IS64%"=="YES" xcopy /Y /C /Q /I "%INSTALL_ROOT%\x64\*.*" "%TARGET_DIR%\x64\"
   echo Copying finished.

:reg
   echo.
   echo Registering COM components ...

   if "%IS64%"=="NO" goto reg32only

   rem ---- 64-bit Windows: register in BOTH registry views --------------------
   if exist "%TARGET_DIR%\x64\dsofile.dll" echo   dsofile.dll (x64, 64-bit registry view)
   if exist "%TARGET_DIR%\x64\dsofile.dll" "%REGSVR_NATIVE%" /s "%TARGET_DIR%\x64\dsofile.dll"
   if not exist "%TARGET_DIR%\x64\dsofile.dll" echo   WARNING: x64\dsofile.dll missing - 64-bit clients cannot use .doc/.xls files!

   if exist "%TARGET_DIR%\x86\dsofile.dll" echo   dsofile.dll (x86, 32-bit registry view)
   if exist "%TARGET_DIR%\x86\dsofile.dll" "%REGSVR_WOW%" /s "%TARGET_DIR%\x86\dsofile.dll"
   if not exist "%TARGET_DIR%\x86\dsofile.dll" echo   WARNING: x86\dsofile.dll missing - 32-bit clients cannot use .doc/.xls files!

   if not exist "%TARGET_DIR%\MT_PropertyExchangeDLL.dll" goto regdone
   echo   MT_PropertyExchangeDLL.dll (64-bit registry view)
   if exist "%REGASM64%" "%REGASM64%" "%TARGET_DIR%\MT_PropertyExchangeDLL.dll" /codebase /tlb:"%TARGET_DIR%\MT_PropertyExchangeDLL.tlb" /nologo
   echo   MT_PropertyExchangeDLL.dll (32-bit registry view)
   if exist "%REGASM32%" "%REGASM32%" "%TARGET_DIR%\MT_PropertyExchangeDLL.dll" /codebase /tlb:"%TARGET_DIR%\MT_PropertyExchangeDLL.tlb" /nologo
   goto regdone

:reg32only
   rem ---- 32-bit Windows ------------------------------------------------------
   if exist "%TARGET_DIR%\x86\dsofile.dll" echo   dsofile.dll (x86)
   if exist "%TARGET_DIR%\x86\dsofile.dll" "%REGSVR_NATIVE%" /s "%TARGET_DIR%\x86\dsofile.dll"
   if not exist "%TARGET_DIR%\MT_PropertyExchangeDLL.dll" goto regdone
   echo   MT_PropertyExchangeDLL.dll
   if exist "%REGASM32%" "%REGASM32%" "%TARGET_DIR%\MT_PropertyExchangeDLL.dll" /codebase /tlb:"%TARGET_DIR%\MT_PropertyExchangeDLL.tlb" /nologo

:regdone
   echo Registration finished.
   if "%1"=="/register" goto fine

REM ==========================================================================
:test
   echo.
   echo Quick test: COM activation check.
   mkdir "%TARGET_DIR%\Test" 2>NUL
   xcopy /Q /S /Y "%INSTALL_ROOT%\test\*.*" "%TARGET_DIR%\Test\" 1>nul 2>nul
   call "%TARGET_DIR%\Test\test-com.bat"
   IF "%ERRORLEVEL%"=="0" COLOR 2F
   echo.
   echo End of test.
   goto fine

REM ==========================================================================
:unreg
   call :dounreg
   if "%1"=="/unregister" goto fine
   if "%1"=="/u" goto uninstall
   goto fine

:uninstall
   echo Removing "%TARGET_DIR%" ...
   if exist "%TARGET_DIR%" rmdir /q /s "%TARGET_DIR%"
   goto fine

REM ==========================================================================
REM  Subroutine: unregister everything (current and legacy locations)
REM ==========================================================================
:dounreg
   echo.
   echo Unregistering COM components (previous installations) ...

   if exist "%TARGET_DIR%\MT_PropertyExchangeDLL.dll" if exist "%REGASM32%" "%REGASM32%" "%TARGET_DIR%\MT_PropertyExchangeDLL.dll" /codebase /tlb:"%TARGET_DIR%\MT_PropertyExchangeDLL.tlb" /nologo /unregister 2>NUL
   if "%IS64%"=="NO" goto dounreg_dso
   if exist "%TARGET_DIR%\MT_PropertyExchangeDLL.dll" if exist "%REGASM64%" "%REGASM64%" "%TARGET_DIR%\MT_PropertyExchangeDLL.dll" /codebase /tlb:"%TARGET_DIR%\MT_PropertyExchangeDLL.tlb" /nologo /unregister 2>NUL

:dounreg_dso
   rem new locations
   if exist "%TARGET_DIR%\x86\dsofile.dll" "%REGSVR_WOW%" /s /u "%TARGET_DIR%\x86\dsofile.dll" 2>NUL
   if exist "%TARGET_DIR%\x64\dsofile.dll" "%REGSVR_NATIVE%" /s /u "%TARGET_DIR%\x64\dsofile.dll" 2>NUL
   if "%IS64%"=="YES" goto dounreg_legacy
   if exist "%TARGET_DIR%\x86\dsofile.dll" "%REGSVR_NATIVE%" /s /u "%TARGET_DIR%\x86\dsofile.dll" 2>NUL

:dounreg_legacy
   rem legacy location: dsofile.dll flat in the target dir (old installer)
   if exist "%TARGET_DIR%\dsofile.dll" "%REGSVR_NATIVE%" /s /u "%TARGET_DIR%\dsofile.dll" 2>NUL
   if "%IS64%"=="YES" if exist "%TARGET_DIR%\dsofile.dll" "%REGSVR_WOW%" /s /u "%TARGET_DIR%\dsofile.dll" 2>NUL
   echo Unregistering finished.
   goto :EOF

REM ==========================================================================
:fine
   echo.
   echo Done. This window will close in a few seconds...
   timeout /t 5 >NUL 2>NUL
   popd 2>NUL

:ende
popd 2>NUL
