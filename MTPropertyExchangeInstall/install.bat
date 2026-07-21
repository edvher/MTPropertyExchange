@echo off
cls

:: ==========================================================================
::  MT_PropertyExchange toolkit - installer
::
::  Installs the toolkit and registers the COM components in BOTH registry
::  views on 64-bit Windows, so that 32-bit callers (legacy SAP GUI, SEAL
::  scripts) and 64-bit callers (SAP Business Client 64-bit) both work.
::
::  Every step is written to a log file (%TEMP%\MT_PropertyExchange_install.log).
::  The window turns GREEN when everything succeeded, RED when something
::  failed (the full log is then shown), and waits for ENTER before closing.
::
::  Expected payload layout (created by stage.bat from the build output):
::     install.bat, uninstall.bat
::     bin\   MT_PropertyExchange.exe, MT_PropertyExchangeDLL.dll, CommandLine.dll,
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

set LOGFILE=%TEMP%\MT_PropertyExchange_install.log

:: --------------------------------------------------------------------------
:: Escape WOW64: when started from a 32-bit parent (e.g. the self-extracting
:: installer stub), this script runs in a 32-bit cmd.exe where %ProgramFiles%
:: means "Program Files (x86)" and System32 tools are silently redirected to
:: their 32-bit versions. Relaunch ourselves in the real 64-bit cmd.exe.
:: --------------------------------------------------------------------------
if not defined PROCESSOR_ARCHITEW6432 goto bitness_ok
if exist "%WINDIR%\sysnative\cmd.exe" "%WINDIR%\sysnative\cmd.exe" /c ""%~f0" %1" & exit /b
:bitness_ok

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
    rem Copy the payload to a stable folder first: when started from the
    rem self-extracting installer the extraction folder disappears as soon
    rem as this (non-elevated) instance returns.
    set STAGE_DIR=%TEMP%\MTPE_Setup
    rmdir /q /s "%STAGE_DIR%" 2>NUL
    xcopy /E /I /Q /Y "%~dp0*" "%STAGE_DIR%\" >NUL
    echo Set UAC = CreateObject^("Shell.Application"^) > "%temp%\getadmin.vbs"
    if "%1" == "" echo UAC.ShellExecute "%STAGE_DIR%\install.bat", "", "", "runas", 1 >> "%temp%\getadmin.vbs"
    if "%1" == "/u" echo UAC.ShellExecute "%STAGE_DIR%\install.bat", "/u", "", "runas", 1 >> "%temp%\getadmin.vbs"
    if "%1" == "/unregister" echo UAC.ShellExecute "%STAGE_DIR%\install.bat", "/unregister", "", "runas", 1 >> "%temp%\getadmin.vbs"
    if "%1" == "/register" echo UAC.ShellExecute "%STAGE_DIR%\install.bat", "/register", "", "runas", 1 >> "%temp%\getadmin.vbs"
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
   set ERRORS=0
   color 8F

   set BUILDINFO=unknown - built before build stamping was introduced
   if exist "%INSTALL_ROOT%\build-info.txt" set /p BUILDINFO=<"%INSTALL_ROOT%\build-info.txt"

    > "%LOGFILE%" echo ================================================================
   >>"%LOGFILE%" echo  MT_PropertyExchange installer log
   >>"%LOGFILE%" echo  Date ......: %DATE% %TIME%
   >>"%LOGFILE%" echo  Build .....: %BUILDINFO%
   >>"%LOGFILE%" echo  User ......: %USERNAME%   Computer: %COMPUTERNAME%
   >>"%LOGFILE%" echo  Source ....: %INSTALL_ROOT%
   >>"%LOGFILE%" echo  Argument ..: %1
   >>"%LOGFILE%" echo  CmdBitness : PROCESSOR_ARCHITECTURE=%PROCESSOR_ARCHITECTURE% ARCHITEW6432=%PROCESSOR_ARCHITEW6432%
   >>"%LOGFILE%" echo ================================================================
   >>"%LOGFILE%" echo --- payload inventory (%INSTALL_ROOT%) ---
   dir /s /b "%INSTALL_ROOT%" >>"%LOGFILE%" 2>&1
   >>"%LOGFILE%" echo --- end of payload inventory ---

   echo.
   if "%1" == "" call :log Installation of MT_PropertyExchange toolkit
   if "%1" == "/u" call :log Uninstall MT_PropertyExchange toolkit
   if "%1" == "/unregister" call :log Unregistering COM components
   if "%1" == "/register" call :log Registering COM components
   echo.

   set TARGET_DIR=%ProgramFiles%\Siemens\MT_PropertyExchange
   set REGASM32=%WINDIR%\Microsoft.NET\Framework\v4.0.30319\regasm.exe
   set REGASM64=%WINDIR%\Microsoft.NET\Framework64\v4.0.30319\regasm.exe
   set REGSVR_NATIVE=%WINDIR%\System32\regsvr32.exe
   set REGSVR_WOW=%WINDIR%\SysWOW64\regsvr32.exe

   set IS64=NO
   if defined ProgramFiles(x86) set IS64=YES
   call :log Target folder: "%TARGET_DIR%"  -  64-bit Windows: %IS64%

   if exist "%REGASM32%" goto regasm32ok
   set /a ERRORS+=1
   call :log ERROR - 32-bit .NET Framework 4.x regasm.exe not found.
:regasm32ok
   if "%IS64%"=="NO" goto regasmchecked
   if exist "%REGASM64%" goto regasmchecked
   set /a ERRORS+=1
   call :log ERROR - 64-bit .NET Framework 4.x regasm.exe not found.
:regasmchecked

   if "%1"=="/u" goto unreg
   if "%1"=="/unregister" goto unreg
   if "%1"=="/register" goto reg

REM ==========================================================================
REM  Full installation: unregister old state, copy files, register, test
REM ==========================================================================
:install
   if exist "%INSTALL_ROOT%\bin\MT_PropertyExchange.exe" goto payloadok
   set /a ERRORS+=1
   call :log ERROR - installer payload is incomplete: "bin\MT_PropertyExchange.exe" not found.
   call :log Run MTPropertyExchangeInstall\stage.bat after building, then retry.
   goto summary
:payloadok
   if defined ProgramFiles(x86) if exist "%ProgramFiles(x86)%\Siemens\MT_PropertyExchange" if /I not "%TARGET_DIR%"=="%ProgramFiles(x86)%\Siemens\MT_PropertyExchange" call :log NOTE - an old installation exists in "Program Files (x86)\Siemens\MT_PropertyExchange". It is removed automatically after a successful installation.
   rem Diagnostic: which running processes have toolkit modules loaded?
   rem (Those processes would lock files and make copy/regasm steps fail.)
   >>"%LOGFILE%" echo --- processes holding toolkit modules ---
   tasklist /M MT_PropertyExchange* /FO TABLE >>"%LOGFILE%" 2>&1
   tasklist /M Interop.Dsofile* /FO TABLE >>"%LOGFILE%" 2>&1
   tasklist /M dsofile* /FO TABLE >>"%LOGFILE%" 2>&1
   >>"%LOGFILE%" echo --- end of process diagnostic ---
   call :dounreg

:copyfiles
   echo.
   if not exist "%TARGET_DIR%" goto docopy
   rem The target folder already exists (previous installation, possibly a
   rem mix of packages). Remove it completely so the installation is clean.
   call :log Target folder already exists - removing it for a clean installation ...
   rmdir /s /q "%TARGET_DIR%" 2>NUL
   if not exist "%TARGET_DIR%" goto docopy
   set /a ERRORS+=1
   call :log ERROR - cannot remove the existing folder, files are in use. Close SAP, Word/Excel and other programs using the toolkit, then run the installer again. The log lists the locking processes.
   goto summary
:docopy
   call :log Copying toolkit to "%TARGET_DIR%" ...
   mkdir "%TARGET_DIR%" 2>NUL
   >>"%LOGFILE%" echo --- xcopy bin ---
   xcopy /Y /F "%INSTALL_ROOT%\bin\*.*" "%TARGET_DIR%\" >>"%LOGFILE%" 2>&1
   if errorlevel 1 (set /a ERRORS+=1 & call :log   ERROR - copying program files failed - a file is probably in use. Close SAP, Word/Excel and other programs using the toolkit, then run the installer again. The log lists the locking processes.) else (call :log   OK - program files copied.)
   >>"%LOGFILE%" echo --- xcopy x86 ---
   xcopy /Y /F /I "%INSTALL_ROOT%\x86\*.*" "%TARGET_DIR%\x86\" >>"%LOGFILE%" 2>&1
   if errorlevel 1 (set /a ERRORS+=1 & call :log   ERROR - copying x86 dsofile failed.) else (call :log   OK - x86 dsofile.dll copied.)
   if "%IS64%"=="NO" goto copydone
   >>"%LOGFILE%" echo --- xcopy x64 ---
   xcopy /Y /F /I "%INSTALL_ROOT%\x64\*.*" "%TARGET_DIR%\x64\" >>"%LOGFILE%" 2>&1
   if errorlevel 1 (set /a ERRORS+=1 & call :log   ERROR - copying x64 dsofile failed.) else (call :log   OK - x64 dsofile.dll copied.)
:copydone
   >>"%LOGFILE%" echo --- installed files (%TARGET_DIR%) ---
   dir /s /b "%TARGET_DIR%" >>"%LOGFILE%" 2>&1
   >>"%LOGFILE%" echo --- end of installed files ---

:reg
   echo.
   call :log Registering COM components ...

   if "%IS64%"=="NO" goto reg32only

   rem ---- 64-bit Windows: register in BOTH registry views --------------------
   if not exist "%TARGET_DIR%\x64\dsofile.dll" goto regdso64missing
   >>"%LOGFILE%" echo --- regsvr32 x64 dsofile ---
   "%REGSVR_NATIVE%" /s "%TARGET_DIR%\x64\dsofile.dll"
   if errorlevel 1 (set /a ERRORS+=1 & call :log   ERROR - registering dsofile.dll x64 failed.) else (call :log   OK - dsofile.dll registered in 64-bit view.)
   goto regdso32
:regdso64missing
   set /a ERRORS+=1
   call :log   ERROR - x64\dsofile.dll missing. 64-bit clients cannot use .doc/.xls files.

:regdso32
   if not exist "%TARGET_DIR%\x86\dsofile.dll" goto regdso32missing
   >>"%LOGFILE%" echo --- regsvr32 x86 dsofile ---
   "%REGSVR_WOW%" /s "%TARGET_DIR%\x86\dsofile.dll"
   if errorlevel 1 (set /a ERRORS+=1 & call :log   ERROR - registering dsofile.dll x86 failed.) else (call :log   OK - dsofile.dll registered in 32-bit view.)
   goto regnet
:regdso32missing
   set /a ERRORS+=1
   call :log   ERROR - x86\dsofile.dll missing. 32-bit clients cannot use .doc/.xls files.

:regnet
   if exist "%TARGET_DIR%\MT_PropertyExchangeDLL.dll" goto regnetdo
   set /a ERRORS+=1
   call :log   ERROR - MT_PropertyExchangeDLL.dll not found in target folder.
   goto regdone
:regnetdo
   rem Remove the previously exported type library first; if it is locked by
   rem a running process, regasm would fail with a misleading access-denied.
   if exist "%TARGET_DIR%\MT_PropertyExchangeDLL.tlb" del /q "%TARGET_DIR%\MT_PropertyExchangeDLL.tlb" 2>NUL
   if exist "%TARGET_DIR%\MT_PropertyExchangeDLL.tlb" call :log   WARNING - the existing type library file is locked by a running process. Close SAP and Office programs and run the installer again if the next steps fail.
   >>"%LOGFILE%" echo --- regasm 64-bit ---
   "%REGASM64%" "%TARGET_DIR%\MT_PropertyExchangeDLL.dll" /codebase /tlb:"%TARGET_DIR%\MT_PropertyExchangeDLL.tlb" /nologo >>"%LOGFILE%" 2>&1
   if errorlevel 1 (set /a ERRORS+=1 & call :log   ERROR - regasm 64-bit failed.) else (call :log   OK - MT_PropertyExchangeDLL registered in 64-bit view.)
   >>"%LOGFILE%" echo --- regasm 32-bit ---
   "%REGASM32%" "%TARGET_DIR%\MT_PropertyExchangeDLL.dll" /codebase /tlb:"%TARGET_DIR%\MT_PropertyExchangeDLL.tlb" /nologo >>"%LOGFILE%" 2>&1
   if errorlevel 1 (set /a ERRORS+=1 & call :log   ERROR - regasm 32-bit failed.) else (call :log   OK - MT_PropertyExchangeDLL registered in 32-bit view.)
   goto regdone

:reg32only
   rem ---- 32-bit Windows ------------------------------------------------------
   if not exist "%TARGET_DIR%\x86\dsofile.dll" goto reg32dsomissing
   >>"%LOGFILE%" echo --- regsvr32 x86 dsofile ---
   "%REGSVR_NATIVE%" /s "%TARGET_DIR%\x86\dsofile.dll"
   if errorlevel 1 (set /a ERRORS+=1 & call :log   ERROR - registering dsofile.dll failed.) else (call :log   OK - dsofile.dll registered.)
   goto reg32net
:reg32dsomissing
   set /a ERRORS+=1
   call :log   ERROR - x86\dsofile.dll missing.
:reg32net
   if not exist "%TARGET_DIR%\MT_PropertyExchangeDLL.dll" goto regdone
   >>"%LOGFILE%" echo --- regasm 32-bit ---
   "%REGASM32%" "%TARGET_DIR%\MT_PropertyExchangeDLL.dll" /codebase /tlb:"%TARGET_DIR%\MT_PropertyExchangeDLL.tlb" /nologo >>"%LOGFILE%" 2>&1
   if errorlevel 1 (set /a ERRORS+=1 & call :log   ERROR - regasm failed.) else (call :log   OK - MT_PropertyExchangeDLL registered.)

:regdone
   call :log Registration finished.
   if "%1"=="/register" goto summary

REM ==========================================================================
:test
   echo.
   call :log Quick test: COM activation check ...
   mkdir "%TARGET_DIR%\Test" 2>NUL
   xcopy /Q /S /Y "%INSTALL_ROOT%\test\*.*" "%TARGET_DIR%\Test\" 1>nul 2>nul
   call "%TARGET_DIR%\Test\test-com.bat" > "%TEMP%\mtpe_testcom.out" 2>&1
   if errorlevel 1 (set /a ERRORS+=1 & call :log   ERROR - COM activation test FAILED.) else (call :log   OK - COM activation test passed.)
   type "%TEMP%\mtpe_testcom.out"
   type "%TEMP%\mtpe_testcom.out" >> "%LOGFILE%"
   del "%TEMP%\mtpe_testcom.out" 2>NUL
   if %ERRORS%==0 call :cleanup_legacy
   goto summary

REM ==========================================================================
:unreg
   call :dounreg
   if "%1"=="/unregister" goto summary
   if "%1"=="/u" goto uninstall
   goto summary

:uninstall
   call :log Removing "%TARGET_DIR%" ...
   if exist "%TARGET_DIR%" rmdir /q /s "%TARGET_DIR%" >>"%LOGFILE%" 2>&1
   if exist "%TARGET_DIR%" (set /a ERRORS+=1 & call :log   ERROR - could not remove target folder, files may be in use.) else (call :log   OK - target folder removed.)
   rem Remove the compatibility junction at the old x86 path, if present.
   rem rmdir without /s deletes only a junction or an empty folder - a real
   rem legacy folder with content is deliberately left alone.
   if defined ProgramFiles(x86) rmdir /q "%ProgramFiles(x86)%\Siemens\MT_PropertyExchange" 2>NUL
   goto summary

REM ==========================================================================
REM  Subroutine: unregister everything (current and legacy locations).
REM  Failures here are logged but NOT counted as errors - on a first
REM  installation there is simply nothing to unregister.
REM ==========================================================================
:dounreg
   echo.
   call :log Unregistering COM components of previous installations ...
   >>"%LOGFILE%" echo --- unregister pass ---

   if exist "%TARGET_DIR%\MT_PropertyExchangeDLL.dll" if exist "%REGASM32%" "%REGASM32%" "%TARGET_DIR%\MT_PropertyExchangeDLL.dll" /codebase /tlb:"%TARGET_DIR%\MT_PropertyExchangeDLL.tlb" /nologo /unregister >>"%LOGFILE%" 2>&1
   if "%IS64%"=="NO" goto dounreg_dso
   if exist "%TARGET_DIR%\MT_PropertyExchangeDLL.dll" if exist "%REGASM64%" "%REGASM64%" "%TARGET_DIR%\MT_PropertyExchangeDLL.dll" /codebase /tlb:"%TARGET_DIR%\MT_PropertyExchangeDLL.tlb" /nologo /unregister >>"%LOGFILE%" 2>&1

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
   call :log Unregister pass finished.
   goto :EOF

REM ==========================================================================
REM  Subroutine: after a fully successful installation, remove leftovers of
REM  previous installations: the obsolete folder in the other Program Files
REM  location (all registry entries already point to the new folder) and the
REM  orphaned DllSurrogate/AppID keys of the abandoned x64 registry hack.
REM  Deliberately NOT removed: dsofile.dll copies in Windows\System32 or
REM  SysWOW64 from very old installers - DSOFile was a shared Microsoft
REM  component and other software may still reference those files.
REM ==========================================================================
:cleanup_legacy
   echo.
   call :log Cleaning up leftovers of previous installations ...
   set LEGACY_DIR=
   if defined ProgramFiles(x86) set LEGACY_DIR=%ProgramFiles(x86)%\Siemens\MT_PropertyExchange
   if not defined LEGACY_DIR goto cleanup_reg
   if /I "%LEGACY_DIR%"=="%TARGET_DIR%" goto cleanup_reg
   if not exist "%LEGACY_DIR%" goto cleanup_link
   call :log   Removing obsolete folder "%LEGACY_DIR%" ...
   rmdir /s /q "%LEGACY_DIR%" 2>NUL
   if exist "%LEGACY_DIR%" goto cleanup_locked
   call :log   OK - obsolete folder removed.
:cleanup_link
   rem Compatibility link: any old hard-coded reference to the (x86) path,
   rem e.g. in SAP customizing or scripts, keeps working via a junction
   rem that points to the real installation folder.
   mklink /J "%LEGACY_DIR%" "%TARGET_DIR%" >NUL 2>NUL
   if exist "%LEGACY_DIR%\MT_PropertyExchange.exe" (call :log   OK - compatibility link created: the old x86 path now points to the new folder.) else (call :log   NOTE - compatibility link could not be created; old hard-coded x86 paths would not work.)
   goto cleanup_reg
:cleanup_locked
   call :log   WARNING - could not remove it completely, files may be locked. Delete it manually.
:cleanup_reg
   rem Orphaned keys of the old DllSurrogate workaround, harmless if absent.
   reg delete "HKCR\Wow6432Node\CLSID\{41A13AC0-103B-40EC-9A52-AEE2C6C846C6}" /v AppID /f >NUL 2>NUL
   reg delete "HKCR\CLSID\{41A13AC0-103B-40EC-9A52-AEE2C6C846C6}" /v AppID /f >NUL 2>NUL
   reg delete "HKCR\Wow6432Node\AppID\{41A13AC0-103B-40EC-9A52-AEE2C6C846C6}" /f >NUL 2>NUL
   reg delete "HKLM\Software\Classes\AppID\{41A13AC0-103B-40EC-9A52-AEE2C6C846C6}" /f >NUL 2>NUL
   call :log   OK - obsolete DllSurrogate registry entries removed - if any were present.
   call :log Cleanup finished.
   goto :EOF

REM ==========================================================================
REM  Subroutine: write a line to screen AND log file
REM ==========================================================================
:log
   echo %*
   >>"%LOGFILE%" echo %*
   goto :EOF

REM ==========================================================================
:summary
   echo.
   echo ================================================================
   if %ERRORS%==0 goto sum_ok
   color CF
   call :log RESULT: FAILED - %ERRORS% error/s occurred.
   echo ================================================================
   echo.
   echo -------- full log --------
   type "%LOGFILE%"
   echo --------------------------
   goto sum_end
:sum_ok
   color 2F
   call :log RESULT: SUCCESS - all steps completed without errors.
   echo ================================================================
:sum_end
   echo.
   echo The full log was saved to:
   echo    %LOGFILE%
   echo.
   set /p DUMMY=Press ENTER to close this window ...
   popd 2>NUL

:ende
popd 2>NUL
