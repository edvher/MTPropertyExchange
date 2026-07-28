@echo off
:: ==========================================================================
::  Builds ONE self-extracting installer EXE from the staged payload:
::     Install\MTPropertyExchangeInstall-<timestamp>.exe
::
::  Prerequisites (see 7-zip\README.md):
::    - the solution and Dsofile are built (Release) and stage.bat has run
::      (this script runs stage.bat itself)
::    - 7-Zip installed (or 7z.exe placed in the 7-zip\ folder)
::    - 7zSD.sfx from the "7-Zip Extra" package placed in the 7-zip\ folder
:: ==========================================================================
setlocal
pushd "%~dp0"
set RC=0

echo.
echo === Step 1/3: staging installer payload =============================
call MTPropertyExchangeInstall\stage.bat
if errorlevel 1 (
   echo ERROR: staging failed - build the solutions first, see README.md.
   set RC=1
   goto ende
)

echo.
echo === Step 2/3: locating 7-Zip ========================================
set SEVENZIP=
if exist "%ProgramFiles%\7-Zip\7z.exe" set SEVENZIP=%ProgramFiles%\7-Zip\7z.exe
if defined SEVENZIP goto have7z
if not defined ProgramFiles(x86) goto try_local
if exist "%ProgramFiles(x86)%\7-Zip\7z.exe" set SEVENZIP=%ProgramFiles(x86)%\7-Zip\7z.exe
if defined SEVENZIP goto have7z
:try_local
if exist "7-zip\7z.exe" set SEVENZIP=7-zip\7z.exe
if defined SEVENZIP goto have7z
for %%i in (7z.exe) do if not "%%~$PATH:i"=="" set SEVENZIP=%%~$PATH:i
if defined SEVENZIP goto have7z
echo ERROR: 7z.exe not found. Install 7-Zip or copy 7z.exe+7z.dll into 7-zip\.
set RC=1
goto ende
:have7z
echo Using: %SEVENZIP%

if exist "7-zip\7zSD.sfx" goto havesfx
echo ERROR: 7-zip\7zSD.sfx not found.
echo Download the "7-Zip Extra" package from https://www.7-zip.org/download.html
echo and copy 7zSD.sfx into the 7-zip\ folder. See 7-zip\README.md.
set RC=1
goto ende
:havesfx

echo.
echo === Step 3/3: packing self-extracting installer =====================

for /f %%i in ('powershell -NoProfile -Command "Get-Date -Format yyyy-MM-dd_HH-mm-ss"') do set TS=%%i
if "%TS%"=="" set TS=unknown-date
set GITREV=
for /f %%i in ('git rev-parse --short HEAD 2^>NUL') do set GITREV=%%i
if not defined GITREV set GITREV=norev

rem Stamp the payload so install.bat can log exactly which build it is.
> MTPropertyExchangeInstall\build-info.txt echo %GITREV% %TS%
echo Build stamp: %GITREV% %TS%

if exist MTPropertyExchangeInstall.7z del MTPropertyExchangeInstall.7z
rem Pack the payload CONTENTS at the archive root (no wrapping folder):
rem 7zSD.sfx runs RunProgram relative to the extraction root, so
rem install.bat must sit directly at the top of the archive.
set SEVENZIP_ABS=%SEVENZIP%
if exist "%~dp0%SEVENZIP%" set SEVENZIP_ABS=%~dp0%SEVENZIP%
pushd MTPropertyExchangeInstall
"%SEVENZIP_ABS%" a -mx=9 "%~dp0MTPropertyExchangeInstall.7z" *
if errorlevel 1 (
   popd
   echo ERROR: packing the payload failed.
   set RC=1
   goto ende
)
popd

mkdir Install 2>NUL
copy /b "7-zip\7zSD.sfx" + "7-zip\MTPE_sfx_config.txt" + MTPropertyExchangeInstall.7z "Install\MTPropertyExchangeInstall-%TS%-%GITREV%.exe" >NUL
if errorlevel 1 (
   echo ERROR: could not assemble the SFX exe.
   set RC=1
   goto ende
)
del MTPropertyExchangeInstall.7z

echo.
echo SUCCESS: Install\MTPropertyExchangeInstall-%TS%-%GITREV%.exe
echo This single file is the complete installer - copy it to a target
echo machine and double-click it.

:ende
popd
endlocal & exit /b %RC%
