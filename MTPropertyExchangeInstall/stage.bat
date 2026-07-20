@echo off
:: ==========================================================================
::  Collects the build output into the installer payload layout used by
::  install.bat. Run this AFTER building:
::    1. Dsofile\dsofile.sln            Release|Win32  and  Release|x64
::    2. MsoPropertyTransferUtils3.5.sln  Release|Any CPU
:: ==========================================================================
setlocal
pushd "%~dp0"

set SRC_MANAGED=..\MT_PropertyExchange.exe\bin\Any\Release
set SRC_DSO_X86=..\Dsofile\bin\Win32\Release
set SRC_DSO_X64=..\Dsofile\bin\x64\Release

set RC=0

if not exist "%SRC_MANAGED%\PropertyExchange.exe" (
   echo ERROR: %SRC_MANAGED%\PropertyExchange.exe not found. Build MsoPropertyTransferUtils3.5.sln Release^|Any CPU first.
   set RC=1
)
if not exist "%SRC_DSO_X86%\dsofile.dll" (
   echo ERROR: %SRC_DSO_X86%\dsofile.dll not found. Build Dsofile\dsofile.sln Release^|x86 first.
   set RC=1
)
if not exist "%SRC_DSO_X64%\dsofile.dll" (
   echo ERROR: %SRC_DSO_X64%\dsofile.dll not found. Build Dsofile\dsofile.sln Release^|x64 first.
   set RC=1
)
if not "%RC%"=="0" goto ende

echo Staging managed binaries ...
mkdir bin 2>NUL
xcopy /Y /C /Q "%SRC_MANAGED%\*.*" bin\

echo Staging native dsofile.dll (x86 and x64) ...
mkdir x86 2>NUL
mkdir x64 2>NUL
copy /Y "%SRC_DSO_X86%\dsofile.dll" x86\ >NUL
if exist "%SRC_DSO_X86%\dsofile.pdb" copy /Y "%SRC_DSO_X86%\dsofile.pdb" x86\ >NUL
copy /Y "%SRC_DSO_X64%\dsofile.dll" x64\ >NUL
if exist "%SRC_DSO_X64%\dsofile.pdb" copy /Y "%SRC_DSO_X64%\dsofile.pdb" x64\ >NUL

echo.
echo Staging finished. This folder can now be copied to a target machine;
echo run install.bat there to install and register the toolkit.

:ende
popd
endlocal & exit /b %RC%
