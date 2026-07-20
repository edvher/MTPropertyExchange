@echo off 
cls

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
    rem if exist "%temp%\mtpe.run" goto gotAdmin

    echo Set UAC = CreateObject^("Shell.Application"^) > "%temp%\getadmin.vbs"
    if "%1" == "" echo UAC.ShellExecute "%~s0", "/unregister", "", "runas", 1 >> "%temp%\getadmin.vbs"
    if "%1" == "/unregister" echo UAC.ShellExecute "%~s0", "/unregister", "", "runas", 1 >> "%temp%\getadmin.vbs"
    if "%1" == "/register" echo UAC.ShellExecute "%~s0", "/register", "", "runas", 1 >> "%temp%\getadmin.vbs"
    if "%1" == "/reg" echo UAC.ShellExecute "%~s0", "/register", "", "runas", 1 >> "%temp%\getadmin.vbs"
    "%temp%\getadmin.vbs"
    exit /B


:gotAdmin
    if exist "%temp%\getadmin.vbs" ( del "%temp%\getadmin.vbs" )
    pushd "%CD%"
    CD /D "%~dp0"
	
:--------------------------------------
:start
   @echo off
   pushd "%~dp0"
   set root=%CD%
   mode con lines=40 cols=90
   color C0
   cls
   echo.
   
   rem ******************************************************************************************
   rem *    Allgemeine Parameter                                                                *
   rem ******************************************************************************************
   set INSTALL_PLM=\\atlnztfile01.ww300.siemens.net\apps600\SAPPLM\Install
   set INSTALL_Root=%root%

   set ARCH=x86
   set SERVER_Vorlagen=\\ww300.siemens.net\DFSroot\info\OFFICE2K\VAI Projekte
   set SERVER_DLLs=%INSTALL_Root%\%ARCH%
   set SERVER_TESTs=%INSTALL_Root%\test
   for /f "tokens=5-8 delims=:., " %%a in ('echo/^|time') do ( set hh=%%a&set nn=%%b&set ss=%%c&set cs=%%d)

   if exist "%WINDIR%\Microsoft.NET\Framework\v4.0.30319\regasm.exe" SET REGASM="%WINDIR%\Microsoft.NET\Framework\v4.0.30319\regasm.exe"
   if exist "%WINDIR%\Microsoft.NET\Framework\v2.0.50727\regasm.exe" SET REGASM="%WINDIR%\Microsoft.NET\Framework\v2.0.50727\regasm.exe"
 

REM ==========================================================================================
:switch
   if "%ARCH%"=="x86" goto inst_x86
   if "%ARCH%"=="x64" goto inst_x64

:inst_x86
   set TARGET_DIR=%ProgramFiles%\Siemens\MT_PropertyExchange
   goto common

:inst_x64
   set TARGET_DIR=%ProgramFiles(x86)%\Siemens\MT_PropertyExchange
   goto common

REM ==========================================================================================
:common
   if "%1"=="/register" goto reg
   if "%1"=="/reg" goto reg

:unreg
   echo Unregistering type libraries ...
   pushd "%TARGET_DIR%"
   if exist MT_PropertyExchangeDLL.dll %REGASM% MT_PropertyExchangeDLL.dll /codebase /tlb:MT_PropertyExchangeDLL.tlb /nologo /unregister
   if exist dsofile.dll regsvr32 /s /u dsofile.dll
   popd
  goto fine

:reg
   echo.
   echo Registering type libraries ...
   pushd "%TARGET_DIR%"
   if exist dsofile.dll regsvr32 /s dsofile.dll
   if exist MT_PropertyExchangeDLL.dll %REGASM% MT_PropertyExchangeDLL.dll /codebase /tlb:MT_PropertyExchangeDLL.tlb /nologo
   popd
   goto fine

:fine
   echo.
   echo Done. 

:ende
popd

