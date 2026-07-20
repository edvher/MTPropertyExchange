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
    if "%1" == "" echo UAC.ShellExecute "%~s0", "", "", "runas", 1 >> "%temp%\getadmin.vbs"
    if "%1" == "/u" echo UAC.ShellExecute "%~s0", "/u", "", "runas", 1 >> "%temp%\getadmin.vbs"
    if "%1" == "/unregister" echo UAC.ShellExecute "%~s0", "/unregister", "", "runas", 1 >> "%temp%\getadmin.vbs"
    if "%1" == "/register" echo UAC.ShellExecute "%~s0", "/register", "", "runas", 1 >> "%temp%\getadmin.vbs"
    rem pause
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
   mode con lines=80 cols=90
   color 8F
   cls
   
   echo.
   if "%1" == "" echo Installation of MT_PropertyExchange toolkit
   if "%1" == "/u" echo Uninstall MT_PropertyExchange toolkit
   if "%1" == "/unregister" echo Unregistering DLLs
   if "%1" == "/register" echo Registering DLLs
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

  if exist "%WINDIR%\Microsoft.NET\Framework\v2.0.50727\regasm.exe" SET REGASM="%WINDIR%\Microsoft.NET\Framework\v2.0.50727\regasm.exe"
  if exist "%WINDIR%\Microsoft.NET\Framework\v4.0.30319\regasm.exe" SET REGASM="%WINDIR%\Microsoft.NET\Framework\v4.0.30319\regasm.exe"

   set CHOICE=%INSTALL_ROOT%\lib\choice.exe  
   set CHOICEPARAM=/C:yn /T:n,5 
   if exist %WINDIR%\system32\choice.exe (
      set CHOICE=%WINDIR%\system32\choice.exe
      set CHOICEPARAM=/C yn /T 5 /D n /M
   )

   if "%1"=="/u" goto switch
   if "%1"=="/unregister" goto switch
   if "%1"=="/register" goto switch

REM    REM ==========================================================================================
REM    echo Copy and register system dlls ...
REM    SET DLL_TO_KILL=dsofile.dll
REM    SET KILLER=
REM :askkill
REM    for /f "tokens=1-3 delims= " %%a in ('tasklist /M "%DLL_TO_KILL%*" /FO TABLE /NH 2^>NUL') do ( 
REM      IF NOT DEFINED KILLER (
REM        SET /P KILLER=%%a is locking %%c. Shall the process be killed [y/n]?
REM      )
REM      IF "!KILLER!"=="y" (
REM        echo killing %%a...
REM        TASKKILL /T /F /PID %%b 
REM        goto askkill
REM      )
REM    ) 
REM    SET KILLER=
REM    for /f "tokens=1-3 delims= " %%a in ('tasklist /M "%DLL_TO_KILL%*" /FO TABLE /NH 2^>NUL') do (
REM      COLOR C0
REM      echo.
REM      echo +================================================================================+
REM      echo |                                                                                | 
REM      echo |  Error: a process is still locking dsofile.dll. Installation cannot continue.  | 
REM      echo |                                                                                | 
REM      echo +================================================================================+
REM      echo.
REM      pause
REM      exit
REM    )
   
   if exist "%WINDIR%\syswow64\dsofile.dll" regsvr32 /s /u "%WINDIR%\syswow64\dsofile.dll"
   if exist "%WINDIR%\syswow64\dsofile.dll" copy "%SERVER_DLLs%\dsofile.dll" "%WINDIR%\syswow64\dsofile.dll" /y >nul
   if exist "%WINDIR%\syswow64\dsofile.dll" regsvr32 /s "%WINDIR%\syswow64\dsofile.dll"
   if exist "%WINDIR%\syswow64\dsofile.dll" goto switch

   if exist "%WINDIR%\system32\dsofile.dll" regsvr32 /s /u "%WINDIR%\system32\dsofile.dll"
   if exist "%WINDIR%\system32\dsofile.dll" copy "%SERVER_DLLs%\dsofile.dll" "%WINDIR%\system32\dsofile.dll" /y >nul
   if exist "%WINDIR%\system32\dsofile.dll" regsvr32 /s "%windir%\system32\dsofile.dll"
   if exist "%WINDIR%\system32\dsofile.dll" goto switch

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
   if "%1"=="/unregister" goto unreg
   if "%1"=="/register" goto reg
   mkdir "%TARGET_DIR%" 2>NUL

:unreg
   echo.
   echo Unregistering type libraries ...
   pushd "%TARGET_DIR%"
   if exist MT_PropertyExchangeDLL.dll echo Unregistering MT_PropertyExchangeDLL.dll ...
   if exist MT_PropertyExchangeDLL.dll %REGASM% MT_PropertyExchangeDLL.dll /codebase /tlb:MT_PropertyExchangeDLL.tlb /nologo /unregister
   if exist dsofile.dll echo Unregistering dsofile.dll ...
   if exist dsofile.dll regsvr32 /s /u dsofile.dll
   popd
   if "%1"=="/unregister" goto fine
   if "%1"=="/u" goto uninstall

:copy
   echo.
   echo Copying toolkit to %TARGET_DIR% ...
   xcopy /Z /C /S /Y /V "%SERVER_DLLs%\*.*" "%TARGET_DIR%"
   echo.
   echo.
   echo Checkup: Folder contents
   for %%a in ("%TARGET_DIR%\*.*") do @echo    %%~nxa
   echo.
   echo Copying finished.

:reg
   echo.
   echo Registering type libraries ...
   pushd "%TARGET_DIR%"
   if exist dsofile.dll echo Register DSOFILE.DLL ...
   if exist dsofile.dll regsvr32 /s dsofile.dll
   echo.
   if exist MT_PropertyExchangeDLL.dll echo Create MT_PropertyExchangeDLL.dll registry file ...
   if exist MT_PropertyExchangeDLL.dll %REGASM% MT_PropertyExchangeDLL.dll /codebase /regfile:MT_PropertyExchangeDLL_%PROCESSOR_ARCHITECTURE%_%COMPUTERNAME%.reg /nologo
   echo.
   if exist MT_PropertyExchangeDLL.dll echo Register MT_PropertyExchangeDLL.dll ...
   if exist MT_PropertyExchangeDLL.dll %REGASM% MT_PropertyExchangeDLL.dll /codebase /tlb:MT_PropertyExchangeDLL.tlb /nologo
   echo.
   popd

   SET EXE="%TARGET_DIR%\MT_PropertyExchange.exe"
   if "%1"=="/register" goto fine

REM ==========================================================================================
:log
   2>NUL call %INSTALL_PLM%\Write_log.bat MT_PropertyTransfer Installation

REM ==========================================================================================
:test
   echo.
   echo Quick test: checking version.
   mkdir "%TARGET_DIR%\Test" 2>NUL
   xcopy /q /S /Y /V "%SERVER_TESTS%\*.*" "%TARGET_DIR%\Test\" 1>nul 2>nul
   call "%TARGET_DIR%\Test\test-com.bat"
   IF "%ERRORLEVEL%"=="0" COLOR 2F
   echo.
   echo End of test.
   goto fine

:uninstall
   Echo Removing "%TARGET_DIR%" ...
   if exist "%TARGET_DIR%" rmdir /q /s "%TARGET_DIR%" 

:fine
   echo.
   echo Done. This window will close in a few seconds...
   %CHOICE% %CHOICEPARAM% "..." >NUL 2>NUL
   popd 2>NUL

:ende
popd 2>NUL
if exist "%temp%\MTPropertyExchangeInstall" (
   pushd %temp% 2>NUL
   xcopy /s /q "%temp%\MTPropertyExchangeInstall" "%temp%\MTPropertyExchangeInstall-%date%-%hh%-%nn%-%ss%\" 1>NUL 2>NUL
   popd 2>NUL
   rmdir /Q /S "%temp%\MTPropertyExchangeInstall" 1>NUL 2>NUL
)

:endoffile
