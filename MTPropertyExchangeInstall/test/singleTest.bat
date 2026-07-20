@echo off
set root=%~dp0

rem echo %root%
rem pause

SET FI=%1
SET PAUSE=NO
IF "%FI%"=="" set PAUSE=YES
IF "%FI%"=="" set FI="%root%\SAP.xls"


if exist "%ProgramFiles%\Siemens\MT_PropertyExchange" set TARGET_DIR=%ProgramFiles%\Siemens\MT_PropertyExchange
if exist "%ProgramFiles(x86)%\Siemens\MT_PropertyExchange" set TARGET_DIR=%ProgramFiles(x86)%\Siemens\MT_PropertyExchange
goto common


REM ==========================================================================================
:common


for /f "tokens=5-8 delims=:., " %%a in ('echo/^|time') do ( set hh=%%a&set nn=%%b&set ss=%%c&set cs=%%d)

pushd "%root%"
path %path%;%TARGET_DIR%;%root%

SET EXE=MT_PropertyExchange.exe -loglevel=DEBUG -logfile=%TEMP%\PropertyExchange.log

@echo off
@echo ________________________________________________________________________________
@echo.
@echo Starting tests with %FI% ....
@echo.
@echo ________________________________________________________________________________
%EXE% -clearlog
@echo ________________________________________________________________________________
@echo Export to text file...

@for /f "tokens=1-5 delims=," %%f IN ('%EXE% -file:%FI% -mode:g') DO type "%%f"
@echo Errorlevel=%ERRORLEVEL%
@echo ________________________________________________________________________________
@echo Read Title2 ...
%EXE% -file=%FI% -mode=r -prop="title2"
@echo Errorlevel=%ERRORLEVEL%
@echo ________________________________________________________________________________
@echo Write Title2 ...
%EXE% -file=%FI% -mode=w -prop="title2" -value="%date% %hh%:%nn%:%ss%.%cs%"
@echo Errorlevel=%ERRORLEVEL%
@echo ________________________________________________________________________________
@echo Read Title2 again ...
%EXE% -file=%FI% -mode=r -prop="title2"
@echo Errorlevel=%ERRORLEVEL%
@echo ________________________________________________________________________________
@echo Read all ...
%EXE% -file=%FI% -mode=r -prop=*
@echo Errorlevel=%ERRORLEVEL%
@echo ________________________________________________________________________________
@echo Export Properties to xml ...
@for /f "tokens=1-5 delims=," %%f IN ('%EXE% -file:%FI% -mode:e') DO type "%%f"
@echo.
@echo ________________________________________________________________________________
@echo Finished tests with %FI% ....
@echo ________________________________________________________________________________
@echo.
@echo In case of errors, send the file "%TEMP%\PropertyExchange.log" to your support.
@echo.
@echo ________________________________________________________________________________
@echo.
@echo off

popd

:ende
if "%PAUSE%"=="YES" pause
