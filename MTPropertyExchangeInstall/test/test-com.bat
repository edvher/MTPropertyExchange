@echo off
:: Verifies that the COM components can be created. On 64-bit Windows the
:: check runs twice: once with the 64-bit script host (as the 64-bit SAP
:: Business Client would) and once with the 32-bit script host (as legacy
:: 32-bit callers would). Both must succeed. Exit code 0 = all passed.
::
:: Note: every generated line uses the ">>file echo text" prefix form on
:: purpose - "echo text 1>>file" would swallow a trailing digit as a file
:: handle redirect and corrupt the generated script.

set VBS=%TEMP%\mtpe-test-com.vbs

 >"%VBS%" echo Dim obj, rc
>>"%VBS%" echo rc = 0
>>"%VBS%" echo On Error Resume Next
>>"%VBS%" echo Set obj = CreateObject("DSOFile.OleDocumentProperties")
>>"%VBS%" echo If Err.Number ^<^> 0 Then
>>"%VBS%" echo     WScript.Echo "COM use of DSOFile.OleDocumentProperties FAILED: 0x" ^& Hex(Err.Number) ^& " " ^& Err.Description
>>"%VBS%" echo     rc = 1
>>"%VBS%" echo Else
>>"%VBS%" echo     WScript.Echo "COM use of DSOFile.OleDocumentProperties ok."
>>"%VBS%" echo End If
>>"%VBS%" echo Err.Clear
>>"%VBS%" echo Set obj = Nothing
>>"%VBS%" echo Set obj = CreateObject("MT_PropertyExchange.Globals")
>>"%VBS%" echo If Err.Number ^<^> 0 Then
>>"%VBS%" echo     WScript.Echo "COM use of MT_PropertyExchange.Globals FAILED: 0x" ^& Hex(Err.Number) ^& " " ^& Err.Description
>>"%VBS%" echo     rc = 1
>>"%VBS%" echo Else
>>"%VBS%" echo     WScript.Echo "COM use of MT_PropertyExchange.Globals ok."
>>"%VBS%" echo     WScript.Echo "DLL Version: " ^& obj.Version()
>>"%VBS%" echo End If
>>"%VBS%" echo Set obj = Nothing
>>"%VBS%" echo WScript.Quit rc

:: Pick the REAL 64-bit and 32-bit script hosts regardless of our own
:: bitness: from a 32-bit process, System32 is redirected to SysWOW64 and
:: the true 64-bit binaries are reachable only via the "sysnative" alias.
set CSCRIPT64=%windir%\System32\cscript.exe
if exist %windir%\sysnative\cscript.exe set CSCRIPT64=%windir%\sysnative\cscript.exe
set CSCRIPT32=%windir%\SysWOW64\cscript.exe

set RC=0

if NOT DEFINED ProgramFiles(x86) goto only32

echo.
echo --- 64-bit COM check (as used by 64-bit SAP Business Client) ---
"%CSCRIPT64%" /nologo "%VBS%"
if errorlevel 1 set RC=1

echo.
echo --- 32-bit COM check (as used by legacy 32-bit callers) ---
"%CSCRIPT32%" /nologo "%VBS%"
if errorlevel 1 set RC=1
goto done

:only32
echo.
echo --- 32-bit COM check ---
%windir%\System32\cscript.exe /nologo "%VBS%"
if errorlevel 1 set RC=1

:done
exit /b %RC%
