@echo off
:: Verifies that the COM components can be created. On 64-bit Windows the
:: check runs twice: once with the 64-bit script host (as the 64-bit SAP
:: Business Client would) and once with the 32-bit script host (as legacy
:: 32-bit callers would). Both must succeed.

echo Dim obj>%TEMP%\test-com.vbs
echo Set obj = WScript.CreateObject("DSOFile.OleDocumentProperties")>>%TEMP%\test-com.vbs
echo If Not IsObject(obj) Then>>%TEMP%\test-com.vbs
echo 	Wscript.Echo("COM use of DSOFile.OleDocumentProperties failed.")>>%TEMP%\test-com.vbs
echo 	WScript.Quit 1>>%TEMP%\test-com.vbs
echo Else>>%TEMP%\test-com.vbs
echo 	Wscript.Echo("COM use of DSOFile.OleDocumentProperties ok.")>>%TEMP%\test-com.vbs
echo End If>>%TEMP%\test-com.vbs
echo Set obj = Nothing>>%TEMP%\test-com.vbs
echo Set obj = WScript.CreateObject("MT_PropertyExchange.Globals")>>%TEMP%\test-com.vbs
echo If Not IsObject(obj) Then>>%TEMP%\test-com.vbs
echo 	Wscript.Echo("COM use of MT_PropertyExchange.Globals failed.")>>%TEMP%\test-com.vbs
echo 	WScript.Quit 1>>%TEMP%\test-com.vbs
echo Else>>%TEMP%\test-com.vbs
echo 	Wscript.Echo("COM use of MT_PropertyExchange.Globals ok.")>>%TEMP%\test-com.vbs
echo 	Wscript.Echo("DLL Version: " ^& obj.Version)>>%TEMP%\test-com.vbs
echo End If>>%TEMP%\test-com.vbs
echo Set obj = Nothing>>%TEMP%\test-com.vbs
echo WScript.Quit 0>>%TEMP%\test-com.vbs

set RC=0

if NOT DEFINED ProgramFiles(x86) goto only32

echo.
echo --- 64-bit COM check (as used by 64-bit SAP Business Client) ---
%windir%\system32\cscript.exe /nologo %TEMP%\test-com.vbs
if errorlevel 1 set RC=1

echo.
echo --- 32-bit COM check (as used by legacy 32-bit callers) ---
%windir%\syswow64\cscript.exe /nologo %TEMP%\test-com.vbs
if errorlevel 1 set RC=1
goto done

:only32
echo.
echo --- 32-bit COM check ---
%windir%\system32\cscript.exe /nologo %TEMP%\test-com.vbs
if errorlevel 1 set RC=1

:done
exit /b %RC%
