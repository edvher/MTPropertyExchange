@echo off
:: Verifies that the COM components can be created via real registry-based
:: COM activation, in the 64-bit and the 32-bit view. Exit code 0 = all ok.
::
:: Host is PowerShell, not cscript: 64-bit cscript's base directory is
:: System32, which trips log4net's config probing on hardened machines and
:: produced false alarms that real hosts (SAP, Office) never see. PowerShell
:: also reports the FULL inner exception chain on failure.

set PS1=%TEMP%\mtpe-test-com.ps1

 >"%PS1%" echo $rc = 0
>>"%PS1%" echo try {
>>"%PS1%" echo     $d = New-Object -ComObject 'DSOFile.OleDocumentProperties'
>>"%PS1%" echo     Write-Output 'COM use of DSOFile.OleDocumentProperties ok.'
>>"%PS1%" echo } catch {
>>"%PS1%" echo     Write-Output 'COM use of DSOFile.OleDocumentProperties FAILED:'
>>"%PS1%" echo     $e = $_.Exception
>>"%PS1%" echo     while ($e) { Write-Output ('   ' + $e.GetType().FullName + ' :: ' + $e.Message); $e = $e.InnerException }
>>"%PS1%" echo     $rc = 1
>>"%PS1%" echo }
>>"%PS1%" echo try {
>>"%PS1%" echo     $o = New-Object -ComObject 'MT_PropertyExchange.Globals'
>>"%PS1%" echo     Write-Output 'COM use of MT_PropertyExchange.Globals ok.'
>>"%PS1%" echo     Write-Output ('DLL Version: ' + $o.Version())
>>"%PS1%" echo     Write-Output ('GetVersion : ' + $o.GetVersion())
>>"%PS1%" echo } catch {
>>"%PS1%" echo     Write-Output 'COM use of MT_PropertyExchange.Globals FAILED:'
>>"%PS1%" echo     $e = $_.Exception
>>"%PS1%" echo     while ($e) { Write-Output ('   ' + $e.GetType().FullName + ' :: ' + $e.Message); $e = $e.InnerException }
>>"%PS1%" echo     $rc = 1
>>"%PS1%" echo }
>>"%PS1%" echo exit $rc

:: Pick the REAL 64-bit and 32-bit hosts regardless of our own bitness:
:: from a 32-bit process System32 is redirected, the true 64-bit binaries
:: are only reachable via the "sysnative" alias.
set PS64=%windir%\System32\WindowsPowerShell\v1.0\powershell.exe
if exist %windir%\sysnative\WindowsPowerShell\v1.0\powershell.exe set PS64=%windir%\sysnative\WindowsPowerShell\v1.0\powershell.exe
set PS32=%windir%\SysWOW64\WindowsPowerShell\v1.0\powershell.exe

set RC=0

if NOT DEFINED ProgramFiles(x86) goto only32

echo.
echo --- 64-bit COM check (as used by 64-bit SAP Business Client) ---
"%PS64%" -NoProfile -ExecutionPolicy Bypass -File "%PS1%"
if errorlevel 1 set RC=1

echo.
echo --- 32-bit COM check (as used by legacy 32-bit callers) ---
"%PS32%" -NoProfile -ExecutionPolicy Bypass -File "%PS1%"
if errorlevel 1 set RC=1
goto done

:only32
echo.
echo --- 32-bit COM check ---
"%windir%\System32\WindowsPowerShell\v1.0\powershell.exe" -NoProfile -ExecutionPolicy Bypass -File "%PS1%"
if errorlevel 1 set RC=1

:done
exit /b %RC%
