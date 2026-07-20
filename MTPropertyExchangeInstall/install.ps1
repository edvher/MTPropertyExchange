<# 
	Hide console window
#>
set-executionpolicy unrestricted
$script:showWindowAsync = Add-Type –memberDefinition @” 
[DllImport("user32.dll")] 
public static extern bool ShowWindowAsync(IntPtr hWnd, int nCmdShow); 
“@ -name “Win32ShowWindowAsync” -namespace Win32Functions –passThru
function Show-PowerShell() { 
	$null = $showWindowAsync::ShowWindowAsync((Get-Process –id $pid).MainWindowHandle, 1) 
}
function Hide-PowerShell() { 
	$null = $showWindowAsync::ShowWindowAsync((Get-Process –id $pid).MainWindowHandle, 2) 
}
#Hide-PowerShell 

<# 
	Self-elevation
#>
$myWindowsID=[System.Security.Principal.WindowsIdentity]::GetCurrent()
$myWindowsPrincipal=new-object System.Security.Principal.WindowsPrincipal($myWindowsID)
$adminRole=[System.Security.Principal.WindowsBuiltInRole]::Administrator
if ($myWindowsPrincipal.IsInRole($adminRole))
{
	$Host.UI.RawUI.WindowTitle = $myInvocation.MyCommand.Definition + “(Elevated)”
	clear-host
}
else {
	$newProcess = new-object System.Diagnostics.ProcessStartInfo “PowerShell”;
	$newProcess.Arguments = $myInvocation.MyCommand.Definition;
	$newProcess.Verb = “runas”;
	[System.Diagnostics.Process]::Start($newProcess);
	exit
}
<# 
	Functions
#>




function Ask ([string]$ask, [string]$title="Question") {
	[System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms") | out-null 
	$result = [System.Windows.Forms.MessageBox]::Show($ask, $title,[System.Windows.Forms.MessageBoxButtons]::YesNo,[System.Windows.Forms.MessageBoxIcon]::Question,[System.Windows.Forms.MessageBoxDefaultButton]::Button1)
	return $result;
}
function Message ([string]$msg) {
	[System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms") | out-null 
	$result = [System.Windows.Forms.MessageBox]::Show($msg, "",[System.Windows.Forms.MessageBoxButtons]::OK,[System.Windows.Forms.MessageBoxIcon]::None)
}

function RegisterDllServer([string]$DllPath, [switch]$Uninstall = $false, [switch]$Verbose = $false, [switch]$Debug = $false)
{
  begin
  {
    $Write = &{if($Verbose -or ($VerbosePreference -ne "SilentlyContinue")) { "Write-Output" } else { "Write-Verbose" }}
  }
  process
  {
    if($_) { $DllPath = $_ }
    
    if(Test-Path $DllPath) {
      $p = new-object System.Diagnostics.Process;
      $si = new-object System.Diagnostics.ProcessStartInfo;

      $si.FileName = "regsvr32.exe";
      $si.Arguments = "`"$DllPath`"";
      if(!$Debug) {
        $si.Arguments = "/s " + $si.Arguments;
      }
      if($Uninstall) {
        $si.Arguments = "/u " + $si.Arguments;
      }
      $p.StartInfo = $si;
      $result = $p.Start();
      $p.WaitForExit();
      [int]$exit = $p.ExitCode; 
      $p.Dispose();

      switch( $exit ) {
        0 { &$Write "`"$DllPath`" Registered Successfully" }
        1 { &$Write "Bad arguments to RegSvr32" }
        2 { &$Write "OLE initilization failed for `"$DllPath`"" }
        3 { &$Write "Failed to load the module `"$DllPath`", you may need to check for problems with dependencies." }
        4 { if($Uninstall) { &$Write "Can't find DllUnregisterServer entry point in the file `"$DllPath`", maybe it's not a .DLL or .OCX?" } else { &$Write "Can't find DllRegisterServer entry point in the file `"$DllPath`", maybe it's not a .DLL or .OCX?" } }
        5 {  &$Write "The assembly `"$DllPath`" was loaded, but the call to DllRegisterServer failed."
            if(!$Debug) {
              &$Write "Call RegisterDllServer again with the -Debug switch to see more information in a MessageBox." 
            }
          }
        default { &$Write "Something went wrong, with Exit Code $exit!" }
      }
      return $exit;  
    } else {
      &$Write "Failed to find the file `"$DllPath`" please check the path."
      return 3;
    }
  }
}
Set-Alias RegSvr32 RegisterDllServer


function RegisterAssembly([string]$DllPath, [switch]$Codebase = $false, [switch]$TypeLib = $false, [switch]$Unregister = $false, [switch]$Verbose = $false)
{
  process
  {
    if($_) { $DllPath = $_ }
    
    if($fso.FileExists($DllPath)) {
      $regasm = ""
  
	 	if($fso.FileExists((Get-Item "Env:Windir").Value+"\Microsoft.NET\Framework\v4.0.30319\regasm.exe") ) {
			$regasm = (Get-Item "Env:Windir").Value + "\Microsoft.NET\Framework\v4.0.30319\";
		}
	 	if($fso.FileExists((Get-Item "Env:Windir").Value+"\Microsoft.NET\Framework\v2.0.50727\regasm.exe") ) {
			$regasm = (Get-Item "Env:Windir").Value + "\Microsoft.NET\Framework\v2.0.50727\";
		}
      $regasm = $regasm + "regasm.exe ";
      
		$args = "`"$DllPath`"";
      if($Unregister) {
        $args = $args + " /unregister";
      }
		if($Codebase) {
        $args = $args + " /codebase";
      }
		if($Typelib) {
        $args = $args + " /tlb:`"" +  [System.IO.Path]::Combine([System.IO.Path]::GetDirectoryName($DllPath), [System.IO.Path]::GetFileNameWithoutExtension($DllPath) + ".tlb") + "`"";
      }
      if(!$Verbose) {
        $args = $args + " /nologo";
      }
		
		$line=$regasm + " " + $args
		cmd /c $line
		
      [int]$exit = $lastExitCode; 

      switch( $exit ) {
        0 {  }
        default { echo "Something went wrong, with Exit Code $exit!" }
      }
      return $exit;  
    } else {
      echo "Failed to find the file `"$DllPath`" please check the path."
      return 3;
    }
  }
}
Set-Alias RegAsm RegisterAssembly



function ReplaceRegistered ($DllPath) {
	if ($fso.FileExists($DllPath) -eq $false) { return $false; }
	if ($fso.FileExists([System.IO.Path]::GetFileName($DllPath)) -eq $false) { return $false; }

	$fi=[System.IO.Path]::GetFileName($DllPath)
	RegSvr32 $DllPath -Uninstall
	cp $fi $DllPath -Force -Verbose
	RegSvr32 $DllPath	
}


<# 
	The script
#>

$dp0 = [System.IO.Path]::GetDirectoryName($myInvocation.MyCommand.Definition)
$ProgramFiles=(Get-Item "Env:ProgramFiles").Value
$Windir=(Get-Item "Env:Windir").Value

$InstallPLM="\\atlnztfile01.ww300.siemens.net\apps600\SAPPLM\Install"
$TargetDir=$ProgramFiles + "\Siemens\MT_PropertyExchange"
$fso = New-Object -ComObject Scripting.FileSystemObject

pushd $dp0
ReplaceRegistered($Windir+"\syswow64\dsofile.dll")
ReplaceRegistered($Windir+"\system32\dsofile.dll")
popd

if ( (Test-Path $TargetDir) -eq $false ) { 
	mkdir $TargetDir | Out-Null 
}
pushd $TargetDir
Regasm ($TargetDir + "MT_PropertyExchangeDLL.dll") -Codebase -Typelib -Unregister
RegSvr32 "dsofile.dll" -Uninstall
popd




<# 
	The end
#>
$result = Ask "Show console output?"
if ($result -eq [System.Windows.Forms.DialogResult]::Yes) {
	Show-PowerShell
	sleep 5
}
exit
