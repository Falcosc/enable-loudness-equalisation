<#
.SYNOPSIS
    automatically add and enables loudness equalisation to any playback device
.DESCRIPTION
    Imports registry keys to add enhancement features and enables loudness equalisation.
    It restarts audio service to apply imported registry settings
.NOTES
    https://github.com/Falcosc/enable-loudness-equalisation
#>

Param(
   [Parameter(Mandatory,HelpMessage='Which Playback Device Name should be configured?')]
   [ValidateLength(3,50)]
   [string]$playbackDeviceName,
   
   [int]$maxDeviceCount=2
)

Add-Type -AssemblyName System.Windows.Forms
function exitWithErrorMsg ([String] $msg){
    Write-Error $msg
    [void][System.Windows.Forms.MessageBox]::Show($msg, $PSCommandPath,
    [System.Windows.Forms.MessageBoxButtons]::OK,
    [System.Windows.Forms.MessageBoxIcon]::Error)
    exit 1
}

$regFile = "$env:temp\SoundEnhancementsTMP.reg"
$enhancementFlagKey = "{fc52a749-4be9-4510-896e-966ba6525980},3"
$enableLoudness = @'
"{d04e05a6-594b-4fb6-a80d-01af5eed7d1d},1"="{62dc1a93-ae24-464c-a43e-452f824c4250}" ;PreMixEffectClsid activates effects
"{d04e05a6-594b-4fb6-a80d-01af5eed7d1d},2"="{637c490d-eee3-4c0a-973f-371958802da2}" ;PostMixEffectClsid activates effects
"{d04e05a6-594b-4fb6-a80d-01af5eed7d1d},3"="{5860E1C5-F95C-4a7a-8EC8-8AEF24F379A1}" ;UserInterfaceClsid shows it in ui
"{fc52a749-4be9-4510-896e-966ba6525980},3"=hex:0b,00,00,00,01,00,00,00,ff,ff,00,00 ;enables loudness equalisation

'@
$missingLoudness = $false

"Windows Registry Editor Version 5.00" > $regFile
$devices = reg query HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\MMDevices\Audio\Render /s /f $playbackDeviceName /d
if(!$?) {
    exitWithErrorMsg "Could not find any device named $playbackDeviceName"
}

$renderer = $devices | Select-String "Render"
$activeRenderer = $renderer | ? { (Get-ItemPropertyValue -Path Registry::$($_ -replace '\\Properties','') -Name DeviceState) -eq 1}
if($activeRenderer.length -lt 1) {
    exitWithErrorMsg "There are $($renderer.length) devices with Name $playbackDeviceName, but non of them is active"
}
if($activeRenderer.length -gt $maxDeviceCount) {
    $devices
    exitWithErrorMsg "Execution aborted, because more then $maxDeviceCount Active Devices found by Name $playbackDeviceName"
}
$renderer | ForEach-Object{
    $fxKeyPath = $_ -replace 'Properties','FxProperties'
    $fxProperties = Get-ItemProperty -Path Registry::$fxKeyPath
    If (($fxProperties -eq $null) -or ($fxProperties.$enhancementFlagKey -eq $null) -or 
        ($fxProperties.$enhancementFlagKey[8] -ne 255) -or ($fxProperties.$enhancementFlagKey[9] -ne 255)) { 
        "[" + $fxKeyPath + "]" >> $regFile
        $enableLoudness >> $regFile
        $missingLoudness = $true
    }
}

If (!$missingLoudness) {
    "Loudness Settings don't need to be enabled"
    Start-Sleep -Seconds 5
    exit 0
}

$currentPrincipal = New-Object Security.Principal.WindowsPrincipal([Security.Principal.WindowsIdentity]::GetCurrent())
if (-not $currentPrincipal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    $arguments = "-File `"$($myInvocation.MyCommand.Definition)`" -playbackDeviceName $playbackDeviceName  -maxDeviceCount $maxDeviceCount"
    Start-Process powershell -Verb runAs -ArgumentList $arguments
    exit
}

"import loudness activation into registry"
$startprocessParams = @{
    FilePath     = "$Env:SystemRoot\REGEDIT.exe"
    ArgumentList = '/s', $regFile
    Verb         = 'RunAs'
    PassThru     = $true
    Wait         = $true
}
$proc = Start-Process @startprocessParams
If($? -eq $false -or $proc.ExitCode -ne 0) {
    exitWithErrorMsg "Failed to import $regFile"
}

"Restart Audio to apply registry settings"
Restart-Service audiosrv -Force


