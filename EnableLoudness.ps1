<#
.SYNOPSIS
    automatically add and enables loudness equalisation to any playback device
.DESCRIPTION
    Imports registry keys to add enhancement features and enables loudness equalisation.
    It restarts audio service to apply imported registry settings
.LINK
    https://github.com/Falcosc/enable-loudness-equalisation
.LINK
.PARAMETER playbackDeviceName
    Searches for Audio Device Names starting with this String
.PARAMETER maxDeviceCount
    Limits the amount of devices to be configured
.PARAMETER releaseTime
    time until audio level is adjusted: from "2" (fast adjustment) up to "7" (slow adjustment)
.EXAMPLE
    PS> .\EnableLoudness.ps1 -playbackDeviceName BE279
    enable loudness equalisation for Audio Devie BE279
.EXAMPLE
    PS> .\EnableLoudness.ps1 -releaseTime 2
    set shortest possible time until audio level is adjusted
#>

Param(
   [Parameter(Mandatory,HelpMessage='Which Playback Device Name should be configured?')]
   [ValidateLength(3,50)]
   [string]$playbackDeviceName,
   
   [ValidateRange(1, 10)]
   [int]$maxDeviceCount=2,

   [ValidateRange(2, 7)]
   [int]$releaseTime=4
)

Add-Type -AssemblyName System.Windows.Forms
function exitWithErrorMsg ([String] $msg){
    [void][System.Windows.Forms.MessageBox]::Show($msg, $PSCommandPath,
    [System.Windows.Forms.MessageBoxButtons]::OK,
    [System.Windows.Forms.MessageBoxIcon]::Error)
    Write-Error $msg
    exit 1
}
function importReg ([String] $file){
    $startprocessParams = @{
        FilePath     = "$Env:SystemRoot\REGEDIT.exe"
        ArgumentList = '/s', $file
        Verb         = 'RunAs'
        PassThru     = $true
        Wait         = $true
    }
    $proc = Start-Process @startprocessParams
    If($? -eq $false -or $proc.ExitCode -ne 0) {
        exitWithErrorMsg "Failed to import $file"
    }
}

$ErrorActionPreference = "Stop"
$PSDefaultParameterValues['Out-File:Encoding'] = 'utf8'
$regFile = "$env:temp\SoundEnhancementsTMP.reg"
$enhancementFlagKey = "{fc52a749-4be9-4510-896e-966ba6525980},3"
$releaseTimeKey = "{9c00eeed-edce-4cd8-ae08-cb05e8ef57a0},3"
$enhancementTabKey = "{d04e05a6-594b-4fb6-a80d-01af5eed7d1d},3"
$enhancementTabValue = "{5860E1C5-F95C-4a7a-8EC8-8AEF24F379A1}"
$releaseTimeStr = $releaseTime.ToString().PadLeft(2,'0')
$fxPropertiesImport = @"
"{d04e05a6-594b-4fb6-a80d-01af5eed7d1d},1"="{62dc1a93-ae24-464c-a43e-452f824c4250}" ;PreMixEffectClsid activates effects
"{d04e05a6-594b-4fb6-a80d-01af5eed7d1d},2"="{637c490d-eee3-4c0a-973f-371958802da2}" ;PostMixEffectClsid activates effects
"{d04e05a6-594b-4fb6-a80d-01af5eed7d1d},3"="{5860E1C5-F95C-4a7a-8EC8-8AEF24F379A1}" ;UserInterfaceClsid shows it in ui
"{d04e05a6-594b-4fb6-a80d-01af5eed7d1d},5"="{62dc1a93-ae24-464c-a43e-452f824c4250}" ;StreamEffectClsid
"{d04e05a6-594b-4fb6-a80d-01af5eed7d1d},6"="{637c490d-eee3-4c0a-973f-371958802da2}" ;ModeEffectClsid
"{fc52a749-4be9-4510-896e-966ba6525980},3"=hex:0b,00,00,00,01,00,00,00,ff,ff,00,00 ;enables loudness equalisation
"{9c00eeed-edce-4cd8-ae08-cb05e8ef57a0},3"=hex:03,00,00,00,01,00,00,00,$releaseTimeStr,00,00,00 ;equalisation release time 2 to 7
"@

$devices = Get-ChildItem "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\MMDevices\Audio\Render\*\Properties"
if($devices.length -eq 0) {
    exitWithErrorMsg "Script does not have access to your Audiodevices, try to run it as Admin."
}

$renderer = @()
foreach($device in $devices) {
    if (($device.GetValueNames() | %{$device.GetValue($_)}) -match $playbackDeviceName) {
        $renderer += Get-ItemProperty $device.PSParentPath
    }
}

if($renderer.length -lt 1) {
    exitWithErrorMsg "Could not find any device named $playbackDeviceName"
}

$activeRenderer = @($renderer | Where-Object -Property DeviceState -eq 1)
if($activeRenderer.length -lt 1) {
    exitWithErrorMsg "There are $($renderer.length) devices with Name $playbackDeviceName, but non of them is active"
}
if($activeRenderer.length -gt $maxDeviceCount) {
    $devices
    exitWithErrorMsg "Execution aborted, because more then $maxDeviceCount Active Devices found by Name $playbackDeviceName"
}

$missingLoudness = $false
"Windows Registry Editor Version 5.00" > $regFile
$activeRenderer | ForEach-Object{
    $fxProperties = Join-Path -Path $_.PSPath -ChildPath FxProperties | Get-ItemProperty -ErrorAction Ignore
    if (($fxProperties -eq $null) -or ($fxProperties.$enhancementFlagKey -eq $null) -or 
        ($fxProperties.$enhancementFlagKey[8] -ne 255) -or ($fxProperties.$enhancementFlagKey[9] -ne 255) -or
        ($fxProperties.$releaseTimeKey -eq $null) -or ($fxProperties.$releaseTimeKey[8] -ne $releaseTime) -or
        ($fxProperties.$enhancementTabKey -eq $null) -or ($fxProperties.$enhancementTabKey -ne $enhancementTabValue)) {
        "[" + $fxKeyPath + "]" >> $regFile
        $fxPropertiesImport >> $regFile
        $missingLoudness = $true
    }
    if ($fxProperties -eq $null) {
        Write-Host -NoNewline "FxProperties is missing '$fxKeyPath'" -ForegroundColor Red
        ", it is very likely that import of $regFile will not work since your driver package did not include effects"
        "Try to install a different driver package version or switch to third party sound processing software."
    }
}

if (!$missingLoudness) {
    "Loudness Settings don't need to be enabled"
    Start-Sleep -Seconds 5
    exit 0
}

$currentPrincipal = New-Object Security.Principal.WindowsPrincipal([Security.Principal.WindowsIdentity]::GetCurrent())
if (-not $currentPrincipal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    $arguments = "-File `"$($myInvocation.MyCommand.Definition)`" -playbackDeviceName $playbackDeviceName -maxDeviceCount $maxDeviceCount"
    Start-Process powershell -Verb runAs -ArgumentList $arguments
    exit
}

"import loudness activation into registry"
importReg $regFile

"Restart Audio to apply registry settings"
Restart-Service audiosrv -Force
