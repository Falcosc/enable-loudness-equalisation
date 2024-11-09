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
    PS> .\ToggleLoudness.ps1 -playbackDeviceName BE279
    enable loudness equalisation for Audio Devie BE279
.EXAMPLE
    PS> .\ToggleLoudness.ps1 -releaseTime 2
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

function Test-LoudnessEqualizationKey {
    param (
        [string]$DeviceId,
        [string]$RegistryBasePath = "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\MMDevices\Audio\Render"
    )

    $RegistryPath = Join-Path -Path $RegistryBasePath -ChildPath $DeviceId
    $keyName = "{fc52a749-4be9-4510-896e-966ba6525980},3"
    $fullPath = Join-Path -Path $RegistryPath -ChildPath "FxProperties"

    if (Test-Path $fullPath) {
        Write-Host "The registry key $keyName exists."
        return $true
    } else {
        Write-Host "The registry key $keyName does not exist."
        return $false
    }
}

function Get-LoudnessEqualizationState {
    param (
        [string]$DeviceId,
        [string]$RegistryBasePath = "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\MMDevices\Audio\Render"
    )

    if (Test-LoudnessEqualizationKey -DeviceId $DeviceId -RegistryBasePath $RegistryBasePath) {
        $RegistryPath = Join-Path -Path $RegistryBasePath -ChildPath $DeviceId
        $keyName = "{fc52a749-4be9-4510-896e-966ba6525980},3"
        $fullPath = Join-Path -Path $RegistryPath -ChildPath "FxProperties"
        $value = (Get-ItemProperty -Path $fullPath).$keyName

        # Convert byte array to a comparable hex string (e.g., "0b,00,00,00,01,00,00,00,ff,ff,00,00")
        $valueHex = ($value | ForEach-Object { "{0:x2}" -f $_ }) -join ','

        Write-Host "The registry key $keyName has the value $valueHex."

        $trueHex = "0b,00,00,00,01,00,00,00,ff,ff,00,00"
        $falseHex = "0b,00,00,00,01,00,00,00,00,00,00,00"

        if ($valueHex -eq $trueHex) {
            return $true
        } elseif ($valueHex -eq $falseHex) {
            return $false
        } else {
            Write-Host "The registry key $keyName has an unexpected value. Assuming false..."
            return $false
        }
    } else {
        Write-Host "The registry key $keyName does not exist. Assuming false..."
        return $false
    }
}

function Get-ToggledLoudnessEqualizationValue {
    param (
        [string]$DeviceId,
        [string]$RegistryBasePath = "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\MMDevices\Audio\Render"
    )

    $currentState = Get-LoudnessEqualizationState -DeviceId $DeviceId -RegistryBasePath $RegistryBasePath

    $trueHex = "0b,00,00,00,01,00,00,00,ff,ff,00,00"
    $falseHex = "0b,00,00,00,01,00,00,00,00,00,00,00"

    if ($currentState -eq $true) {
        return $falseHex
    } else {
        return $trueHex
    }
}

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
$releaseTimeStr = $releaseTime.ToString().PadLeft(2,'0')
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

"Windows Registry Editor Version 5.00" > $regFile
$activeRenderer | ForEach-Object{
    $fxKeyPath = Join-Path -Path $_.PSPath.Replace("Microsoft.PowerShell.Core\Registry::", "") -ChildPath FxProperties
    $fxProperties = Get-ItemProperty -Path Registry::$fxKeyPath -ErrorAction Ignore
    "[" + $fxKeyPath + "]" >> $regFile

    $toggleValue = Get-ToggledLoudnessEqualizationValue -DeviceId $_.PSChildName
    Write-Host "Setting value to $toggleValue"

    $fxPropertiesImport = @"
"{d04e05a6-594b-4fb6-a80d-01af5eed7d1d},1"="{62dc1a93-ae24-464c-a43e-452f824c4250}" ;PreMixEffectClsid activates effects
"{d04e05a6-594b-4fb6-a80d-01af5eed7d1d},2"="{637c490d-eee3-4c0a-973f-371958802da2}" ;PostMixEffectClsid activates effects
"{d04e05a6-594b-4fb6-a80d-01af5eed7d1d},3"="{5860E1C5-F95C-4a7a-8EC8-8AEF24F379A1}" ;UserInterfaceClsid shows it in ui
"{d04e05a6-594b-4fb6-a80d-01af5eed7d1d},5"="{62dc1a93-ae24-464c-a43e-452f824c4250}" ;StreamEffectClsid
"{d04e05a6-594b-4fb6-a80d-01af5eed7d1d},6"="{637c490d-eee3-4c0a-973f-371958802da2}" ;ModeEffectClsid
"{fc52a749-4be9-4510-896e-966ba6525980},3"=hex:$toggleValue ;enables or disables loudness equalisation
"{9c00eeed-edce-4cd8-ae08-cb05e8ef57a0},3"=hex:03,00,00,00,01,00,00,00,$releaseTimeStr,00,00,00 ;equalisation release time 2 to 7
"@

    $fxPropertiesImport >> $regFile
    if ($fxProperties -eq $null) {
        Write-Host -NoNewline "FxProperties is missing '$fxKeyPath'" -ForegroundColor Red
        ", it is very likely that import of $regFile will not work since your driver package did not include effects"
        "Try to install a different driver package version or switch to third party sound processing software."
    }
}

$currentPrincipal = New-Object Security.Principal.WindowsPrincipal([Security.Principal.WindowsIdentity]::GetCurrent())
if (-not $currentPrincipal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    Write-Host "Please run this script as Administrator."
    exit
}

"import loudness activation into registry"
importReg $regFile

"Restart Audio to apply registry settings"
Restart-Service audiosrv -Force
