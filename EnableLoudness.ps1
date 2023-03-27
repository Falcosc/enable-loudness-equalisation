<#
.SYNOPSIS
    automatically adds and enables loudness equalisation to any playback device
.DESCRIPTION
    Imports registry keys to add enhancement features and enables loudness equalisation.
    It restarts audio service to apply imported registry settings
.NOTES
    https://github.com/Falcosc/enable-loudness-equalisation
#>
Param(
   [Parameter(Mandatory=$true)]
   [string]$playbackDeviceName
) 

$ErrorActionPreference = 'Inquire' #prevents console exit to investigate instead of ignore errors
$regFile = "$env:temp\SoundEnhancementsTMP.reg"
$enhancementFlagsKey = "{fc52a749-4be9-4510-896e-966ba6525980},3"
$enableLoudness = @'
"{d04e05a6-594b-4fb6-a80d-01af5eed7d1d},1"="{62dc1a93-ae24-464c-a43e-452f824c4250}"
"{d04e05a6-594b-4fb6-a80d-01af5eed7d1d},2"="{637c490d-eee3-4c0a-973f-371958802da2}"
"{d04e05a6-594b-4fb6-a80d-01af5eed7d1d},3"="{5860E1C5-F95C-4a7a-8EC8-8AEF24F379A1}"
"{fc52a749-4be9-4510-896e-966ba6525980},3"=hex:0b,00,00,00,01,00,00,00,ff,ff,00,00

'@ #Audio enhancement settings are hardcoded in hex:0b,00,00,00,01,00,00,00,ff,ff,00,00 change it if you have other ones
$missingLoudness = $false

"Windows Registry Editor Version 5.00" > $regFile
$devices = reg query HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\MMDevices\Audio\Render /s /f $playbackDeviceName /d
if(!$?) {
    Write-Error "Could not find any device named $playbackDeviceName"
}

$devices | Select-String "Render" | ForEach-Object{
    $keyPath = $_ -replace 'Properties','FxProperties'
    $fxProperties = Get-ItemProperty -Path Registry::$keyPath
    If (($fxProperties -eq $null) -or ($fxProperties.$enhancementFlagsKey -eq $null) -or 
        ($fxProperties.$enhancementFlagsKey[8] -ne 255) -or ($fxProperties.$enhancementFlagsKey[9] -ne 255)) { 
        "[" + $keyPath + "]" >> $regFile
        $enableLoudness >> $regFile
        $missingLoudness = $true
    }
}

If (!$missingLoudness) {
    "Loudness Settings don't need to be added"
    exit 0
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
    Write-Error "Failed to import $regFile"
}

"Restart Audio to apply registry settings"
Restart-Service audiosrv -Force


