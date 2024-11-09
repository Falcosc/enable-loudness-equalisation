#Requires AutoHotkey v2.0

full_command_line := DllCall("GetCommandLine", "str")

if not (A_IsAdmin or RegExMatch(full_command_line, " /restart(?!\S)"))
{
    try
    {
        if A_IsCompiled
            Run '*RunAs "' A_ScriptFullPath '" /restart'
        else
            Run '*RunAs "' A_AhkPath '" /restart "' A_ScriptFullPath '"'
    }
    ExitApp
}

MyGui := Gui(, "Toggle Loudness Equalisation")
MyBtn := MyGui.Add("Button", "w100 h50", "Toggle")

ToggleLoudness(*)
{
    playbackDeviceName := EnvGet("HeadphonesName")
    releaseTime := EnvGet("ReleaseTime")

    if !playbackDeviceName
    {
        MsgBox("Error: HeadphonesName environment variable not set.")
        return
    }

    if !releaseTime
    {
        MsgBox("Error: ReleaseTime environment variable not set.")
        return
    }
    ; The script is assumed to be in the same directory as the AHK script or in the PATH
    Run('powershell.exe -WindowStyle hidden -File "ToggleLoudness.ps1" -playbackDeviceName "' playbackDeviceName '" -maxDeviceCount 1 -releaseTime ' releaseTime)
}

MyBtn.OnEvent("Click", ToggleLoudness)

MyGui.Show("w400 h100")