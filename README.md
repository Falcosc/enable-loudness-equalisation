# Enable Loudness Equalisation
Automatically adds and enables loudness equalisation to any playback device.

Only works if your selected driver supports enhancements for speakers, but didn't expose this support for any other output devices. This script will  expose any existing support, but can not work if the driver doesn't ship any.

| before execution | after execution |
| --------------- | -------------- |
| ![Enhancements Missing](EnhancementsMissing.png)  | ![Enhancements Added](EnhancementsAdded.png)  |

If you are looking for bass boost, you can use the more complex version of this script https://github.com/Falcosc/enable-bass-boost

# How to Download and Run
run in powershell
```
Invoke-WebRequest https://raw.githubusercontent.com/Falcosc/enable-loudness-equalisation/main/EnableLoudness.ps1 -OutFile $env:HOMEPATH\EnableLoudness.ps1
Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser
. $env:HOMEPATH\EnableLoudness.ps1
```
Or if you want to set the fastest possible time until sound level gets adjusted (unpleasend to daily usage but gives a competitive edge on video games where dynamic audiolevel adjustments are not banned)
```
. $env:HOMEPATH\EnableLoudness.ps1 -releaseTime 2
```

## Using the Toggle Version with GUI
This script includes a toggle version with an AutoHotkey (AHK) GUI script for easier use:
1. **Install [AutoHotkey v2.0+](https://www.autohotkey.com/)** if you haven't already.
2. Save the `ToggleGui.ahk` script in the same folder as `EnableLoudness.ps1`.
3. Run `ToggleGui.ahk` to open a simple window with a button that toggles loudness equalisation when clicked.

### Environment Variables for the Toggle Script
For the toggle script to work correctly, the following environment variables must be set:
- **`HeadphonesName`**: Specifies the playback device name. This should match the beginning of the device name as shown in your system.
  - Example: 
    ```cmd
    setx HeadphonesName "YourDeviceName"
    ```
- **`ReleaseTime`**: Sets the release time for audio level adjustment, from 2 (fastest) to 7 (slowest).
  - Example:
    ```cmd
    setx ReleaseTime "4"
    ```

> **Note**: Setting these variables ensures the toggle script works with the intended playback device and adjustment speed.

### How to Set Environment Variables
1. Open Command Prompt as an administrator.
2. Use `setx` to set the environment variables:
   ```cmd
   setx HeadphonesName "YourDeviceName"
   setx ReleaseTime "4"
   ```
3. Restart Command Prompt or PowerShell to apply the changes, or reboot your system for a global update.

# When is it needed?
- HDMI, Display Port, Digital Optical Output playback devices usually doesn't have it
- if you can not find an audio driver version which adds loudness equalisation to any of your playback devices
- you can't enable it globally in your driver

# Why does it need to be scripted?
- if you want to toggle it via hotkey
- updates are messing with your audio drivers
- some use cases lead into re-registration of HDMI or DisplayPort playback devices, which will purge your settings every time

# What does it do?
1. search for all active playback devices by name in registry
1. imports audio enhancement settings
    - PreMixEffectClsid and PostMixEffectClsid
    - StreamEffectClsid and ModeEffectClsid
    - Enhancement Tab UI defnition
    - loudness equalisation flag
    - release time value
1. restarts audio service to apply changed registry values

# Known Issues
- all setting flags stored in `fc52a749-4be9-4510-896e-966ba6525980` get overwritten, instead of just enabling loudness equalisation
- flags key are different across Windows versions `fc52a749-4be9-4510-896e-966ba6525980` used in this script works for Windows 11, maybe 10 as well.
- If the playback device gets re-detected the audio service reboot maybe sets volume to default 100%
- Sound Settings UI shows 0% volume if it was open during restart (reopening fixes it)
- Restarting audio service after sleep does break the taskbar tray icon volume slider in some situations
    - mediakeys and sound settings UI volume controll still works fine
    - tray icon slider gets fixed with full reboot
- does not work if your driver doesn't have any enhancements, try a different one
- incompatible devices will be unable to output audio until settings are restored

# Restore Settings
Most drivers restore settings if the registry key get removed, that would be the manual way to restore.
Over UI we found the following way to reset your settings [#156](https://github.com/Falcosc/enable-loudness-equalisation/issues/28)
1. Device Manager
1. Sound, video and game controllers
1. Right-click on your Audio Device
1. Uninstall device, DO NOT check “Delete driver software”
1. Reboot

# Install as Task
1. Open Task Scheduler
1. Action -> Create Task...
1. General -> Run with highest privileges
  
    ![Run with highest privileges](TaskAdmin.png)
1. Triggers -> New...
  
    ![Additional Triggers](TaskTrigger.png)
1. Actions -> New...
    - Action: Start a program
    - Program: powershell
    - Add arguments: `-WindowStyle hidden -f %HOMEPATH%\EnableLoudness.ps1 -playbackDeviceName BE279`
1. To test it you could use an invalid DeviceName like "-playbackDeviceName XXX" then you will see an error message pop-up after login
  
    ![Test Error](ErrorTest.png)
