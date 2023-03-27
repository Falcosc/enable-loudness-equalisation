# enable-loudness-equalisation
automatically adds and enables loudness equalisation to any playback device

| before exection | after exection |
| --------------- | -------------- |
| ![Enhancements Missing](EnhancementsMissing.png)  | ![Enhancements Added](EnhancementsAdded.png)  |

# how to run
one time run as admin in powershell
```
EnableLoudness.ps1 -playbackDeviceName <name of your playback device>
```
reusable shortcut
1. target: `powershell.exe -f EnableLoudness.ps1 -playbackDeviceName <name of your playback device>`
2. advanced: run as admin

# When is it needed?
- HDMI or Display Port Playback devices usually doesn't have it
- if you can not find audio driver version which adds loudness equalisation to any of your playback devices
- you can't enable it globally in your driver

# Why does it need to be scripted?
- if you want to toogle it via hotkey
- updates are messing with your audio drivers
- some usecases lead into reregistration of your HDMI or Displayport playback devices which will purge your settings every time

# What does it do?
1. search for all playback devices by name in registry
1. imports audio enhancement settings and sets loudness equalisation flag
1. restarts audio service to apply changed registry values

# Known Issues
- all settings flags stored in `fc52a749-4be9-4510-896e-966ba6525980` get overwritten with instead of just enabling loudness equalisation
- flags key is different accros windows version `fc52a749-4be9-4510-896e-966ba6525980` used in this script works for windows 11, maybe 10 as well.
