# Intro

This application is XiaoMi/MiJia Gateway Monitor, supported 
do some actions by your c# script with the ZigBee device state

Note: it's only support ZigBee device attached to gateway, 
because gateway report the ZigBee device info only.

## Development Environment

### Hardware

1. Aqara/XiaoMi/MiJia Gateway
1. Some ZigBee device attached to gateway above

### Software

1. Windows 10 Home x64
1. Visual Studio Express 2015 for Desktop
1. dotnet framework 4.5 of up
1. [Elton.Aqara](https://github.com/eltonfan/aqara-dotnet-sdk)
1. Roslyn Script Library 1.3.2 (not recommanded upgrade to newest, I try the 2.x version, can not loading assembly dll).
1. AutoIt3.net (AutoItX)
1. NAudio
1. NewtoneSoft.Json

## Features

1. Display gateway heartbeat states every 1s
1. Response device state changing event
1. Custom C# script that runs in the above event
1. Script can read the latest device state & heartbeat info
1. Script can call AutoItX method to automation OS
1. Add some functions for script global call, for examples:
	1. Beep (AutoItX not include these methods)
	1. Sleep
	1. MonitorOff
	1. WinList("notepad") (AutoItX not include this method in .net)
	1. Kill(" firefox$") (process name or title, supoort simple regex syntax)
	1. Media Play/Pause/Stop like global press Play/Pause/Stop key
	1. Mute/UnMute/ToggleMute (support all devices with * or specified devices with device name)
	1. MuteApp/UnMuteApp/ToggleMuteApp (support all app with * or specified app with app name/title)
	1. MediaIsActive Detect Device/App is playing,
	1. Minimize Window (support all window with "*" of null, or specified app with app name/title)

Note: You can calling external utils for more action with AutoItX methods (Run/RunAs/RunWait/RunAsWait), like nircmd 

## Example of gateway config

Path: `<APP>\config\aqara.json` , gateway and device friendly name info took from MiJia App gateway for development page

```json
{
  "gateways": [
    {
      "mac": "00:00:00:00:00:00",
      "password": "****************",
      "devices": [
        {
          "model": "lumi.ctrl_86plug.v1",
          "did": "lumi.00000000000000",
          "name": "<friendly name>"
        }
      ]
    }
  ]
}

```

## Example C# Script

Path: `<APP>\actions.csx`

```csharp
if(IsTest || Device["走道人体传感器"].State.Equals("motion", StringComparison.CurrentCultureIgnoreCase))
{
  Kill(new string[] { "mpc-be64.exe", "zPlayer UWP.exe", });

  var firefox = @"[REGEXPTITLE:(?i) firefox]";
  if(AutoItX.WinExists(firefox) == 1)
  {
    var title = AutoItX.WinGetTitle(firefox);
    if(Regex.IsMatch(title, @"((qidian)|(17k))", RegexOptions.IgnoreCase))
    {
      if(IsTest || AutoItX.IsAdmin() == 1)
      {
        // Set Window State / Send method need permissions for app RunAsAdminstrator
		AutoItX.WinSetState(firefox, "", AutoItX.SW_MAXIMIZE);
        AutoItX.WinActivate(firefox);
        AutoItX.Send("^1");
        AutoItX.Sleep(50);
        AutoItX.Send("{BROWSER_REFRESH}");
        //Minimize(" firefox");
      }
      else
      {
		// Minimize All Window like Win+M
        Minimize();
      }
      //Mute("USB音箱 (JBL Pebbles)");
      //Mute();
      if(MediaIsActive("firefox"))
      {
        MuteApp("firefox");
      }
    }
  }
  Beep();
  //AutoItX.SendKeys("VOLUME_MUTE");
  //Reset("走道人体传感器");
}

if(Device["书房门"].State.Equals("close", StringComparison.CurrentCultureIgnoreCase))
{
  MonitorOff();
  //Device["书房门"].Reset();
  //MediaPause();
  //if(MediasIsActive(new string[]{"cloudmusic", "wallpaper", "firefox"}))
  if(MediaIsActive())
  {
    Mute();
  }
}

// Reset all device state
Reset();

// these values will return to system
var NoMontion_Aisle=$"Duration:{Device["走道人体传感器"].StateDuration}";
var NoMontion_Living=$"Duration:{Device["客厅人体传感器"].StateDuration}";
var OpenDoor_Study=$"Duration:{Device["书房门"].StateDuration}";
var OpenDoor_Main=$"Duration:{Device["入户正门"].StateDuration}";

```