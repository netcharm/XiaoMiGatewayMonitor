
if (Device["书房门"].State.Equals("close", StringComparison.CurrentCultureIgnoreCase))
{
    MonitorOff();
    if (MediaIsActive())
    {
        Mute();
    }
}

// Reset all device state
Reset();

// these values will return to system
var OpenDoor_Study=$"Duration:{Device["书房门"].StateDuration}";
