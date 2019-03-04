// Do something when some sensor status changes
// 当某些传感器状态发生变化时做某些事情
if (Device["书房门"].State.Equals("close", StringComparison.CurrentCultureIgnoreCase))
{
    MonitorOff();
    if (MediaIsOut())
    {
        Mute();
    }
}

// Reset the state of each sensor, you must usually have this statement.
// 重置各传感器的状态, 通常必须有此语句.
Reset();


// The variable returned by the last run script can be loaded, and 
// the modified variable will also be returned for the next load.
// 可以加载上次运行的脚本返回的变量, 修改过变量也将会返回以便下次载入
var RunAsAdmin = GetVar<string>("RunAsAdmin");
RunAsAdmin = IsAdmin;
// these values will return to system
// 这些值将返回系统
var OpenDoor_Study=$"Duration:{Device["书房门"].StateDuration}";
