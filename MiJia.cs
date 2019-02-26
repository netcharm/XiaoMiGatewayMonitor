using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Linq.Expressions;
using System.Net;
using System.Net.Sockets;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

using AutoIt;
using Elton.Aqara;
using KnownFolderPaths;
using Microsoft.CodeAnalysis.CSharp.Scripting;
using Microsoft.CodeAnalysis.Scripting;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;

namespace MiJia
{
    #region Gateway Helper routines
    public class DeviceInfo
    {
        public string Cmd { get; set; }
        [JsonProperty("mac")]
        public string MAC { get; set; }
        [JsonProperty("password")]
        public string Password { get; set; }
        [JsonProperty("model")]
        public string Model { get; set; }
        public string Sid { get; set; }
        [JsonProperty("short_id")]
        public string ShortId { get; set; }
        [JsonProperty("did")]
        public string DeviceId { get; set; }
        public string Token { get; set; }
        public string Data { get; set; }
        //[JsonProperty("devices")]
        //public AqaraDeviceConfig[] Devices { get; set; }
        //[JsonProperty("gateways")]
        //public AqaraGatewayConfig[] Gateways { get; set; }
    }

    public class Gateway
    {
        public string MulticastIP { get; set; } = "224.0.0.50";
        public int MulticastPort { get; set; } = 4321;
        public string RemoteIP { get; set; } = "192.168.1.244";
        public int RemotePort { get; set; } = 9898;
        public string Token { get; set; }

        internal int SendCmd(string cmd)
        {
            IPAddress mip = IPAddress.Parse(MulticastIP);
            IPEndPoint mep = new IPEndPoint(mip, MulticastPort);

            IPAddress rip = IPAddress.Parse(RemoteIP);
            IPEndPoint rep = new IPEndPoint(rip, RemotePort);

            UdpClient udp = new UdpClient();
            udp.JoinMulticastGroup(mip);
            udp.Client.SetSocketOption(SocketOptionLevel.Socket, SocketOptionName.ReuseAddress, true);
            udp.ExclusiveAddressUse = false;
            udp.Connect("127.0.0.1", RemotePort);

            byte[] bs = Encoding.Default.GetBytes(cmd);
            return udp.Send(bs, bs.Length);
        }

        internal async Task<string> SendCmd(string cmd, string server, int port)
        {
            string result = string.Empty;

            UdpClient udp = new UdpClient();
            udp.Client.SetSocketOption(SocketOptionLevel.Socket, SocketOptionName.ReuseAddress, true);
            udp.ExclusiveAddressUse = false;
            udp.Connect(server, port);

            byte[] bs = Encoding.Default.GetBytes(cmd);
            var ret = udp.Send(bs, bs.Length);
            var received = await udp.ReceiveAsync();
            result = Encoding.Default.GetString(received.Buffer);

            return (result);
        }

        public async Task<List<string>> Listen()
        {
            List<string> result = new List<string>();

            IPAddress mip = IPAddress.Parse(MulticastIP);
            IPEndPoint mep = new IPEndPoint(mip, MulticastPort);

            IPAddress rip = IPAddress.Parse(RemoteIP);
            IPEndPoint rep = new IPEndPoint(rip, RemotePort);

            UdpClient udp = new UdpClient();
            udp.JoinMulticastGroup(mip);
            udp.Client.SetSocketOption(SocketOptionLevel.Socket, SocketOptionName.ReuseAddress, true);
            udp.ExclusiveAddressUse = false;

            udp.Client.Bind(new IPEndPoint(IPAddress.Any, RemotePort));

            while (true)
            {
                var received = await udp.ReceiveAsync();
                var line = Encoding.Default.GetString(received.Buffer);
                if (!string.IsNullOrEmpty(line))
                {
                    result.Add(line);
                    break;
                }
            }

            return (result);
        }

        public async Task<IList<string>> GetDevice()
        {
            List<string> result = new List<string>();

            //var ret = SendCmd("{\"cmd\":\"whois\"}");
            var ret = await SendCmd("{\"cmd\" : \"get_id_list\"}", RemoteIP, RemotePort);

            JsonSerializer serializer = new JsonSerializer();
            DeviceInfo token = JsonConvert.DeserializeObject<DeviceInfo>(ret);
            if (!string.IsNullOrEmpty(token.Data))
            {
                var data = JsonConvert.DeserializeObject<string[]>(token.Data);
                foreach (var device in data)
                    result.Add(device);
            }

            return (result);
        }

        public async Task<IList<string>> GetDevice(string deviceId)
        {
            List<string> result = new List<string>();

            if (!string.IsNullOrEmpty(deviceId))
            {
                var ret = await SendCmd($"{{\"cmd\":\"read\", \"sid\":\"{deviceId}\"}}", RemoteIP, RemotePort);

                JsonSerializer serializer = new JsonSerializer();
                DeviceInfo token = JsonConvert.DeserializeObject<DeviceInfo>(ret);
                if (!string.IsNullOrEmpty(token.Data))
                {
                    JToken data = JsonConvert.DeserializeObject<JToken>(token.Data);
                    if (data["status"] != null)
                    {
                        var status = data["status"];
                    }
                }
            }

            return (result);
        }
    }
    #endregion

    public enum ConditionMode { AND, OR, NOR, XOR };
    public class Condition<T>
    {
        public ConditionMode Mode;
        public KeyValuePair<string, T> Param { get; set; }
    }

    public enum ActionMode { Close, Minimize, Maximize, Mute };
    public class Action<T>
    {
        public string Name { get; set; }
        public ActionMode Mode { get; set; }
        public IList<Condition<T>> Conditions { get; set; }
        public IList<string> Param { get; set; }
    }

    public static class Extensions
    {
        private delegate void SetPropertyThreadSafeDelegate<TResult>(
            Control @this,
            Expression<Func<TResult>> property,
            TResult value);

        public static void SetPropertyThreadSafe<TResult>(
            this Control @this,
            Expression<Func<TResult>> property,
            TResult value)
        {
            var propertyInfo = (property.Body as MemberExpression).Member as PropertyInfo;

            if (propertyInfo == null ||
                !@this.GetType().IsSubclassOf(propertyInfo.ReflectedType) ||
                @this.GetType().GetProperty(
                    propertyInfo.Name,
                    propertyInfo.PropertyType) == null)
            {
                throw new ArgumentException("The lambda expression 'property' must reference a valid property on this Control.");
            }

            if (@this.InvokeRequired)
            {
                @this.Invoke(new SetPropertyThreadSafeDelegate<TResult>
                (SetPropertyThreadSafe),
                new object[] { @this, property, value });
            }
            else
            {
                @this.GetType().InvokeMember(
                    propertyInfo.Name,
                    BindingFlags.SetProperty,
                    null,
                    @this,
                    new object[] { value });
            }
        }
    }

    public class ScriptEngine
    {
        private string APPFOLDER = Path.GetDirectoryName(Application.ExecutablePath);
        private string SKYFOLDER = Path.Combine(KnownFolders.GetPath(KnownFolder.SkyDrive), @"ApplicationData\ConnectedHome");
        private string DOCFOLDER = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.CommonDocuments), @"Elton\ConnectedHome\");
        private string USERNAME = Environment.UserName;

        public TextBox Logger { get; set; } = null;

        #region Kill Process
        internal static bool KillProcess(string processName)
        {
            bool result = true;

            //Process[] pslist = Process.GetProcesses();
            var pName = Path.GetFileNameWithoutExtension(processName);
            Process[] pslist = Process.GetProcessesByName(pName);
            foreach (var ps in pslist)
            {
                if (ps.ProcessName.EndsWith(pName, StringComparison.CurrentCultureIgnoreCase))
                {
                    ps.Kill();
                    System.Media.SystemSounds.Asterisk.Play();
                }
            }

            return (result);
        }

        internal static bool KillProcess(int pid)
        {
            bool result = true;

            if (pid > 0)
            {
                Process ps = Process.GetProcessById(pid);
                if (ps is Process)
                    ps.Kill();
            }
            else result = false;

            return (result);
        }

        internal static void KillProcess(string[] processNames)
        {
            foreach (var processName in processNames)
            {
                try
                {
                    var pid = AutoItX.ProcessExists(processName);
                    if (pid > 0)
                    {
                        var ret = AutoItX.ProcessClose(processName);
                        if (ret > 0) KillProcess(pid);
                        //if (ret > 0) KillProcess(processName);
                    }
                    //KillProcess(processName);
                }
                catch (Exception) { }
            }
        }
        #endregion

        #region Power On/Off Monitor
        private const int LCI_WM_SYSCommand = 274;
        private const int LCI_SC_MonitorPower = 61808;
        private const int LCI_Power_Off = 2;
        private const int LCI_Power_On = -1;

        //[DllImport("coredll.dll")]
        //private static extern IntPtr SendMessage(IntPtr hWnd, int Msg, IntPtr wParam, IntPtr lParam);

        [DllImport("user32.dll")]
        public static extern int SendMessage(int hWnd, uint Msg, int wParam, int lParam);

        internal static void Monitor(bool on)
        {
            var handle = AutoItX.WinGetHandle("[CLASS:Progman]");
            if (handle != IntPtr.Zero)
            {
                if (on)
                    SendMessage(handle.ToInt32(), LCI_WM_SYSCommand, LCI_SC_MonitorPower, LCI_Power_On);
                else
                    SendMessage(handle.ToInt32(), LCI_WM_SYSCommand, LCI_SC_MonitorPower, LCI_Power_Off);
            }
        }
        #endregion

        #region MiJiaGateway routines
        private Timer timerRefresh = null;

        private async void TimerRefresh_Tick(object sender, EventArgs e)
        {
            var gateway = client.Gateways.Values.FirstOrDefault();
            if (gateway == null)
            {
                if (Logger is TextBox) Logger.Text = string.Empty;
            }
            else
            {
                if (Logger is TextBox)
                {
                    var sb = new StringBuilder();
                    sb.AppendLine(gateway.EndPoint?.ToString());
                    sb.AppendLine(gateway.LatestTimestamp.ToString());
                    sb.AppendLine(gateway.Token);
                    sb.AppendLine("--------------------------------");

                    foreach (var device in gateway.Devices.Values)
                    {
                        sb.AppendLine($"{device.Name}");
                        foreach (var pair in device.States)
                        {
                            sb.Append($"{pair.Key} = {pair.Value.Value} ");
                        }
                        sb.AppendLine();
                    }
                    Logger.Text = sb.ToString();
                }
            }
            await RunScript();
        }

        private AqaraConfig config = null;
        private AqaraClient client = null;
        private Dictionary<string, DEVICE> Devices = new Dictionary<string, DEVICE>();

        private async void DeviceStateChanged(object sender, StateChangedEventArgs e)
        {
            if (Devices.ContainsKey(e.Device.Name) && Devices[e.Device.Name] is DEVICE)
            {
                Devices[e.Device.Name].State = e.NewData;
                Devices[e.Device.Name].Info = e.Device;
                Devices[e.Device.Name].StateDuration = 0;
            }
            else
                Devices[e.Device.Name] = new DEVICE() { State = e.NewData, Info = e.Device };

            await RunScript();
        }

        internal void InitMiJiaGateway(string basepath, string configFile)
        {
            if (File.Exists(Path.Combine(basepath, "aqara.json")))
            {
                config = AqaraConfig.Parse(File.ReadAllText(configFile));
            }
            else if (File.Exists(configFile))
            {
                config = AqaraConfig.Parse(File.ReadAllText(configFile));
            }
            client = new AqaraClient(config);
            client.DeviceStateChanged += DeviceStateChanged;
            Task.Run(() => { client.DoWork(null); });

            timerRefresh = new Timer();
            timerRefresh.Interval = 1000;
            timerRefresh.Tick += TimerRefresh_Tick;
            timerRefresh.Start();
        }
        #endregion

        #region CSharp Script routines
        private bool sciptRunning = false;
        private ScriptOptions scriptOptions = ScriptOptions.Default;
        public string ScriptContext { get; set; } = string.Empty;
        //internal Script script;

        internal ScriptOptions InitScriptEngine()
        {
            scriptOptions = ScriptOptions.Default;
            //options = options.AddReferences(AppDomain.CurrentDomain.GetAssemblies());
            scriptOptions = scriptOptions.AddReferences(new Assembly[] {
                    Assembly.GetAssembly(typeof(Path)),
                    Assembly.GetAssembly(typeof(AutoItX)),
                    Assembly.GetAssembly(typeof(JsonConvert)),
                    Assembly.GetAssembly(typeof(AqaraDevice)),
                    Assembly.GetCallingAssembly(),
                    Assembly.GetEntryAssembly(),
                    Assembly.GetExecutingAssembly(),
                    Assembly.GetAssembly(typeof(System.Dynamic.DynamicObject)),  // System.Code
                    Assembly.GetAssembly(typeof(Microsoft.CSharp.RuntimeBinder.CSharpArgumentInfo)),  // Microsoft.CSharp
                    Assembly.GetAssembly(typeof(System.Dynamic.ExpandoObject))  // System.Dynamic                    
                });
            scriptOptions = scriptOptions.AddImports(new string[] {
            //scriptOptions = scriptOptions.AddNamespaces(new string[] {
                    "System",
                    "System.Dynamic",
                    "AutoIt",
                    "Newtonsoft.Json",
                    "Elton.Aqara",
                    "MiJia",
                });
            //options = options.AddReferences(
            //    Assembly.GetAssembly(typeof(Path)),
            //    Assembly.GetAssembly(typeof(AutoItX)),
            //    Assembly.GetAssembly(typeof(JsonConvert)),
            //    Assembly.GetAssembly(typeof(AqaraDevice)),
            //    Assembly.GetAssembly(GetType()),
            //    Assembly.GetAssembly(typeof(System.Dynamic.DynamicObject)),  // System.Code
            //    Assembly.GetAssembly(typeof(Microsoft.CSharp.RuntimeBinder.CSharpArgumentInfo)),  // Microsoft.CSharp
            //    Assembly.GetAssembly(typeof(System.Dynamic.ExpandoObject))  // System.Dynamic                    
            //);
            //options = options.AddNamespaces(
            //    "System",
            //    "System.Dynamic",
            //    "AutoIt",
            //    "Newtonsoft.Json",
            //    "Elton.Aqara",
            //    "MiJia"
            //);

            var sf = Path.Combine(APPFOLDER, "actions.csx");
            if (File.Exists(sf))
                ScriptContext = File.ReadAllText(sf);

            return (scriptOptions);
        }

        internal void Init(string basepath, string configFile, TextBox logger)
        {
            InitMiJiaGateway(basepath, configFile);
            InitScriptEngine();
            if(logger is TextBox) Logger = logger;
        }

        internal async Task<ScriptState> RunScript(bool AutoReset = false)
        {
            ScriptState result = null;
            if (sciptRunning) return (result);
            sciptRunning = true;

            var gateway = client.Gateways.Values.FirstOrDefault();
            if (gateway == null)
            {
                if(Logger is TextBox)
                    Logger.SetPropertyThreadSafe(() => Logger.Text, string.Empty);
            }
            else
            {
                foreach (var device in gateway.Devices.Values)
                {
                    if (!Devices.ContainsKey(device.Name))
                        Devices[device.Name] = new DEVICE() { State = string.Empty, Info = device };
                    else
                        Devices[device.Name].StateDuration++;
                }

                try
                {
                    var globals = new Globals()
                    {
                        Device = Devices,
                    };
                    result = await CSharpScript.RunAsync(ScriptContext, scriptOptions, globals);
                    if (AutoReset) globals.Reset();

                    StringBuilder sb = new StringBuilder();
                    sb.AppendLine("--------------------------------");
                    foreach (var v in result.Variables)
                    {
                        if (v.Name.Equals("Device", StringComparison.CurrentCultureIgnoreCase)) continue;
                        //sb.AppendLine($"{v.Name} = {v.Value}, {v.Type}");
                        sb.AppendLine($"{v.Name} = {v.Value}");
                    }
                    if (Logger is TextBox)
                        Logger.SetPropertyThreadSafe(() => Logger.Text, Logger.Text + sb.ToString());
                }
                catch (Exception ex)
                {
                    if (Logger is TextBox)
                        Logger.SetPropertyThreadSafe(() => Logger.Text, Logger.Text + ex.Message);
                }
            }

            sciptRunning = false;
            return (result);
        }
        #endregion

    }

    public class DEVICE
    {
        public string State { get; set; } = string.Empty;
        public uint StateDuration { get; set; } = 0;
        public AqaraDevice Info { get; set; }

        public void Reset()
        {
            State = string.Empty;
        }
    }

    public class Globals
    {
        public Dictionary<string, DEVICE> Device = new Dictionary<string, DEVICE>();

        public void Reset(string device = "*")
        {
            foreach (var dev in Device)
            {
                if (string.IsNullOrEmpty(device) || device.Equals("*") || device.Equals(dev.Key, StringComparison.CurrentCulture))
                    dev.Value.Reset();
            }
        }

        public void Minimize(string window="*")
        {
            if (string.IsNullOrEmpty(window) || window.Equals("*"))
                AutoItX.WinMinimizeAll();
            else
                AutoItX.WinSetState($"[REGEXPTITLE:(?i){window}]", "", AutoItX.SW_MINIMIZE);
        }

        public void Minimize(string[] windows)
        {
            if (windows.Length == 0)
                AutoItX.WinMinimizeAll();
            else
            {
                foreach (var win in windows)
                {
                    Minimize(win);
                }
            }
        }

        public void Kill(int pid)
        {
            ScriptEngine.KillProcess(pid);
        }

        public void Kill(string process)
        {
            ScriptEngine.KillProcess(process);
        }

        public void Kill(string[] processList)
        {
            ScriptEngine.KillProcess(processList);
        }

        public void Monitor(bool on)
        {
            ScriptEngine.Monitor(on);
        }

        public void MonitorOn()
        {
            ScriptEngine.Monitor(true);
        }

        public void MonitorOff()
        {
            ScriptEngine.Monitor(false);
        }

        public void Mute(bool on, string device = "*")
        {
            try
            {
                //Instantiate an Enumerator to find audio devices
                NAudio.CoreAudioApi.MMDeviceEnumerator MMDE = new NAudio.CoreAudioApi.MMDeviceEnumerator();
                //Get all the devices, no matter what condition or status
                NAudio.CoreAudioApi.MMDeviceCollection DevCol = MMDE.EnumerateAudioEndPoints(NAudio.CoreAudioApi.DataFlow.All, NAudio.CoreAudioApi.DeviceState.All);
                //Loop through all devices
                foreach (NAudio.CoreAudioApi.MMDevice dev in DevCol)
                {
                    try
                    {
                        if (dev.State == NAudio.CoreAudioApi.DeviceState.Active)
                        {
                            //Show us the human understandable name of the device
#if DEBUG
                            Debug.Print(dev.FriendlyName);
#endif
                            if (string.IsNullOrEmpty(device) || device.Equals("*") || dev.FriendlyName.Equals(device, StringComparison.CurrentCultureIgnoreCase))
                            {
                                //Mute it
                                if (on)
                                    dev.AudioEndpointVolume.Mute = true;
                                else
                                    dev.AudioEndpointVolume.Mute = false;
                            }
                        }
                    }
                    catch (Exception)
                    {
                        //Do something with exception when an audio endpoint could not be muted
                    }
                }
            }
            catch (Exception)
            {
                //When something happend that prevent us to iterate through the devices
            }

            //AutoItX.Send("{VOLUME_MUTE}");
            //AutoItX.Sleep(10);
        }

        public void Mute(string device = "*")
        {
            Mute(true, device);
        }

        public void UnMute(string device = "*")
        {
            Mute(false, device);
        }

    }
}
