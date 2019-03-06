using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Linq.Expressions;
using System.Management;
using System.Net;
using System.Net.Sockets;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Dynamic;

using AutoIt;
using Elton.Aqara;
using KnownFolderPaths;
using Microsoft.CodeAnalysis.CSharp.Scripting;
using Microsoft.CodeAnalysis.Scripting;
using Microsoft.CodeAnalysis.Scripting.Hosting;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using NAudio.CoreAudioApi;

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

        public static void Update(this Control control, string text)
        {
            control.Invoke((MethodInvoker)delegate {
                // Running on the UI thread
                control.Text = text;
            });
        }

        public static string PaddingLeft(this string text, int width, char paddingchar = ' ')
        {
            var plen = width - Encoding.GetEncoding("gbk").GetBytes(text).Length;
            var padding = new string(paddingchar, plen > 0 ? plen : 0);
            return ($"{padding}{text}");
        }

        public static string PaddingRight(this string text, int width, char paddingchar = ' ')
        {
            var plen = width - Encoding.GetEncoding("gbk").GetBytes(text).Length;
            var padding = new string(paddingchar, plen > 0 ? plen : 0);
            return ($"{text}{padding}");
        }

        #region MiJia Device Exts
        public static bool IsOpen(this AqaraDevice device)
        {
            bool result = false;

            if(device is AqaraDevice)
            {
                if(device.States.ContainsKey("state"))
                {

                }
            }

            return (result);
        }

        #endregion
    }

    public class ScriptEngine
    {
        private string APPFOLDER = Path.GetDirectoryName(Application.ExecutablePath);
        private string SKYFOLDER = Path.Combine(KnownFolders.GetPath(KnownFolder.SkyDrive), @"ApplicationData\ConnectedHome");
        private string DOCFOLDER = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.CommonDocuments), @"Elton\ConnectedHome\");
        private string USERNAME = Environment.UserName;

        public bool Pausing { get; set; } = false;
        public TextBox Logger { get; set; } = null;

        #region MiJiaGateway routines
        private Timer timerRefresh = null;
        private int interval = 1000;
        public int Interval
        {
            get { return (interval); }
            set
            {
                interval = value;
                if (timerRefresh is Timer)
                {
                    if(timerRefresh.Enabled) timerRefresh.Stop();
                    timerRefresh.Interval = value;
                    timerRefresh.Start();
                }
            }
        }

        private async void TimerRefresh_Tick(object sender, EventArgs e)
        {
            var gateway = client.Gateways.Values.FirstOrDefault();
            if (gateway == null)
            {
                if (Logger is TextBox) Logger.Update(string.Empty); //Logger.Text = string.Empty;
            }
            else
            {
                var maxlen_device = 0;
                foreach (var device in gateway.Devices.Values)
                {
                    if (!Devices.ContainsKey(device.Name))
                        Devices[device.Name] = device;
                    //Devices[device.Name] = new DEVICE() {
                    //    client = client,
                    //    Properties = device.States.ToDictionary(s => s.Key, s => s.Value.Value),
                    //    Info = device
                    //};
                    else
                        Devices[device.Name].StateDuration++;

                    var len_device = Encoding.GetEncoding("gbk").GetBytes(device.Name).Length;
                    if (len_device > maxlen_device) maxlen_device = len_device;
                }

                if (Logger is TextBox)
                {
                    var sb = new StringBuilder();
                    sb.AppendLine($"{gateway.EndPoint?.ToString()}[{gateway.Id}]]");
                    sb.AppendLine($"{gateway.LatestTimestamp.ToString()} : {gateway.Token}");
                    sb.AppendLine("".PaddingRight(72, '-'));

                    foreach (var device in gateway.Devices.Values)
                    {
                        List<string> psl = new List<string>();
                        foreach (var pair in device.States)
                        {
                            psl.Add($"{pair.Key} = {pair.Value.Value}");
                        }
                        sb.AppendLine($"{device.Name.PaddingRight(maxlen_device)}[{string.Join(",", psl)}]");
                    }
                    Logger.Update(sb.ToString());
                }
            }
            if(!Pausing) await RunScript();
        }

        private AqaraConfig config = null;
        private AqaraClient client = null;
        //private Dictionary<string, DEVICE> Devices = new Dictionary<string, DEVICE>();
        private Dictionary<string, dynamic> Devices = new Dictionary<string, dynamic>();

        private async void DeviceStateChanged(object sender, StateChangedEventArgs e)
        {
            //if (Devices.ContainsKey(e.Device.Name) && Devices[e.Device.Name] is DEVICE)
            //{
            //    Devices[e.Device.Name].State = e.NewData;
            //    Devices[e.Device.Name].StateName = e.StateName;
            //    Devices[e.Device.Name].Properties = e.Device.States.ToDictionary(s => s.Key, s => s.Value.Value);
            //    Devices[e.Device.Name].Info = e.Device;
            //    Devices[e.Device.Name].StateDuration = 0;
            //    if (e.Device is AqaraDevice)
            //    {
            //        Devices[e.Device.Name].Info.NewStateName = e.StateName;
            //        Devices[e.Device.Name].Info.NewState.Key = e.StateName;
            //        Devices[e.Device.Name].Info.NewState.Value = e.NewData;
            //        Devices[e.Device.Name].Info.StateDuration = 0;
            //    }
            //}
            //else
            //{
            //    Devices[e.Device.Name] = new DEVICE()
            //    {
            //        client = client,
            //        State = e.NewData,
            //        Properties = new Dictionary<string, string>(),
            //        StateName = e.StateName,
            //        Info = e.Device,
            //    };
            //    if (e.Device is AqaraDevice)
            //    {
            //        Devices[e.Device.Name].Info.NewStateName = e.StateName;
            //        Devices[e.Device.Name].Info.NewState.Key = e.StateName;
            //        Devices[e.Device.Name].Info.NewState.Value = e.NewData;
            //    }
            //}
            Devices[e.Device.Name] = e.Device;
            if (e.Device is AqaraDevice)
            {
                Devices[e.Device.Name].NewStateName = e.StateName;
                Devices[e.Device.Name].NewStateValue = e.NewData;
                Devices[e.Device.Name].StateDuration = 0;
            }

            if (!Pausing) await RunScript();
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
            timerRefresh.Tick += TimerRefresh_Tick;
            timerRefresh.Interval = Interval;
            timerRefresh.Start();
        }

        #endregion

        #region CSharp Script routines
        private bool sciptRunning = false;
        private string scriptContext = string.Empty;
        public string ScriptContext
        {
            get { return (scriptContext); }
            set { Load(value); }
        }

        private Globals globals = new Globals();
        private System.Threading.CancellationToken cancelToken = new System.Threading.CancellationToken();

        private Script script;
        private InteractiveAssemblyLoader loader = new InteractiveAssemblyLoader();
        private ScriptOptions scriptOptions = ScriptOptions.Default;

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
                    Assembly.GetAssembly(typeof(System.Globalization.CultureInfo)),
                    Assembly.GetAssembly(typeof(Math)),
                    Assembly.GetAssembly(typeof(Regex)),
                    Assembly.GetAssembly(typeof(DynamicObject)),  // System.Dynamic
                    Assembly.GetAssembly(typeof(ExpandoObject)), // System.Dynamic                    
                    Assembly.GetAssembly(typeof(Microsoft.CSharp.RuntimeBinder.CSharpArgumentInfo)),  // Microsoft.CSharp
                });
            scriptOptions = scriptOptions.AddImports(new string[] {
                    "System",
                    "System.Collections.Generic",
                    "System.Dynamic",
                    "System.Globalization",
                    "System.IO",
                    "System.Math",
                    "System.Text",
                    "System.Text.RegularExpressions",
                    "AutoIt",
                    "Newtonsoft.Json",
                    "Elton.Aqara",
                    "MiJia",
                });

            var sf = Path.Combine(APPFOLDER, "actions.csx");
            if (File.Exists(sf))
                Load(File.ReadAllText(sf));

            return (scriptOptions);
        }

        public void Load(string context = "")
        {
            if(!string.IsNullOrEmpty(context)) scriptContext = context;
            script = CSharpScript.Create(scriptContext, scriptOptions, typeof(Globals), loader);
            script.Compile();
        }

        internal void Init(string basepath, string configFile, TextBox logger)
        {
            InitMiJiaGateway(basepath, configFile);
            InitScriptEngine();
            if(logger is TextBox) Logger = logger;
        }

        internal async Task<ScriptState> RunScript(bool AutoReset = false, bool IsTest = false)
        {
            ScriptState result = null;
            if (sciptRunning || string.IsNullOrEmpty(scriptContext))
            {
                return (result);
            }

            sciptRunning = true;

            var gateway = client.Gateways.Values.FirstOrDefault();
            if (gateway == null)
            {
                if (Logger is TextBox)
                    Logger.Update(string.Empty);
                    //Logger.SetPropertyThreadSafe(() => Logger.Text, string.Empty);
            }
            else
            {
                try
                {
                    globals.device = Devices;
                    globals.isTest = IsTest;
                    globals.Logger.Clear();

                    //result = await CSharpScript.RunAsync(ScriptContext, scriptOptions, globals);
                    if (!(script is Script)) Load();
                    result = await script.RunAsync(globals, cancelToken);

                    if (AutoReset) globals.Reset();

                    StringBuilder sb = new StringBuilder();
                    sb.AppendLine("".PaddingRight(72, '-'));// "--------------------------------");
                    foreach (var line in globals.Logger)
                    {
                        sb.AppendLine(line);
                    }
                    if(globals.Logger.Count>0)
                        sb.AppendLine("".PaddingRight(72, '-'));// "--------------------------------");
                    foreach (var v in result.Variables)
                    {
                        if (v.Name.Equals("Device", StringComparison.CurrentCultureIgnoreCase)) continue;
                        sb.AppendLine($"{v.Name} = {v.Value}");
                        globals.vars[v.Name] = v.Value;
                    }
                    if(result.Variables.Count()>0)
                        sb.AppendLine("".PaddingRight(72, '-'));// "--------------------------------");
                    if (Logger is TextBox)
                        Logger.Update(Logger.Text + sb.ToString());
                    //Logger.SetPropertyThreadSafe(() => Logger.Text, Logger.Text + sb.ToString());
                }
                catch (Exception ex)
                {
                    if (Logger is TextBox)
                        Logger.Update(Logger.Text + ex.Message);
                        //Logger.SetPropertyThreadSafe(() => Logger.Text, Logger.Text + ex.Message);
                }
            }

            sciptRunning = false;
            return (result);
        }
        #endregion

    }

    public class DEVICE
    {
        internal AqaraClient client = default(AqaraClient);
        public string State { get; internal set; } = string.Empty;
        public string StateName { get; internal set; } = string.Empty;
        public uint StateDuration { get; set; } = 0;
        public Dictionary<string, string> Properties { get; set; } = new Dictionary<string, string>();
        //private dynamic device = new ExpandoObject();
        //public dynamic Info
        //{
        //    get
        //    {
        //        return (device);
        //    }
        //    set
        //    {
        //        device = value;
        //        device.
        //    }
        //}
        public dynamic Info { get; set; }
        //public AqaraDevice Info { get; set; } = default(AqaraDevice);

        public bool Open { get; }

        public void SetState(string key, string value)
        {
            if (client is AqaraClient)
            {
                List<KeyValuePair<string, string>> states = new List<KeyValuePair<string, string>>();
                KeyValuePair<string, string> kv = new KeyValuePair<string, string>(key, value);
                states.Add(kv);
                SetStates(states);
            }
        }

        public void SetStates(IEnumerable<KeyValuePair<string, string>> states)
        {
            if (client is AqaraClient && states is IEnumerable<KeyValuePair<string, string>>)
            {
                client.SendWriteCommand(Info as AqaraDevice, states);
            }
        }

        public void Reset()
        {
            State = string.Empty;
            StateName = string.Empty;
            if (Info is AqaraDevice)
            {
                Info.NewStateName = string.Empty;
            }
        }
    }

    public class Globals
    {
        public enum MUTE_MODE { Mute, UnMute, Toggle, Background }

        internal bool isTest = false;
        public bool IsTest { get { return (isTest); } }

        internal Dictionary<string, object> vars = new Dictionary<string, object>();
        public T GetVar<T>(string vn)
        {
            T result = default(T);
            try
            {
                if (vars.ContainsKey(vn)) result = (T)vars[vn];
                else result = default(T);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            return (result);
        }

        private class ProcInfo
        {
            uint PID { get; set; } = 0;
            uint Parent { get; set; } = 0;
            string Name { get; set; } = string.Empty;
            string Title { get; set; } = string.Empty;
            Process Info { get; set; } = default(Process);
        }

        public bool IsAdmin { get; set; } = AutoItX.IsAdmin() == 1 ? true : false;
        Dictionary<uint, Process> procs = null;
        private ManagementEventWatcher _watcherStart;
        private ManagementEventWatcher _watcherStop;
        public Globals()
        {
            if (IsAdmin)
            {
                _watcherStart = new ManagementEventWatcher("SELECT ProcessID, ProcessName FROM Win32_ProcessStartTrace");
                _watcherStart.EventArrived += ProcessStarted;
                _watcherStop = new ManagementEventWatcher("SELECT * FROM Win32_ProcessStopTrace");
                _watcherStop.EventArrived += ProcessStoped;
                _watcherStart.Start();
                _watcherStop.Start();
            }

            procs = Process.GetProcesses().ToDictionary(p => (uint)p.Id, p => p);
        }

        ~Globals()
        {
            if (IsAdmin)
            {
                _watcherStop.Stop();
                _watcherStart.Stop();
            }
        }

        #region MiJia Gateway/ZigBee Device
        //internal Dictionary<string, DEVICE> device = new Dictionary<string, DEVICE>();
        //public Dictionary<string, DEVICE> Device { get { return (device); } }
        internal Dictionary<string, dynamic> device = new Dictionary<string, dynamic>();
        public Dictionary<string, dynamic> Device { get { return (device); } }

        public void Reset(string devname = default(string))
        {
            foreach (var dev in device)
            {
                if (string.IsNullOrEmpty(devname) || device.Equals("*") || devname.Equals(dev.Key, StringComparison.InvariantCulture))
                    dev.Value.NewStateName = string.Empty;
                    //dev.Value.Reset();
            }
            isTest = false;
        }
        #endregion

        #region Windows routines
        private const int SW_SHOWNORMAL = 1;
        private const int SW_SHOWMINIMIZED = 2;
        private const int SW_SHOWMAXIMIZED = 3;

        [DllImport("user32.dll", CharSet = CharSet.Auto)]
        private static extern bool ShowWindowAsync(IntPtr hWnd, int nCmdShow);

        public void Minimize(string window = default(string))
        {
            if (string.IsNullOrEmpty(window) || window.Equals("*"))
                AutoItX.WinMinimizeAll();
            else
            {
                AutoItX.WinSetState($@"[REGEXPTITLE:(?i){window}]", "", AutoItX.SW_MINIMIZE);
                //if (IsAdmin)
                //    AutoItX.WinSetState($@"[REGEXPTITLE:(?i){window}]", "", AutoItX.SW_MINIMIZE);
                //else
                //{
                //    var handle = AutoItX.WinGetHandle($@"[REGEXPTITLE:(?i){window}]");
                //    if (handle == IntPtr.Zero) return;
                //    if (AutoItX.WinGetState(handle) != AutoItX.SW_MINIMIZE)
                //        ShowWindowAsync(handle, SW_SHOWMINIMIZED);
                //}
            }
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

        public Dictionary<IntPtr, string> WinList(string title)
        {
            Dictionary<IntPtr, string> result = new Dictionary<IntPtr, string>();

            var plist = Process.GetProcesses();
            foreach (var p in plist)
            {
                var pid = new IntPtr(p.Id);
                var ptitle = p.MainWindowTitle; // AutoItX.WinGetTitle(pid);
                var pname = p.ProcessName;
                if (Regex.IsMatch(pname, $@"{Path.GetFileNameWithoutExtension(title)}", RegexOptions.IgnoreCase) ||
                    Regex.IsMatch(ptitle, $@"{title}", RegexOptions.IgnoreCase))
                {
                    if (p.SafeHandle.IsInvalid) continue;
                    if (p.MainWindowHandle == IntPtr.Zero) continue;
                    result[pid] = ptitle;
                }
            }

            return (result);
        }

        #region UWP Window Process Name : https://stackoverflow.com/a/50554419/1842521
        internal struct WINDOWINFO
        {
            public uint ownerpid;
            public uint childpid;
        }

        [DllImport("user32.dll", EntryPoint = "FindWindow", CharSet = CharSet.Auto)]
        private static extern int FindWindow(string sClass, string sWindow);

        [DllImport("user32.dll", EntryPoint = "FindWindowEx", CharSet = CharSet.Auto)]
        private static extern IntPtr FindWindowEx(IntPtr hwndParent, IntPtr hwndChildAfter, string lpszClass, string lpszWindow);

        [DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]
        private static extern int GetWindowText(IntPtr hWnd, StringBuilder lpString, int nMaxCount);

        [DllImport("User32.Dll", CharSet = CharSet.Auto)]
        private static extern int GetClassName(int hwnd, string lpClassName, int nMaxCount);

        [DllImport("user32.dll", SetLastError = true)]
        private static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint lpdwProcessId);
        
        [DllImport("user32", SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        private static extern bool EnumChildWindows(IntPtr hWndParent, EnumWindowProc lpEnumFunc, IntPtr lParam);

        [DllImport("kernel32.dll", SetLastError = true)]
        private static extern bool QueryFullProcessImageName([In]IntPtr hProcess, [In]int dwFlags, [Out]StringBuilder lpExeName, ref int lpdwSize);

        [DllImport("kernel32.dll", SetLastError = true)]
        private static extern IntPtr OpenProcess(UInt32 dwDesiredAccess, [MarshalAs(UnmanagedType.Bool)]Boolean bInheritHandle, Int32 dwProcessId);

        private delegate bool EnumWindowProc(IntPtr hWnd, IntPtr parameter);
        delegate void WinEventDelegate(IntPtr hWinEventHook, uint eventType, IntPtr hwnd, int idObject, int idChild, uint dwEventThread, uint dwmsEventTime);

        private const uint PROCESS_QUERY_INFORMATION = 0x400;
        private const uint PROCESS_VM_READ = 0x010;

        private const uint WINEVENT_OUTOFCONTEXT = 0;
        private const uint EVENT_SYSTEM_FOREGROUND = 3;

        private static bool EnumChildWindowsCallback(IntPtr hWnd, IntPtr lParam)
        {
            WINDOWINFO info = (WINDOWINFO)Marshal.PtrToStructure(lParam, typeof(WINDOWINFO));

            uint pID;
            GetWindowThreadProcessId(hWnd, out pID);

            if (pID != info.ownerpid) info.childpid = pID;

            Marshal.StructureToPtr(info, lParam, true);

            return true;
        }

        private List<IntPtr> GetWindowHandles(string classname, IntPtr hParent, int maxCount)
        {
            //ManagementClass mngcls = new ManagementClass("Win32_Process");
            //foreach (ManagementObject instance in mngcls.GetInstances())
            //{
            //    Console.Write("ID: " + instance["ProcessId"]);
            //}

            List<IntPtr> result = new List<IntPtr>();
            int ct = 0;
            IntPtr prevChild = IntPtr.Zero;
            IntPtr currChild = IntPtr.Zero;
            while (ct < maxCount)
            {
                //currChild = FindWindowEx(hParent, prevChild, null, null);
                currChild = FindWindowEx(hParent, prevChild, classname, null);
                if (currChild == IntPtr.Zero) break;
                result.Add(currChild);
                prevChild = currChild;
                ct++;
            }
            return result;
        }

        private string UWP_AppName(IntPtr hWnd, uint pID)
        {
            WINDOWINFO windowinfo = new WINDOWINFO();
            windowinfo.ownerpid = pID;
            windowinfo.childpid = windowinfo.ownerpid;

            IntPtr pWindowinfo = Marshal.AllocHGlobal(Marshal.SizeOf(windowinfo));

            Marshal.StructureToPtr(windowinfo, pWindowinfo, false);

            EnumWindowProc lpEnumFunc = new EnumWindowProc(EnumChildWindowsCallback);
            EnumChildWindows(hWnd, lpEnumFunc, pWindowinfo);

            windowinfo = (WINDOWINFO)Marshal.PtrToStructure(pWindowinfo, typeof(WINDOWINFO));

            IntPtr proc;
            if ((proc = OpenProcess(PROCESS_QUERY_INFORMATION | PROCESS_VM_READ, false, (int)windowinfo.childpid)) == IntPtr.Zero) return null;

            int capacity = 2000;
            StringBuilder sb = new StringBuilder(capacity);
            QueryFullProcessImageName(proc, 0, sb, ref capacity);

            Marshal.FreeHGlobal(pWindowinfo);

            return sb.ToString(0, capacity);
        }

        private List<string> UWP_AppNames()
        {
            List<string> result = new List<string>();

            var UWPs = GetWindowHandles("Windows.UI.Core.CoreWindow", IntPtr.Zero, 100);
            foreach (var uwp in UWPs)
            {
                uint pid = 0;
                GetWindowThreadProcessId(uwp, out pid);
                var up = UWP_AppName(uwp, pid);
                if (result.Contains(up)) continue;
                result.Add(up);
            }

            return (result);
        }

        private List<string> UWP_AppNames(string title)
        {
            List<string> result = new List<string>();

            var UWPs = GetWindowHandles("Windows.UI.Core.CoreWindow", IntPtr.Zero, 100);
            foreach (var uwp in UWPs)
            {
                uint pid = 0;
                GetWindowThreadProcessId(uwp, out pid);
                StringBuilder sb = new StringBuilder(2000);
                var len = GetWindowText(uwp, sb, 2000);
                var ut = sb.ToString();
                var up = UWP_AppName(uwp, pid);
                if (result.Contains(up)) continue;
                if (string.IsNullOrEmpty(title) || title.Equals("*") ||
                    Regex.IsMatch(up, title, RegexOptions.IgnoreCase)||
                    Regex.IsMatch(ut, title, RegexOptions.IgnoreCase))
                {
                    result.Add(up);
                }
            }

            return (result);
        }

        private List<string> UWP_AppNames(uint pID)
        {
            List<string> result = new List<string>();

            var Hosts = GetWindowHandles("ApplicationFrameWindow", IntPtr.Zero, 100);
            foreach (var host in Hosts)
            {
                var childs = GetWindowHandles("Windows.UI.Core.CoreWindow", host, 100);
                if (childs.Count() > 0) {
                    var pname = UWP_AppName(host, pID);
                    if(!result.Contains(pname))
                        result.Add(UWP_AppName(host, pID));
                }
            }

            return (result);
        }

        #endregion
        #endregion

        #region Process routines
        private void ProcessStarted(object sender, EventArrivedEventArgs e)
        {
            // add proc to proc dict
            var procinfo = e.NewEvent.Properties;
            var pid = Convert.ToUInt32(procinfo["ProcessID"].Value);
            var proc = GetProcessById(pid);
            procs[pid] = proc;
        }

        private void ProcessStoped(object sender, EventArrivedEventArgs e)
        {
            // remove proc to proc dict
            var procinfo = e.NewEvent.Properties;
            var pid = Convert.ToUInt32(procinfo["ProcessID"].Value);
            procs.Remove(pid);
        }

        public void Affinity(int pid, int value = 0)
        {
            if (pid > 0 && IsAdmin)
            {
                Process proc = procs.ContainsKey((uint)pid) ? procs[(uint)pid] : GetProcessById(pid);
                if (proc is Process)
                {
                    proc.ProcessorAffinity = new IntPtr(value);
                }
            }
        }

        public void Affinity(uint pid, int value = 0)
        {
            if (pid > 0 && IsAdmin)
            {
                Process proc = procs.ContainsKey(pid) ? procs[pid] : GetProcessById(pid);
                if (proc is Process)
                {
                    proc.ProcessorAffinity = new IntPtr(value);
                }
            }
        }

        public Process GetProcessById(int pid)
        {
            Process result = null;

            var procs = Process.GetProcesses().Where(p => p.Id == pid);
            if (procs.Count() > 0)
            {
                result = procs.First();
            }

            return (result);
        }

        public Process GetProcessById(uint pid)
        {
            Process result = null;

            var procs = Process.GetProcesses().Where(p => p.Id == pid);
            if (procs.Count() > 0)
            {
                result = procs.First();
            }

            return (result);
        }

        public List<Process> GetProcessesByName(string name, bool regex = false)
        {
            List<Process> result = null;

            name = Path.GetFileNameWithoutExtension(name);
            if (regex)
                result = Process.GetProcesses().Where(p => Regex.IsMatch(p.ProcessName, $"{name}", RegexOptions.IgnoreCase)).ToList();
            else
                result = Process.GetProcesses().Where(p => p.ProcessName.Equals(name, StringComparison.OrdinalIgnoreCase)).ToList();

            return (result);
        }

        public List<Process> GetProcessesByName(string[] names, bool regex = false)
        {
            List<Process> result = new List<Process>();

            if (regex)
            {
                var namelist = $"({string.Join("|", names.Select(s => $"({Path.GetFileNameWithoutExtension(s)})"))})";
                result = Process.GetProcesses().Where(p => Regex.IsMatch(p.ProcessName, $"{namelist}", RegexOptions.IgnoreCase)).ToList();
            }
            else
            {
                result = Process.GetProcesses().Where(p => names.Contains(p.ProcessName, StringComparer.OrdinalIgnoreCase)).ToList();
            }

            return (result);
        }

        public List<Process> GetProcessesByTitle(string title, bool regex = false)
        {
            List<Process> result = new List<Process>();

            title = title.Trim('*');
            if (regex)
                result = Process.GetProcesses().Where(p => Regex.IsMatch(p.MainWindowTitle, $"{title}", RegexOptions.IgnoreCase)).ToList();
            else
                result = Process.GetProcesses().Where(p => p.MainWindowTitle.Equals(title, StringComparison.OrdinalIgnoreCase)).ToList();

            var Hosts = result.Where(p => p.ProcessName.Equals("ApplicationFrameHost", StringComparison.OrdinalIgnoreCase));
            if (Hosts.Count() > 0)
            {
                var proc = Hosts.First();
                //var pnames = UWP_AppNames((uint)proc.Id);
                var pnames = UWP_AppNames(title);
                foreach (var pname in pnames)
                {
                    var uwps = GetProcessesByName(pname, regex);
                    if (uwps.Count() > 0)
                    {
                        var app = uwps.First();
                        if (app.Id != proc.Id && result.Where(p => p.Id == app.Id).Count() == 0) result.Add(app);
                    }                        
                }
                result.Remove(proc);
            }

            return (result);
        }

        public List<Process> GetProcessesByTitle(string[] titles, bool regex = false)
        {
            List<Process> result = null;

            var namelist = $"({string.Join("|", titles.Select(s => $"({s.Trim('*')})"))})";
            if (regex)
            {
                result = Process.GetProcesses().Where(p => Regex.IsMatch(p.MainWindowTitle, $"{namelist}", RegexOptions.IgnoreCase)).ToList();
            }
            else
            {
                result = Process.GetProcesses().Where(p => titles.Contains(p.MainWindowTitle, StringComparer.OrdinalIgnoreCase)).ToList();
            }

            var Hosts = result.Where(p => p.ProcessName.Equals("ApplicationFrameHost", StringComparison.OrdinalIgnoreCase));
            if (Hosts.Count() > 0)
            {
                var proc = Hosts.First();
                //var pnames = UWP_AppNames((uint)proc.Id);
                var pnames = UWP_AppNames(namelist);
                foreach (var pname in pnames)
                {
                    var uwps = GetProcessesByName(pname, regex);
                    if (uwps.Count() > 0)
                    {
                        var app = uwps.First();
                        if (app.Id != proc.Id && result.Where(p => p.Id == app.Id).Count() == 0) result.Add(app);
                    }
                }
                result.Remove(proc);
            }

            return (result);
        }

        public string ProcessName(int pid)
        {
            string result = string.Empty;

            if (pid > 0)//&& IsAdmin)
            {
                Process proc = procs.ContainsKey((uint)pid) ? procs[(uint)pid] : GetProcessById(pid);
                if (proc is Process) result = proc.ProcessName;
            }

            return (result);
        }

        public string ProcessName(uint pid)
        {
            string result = string.Empty;

            if (pid > 0)//&& IsAdmin)
            {
                Process proc = procs.ContainsKey(pid) ? procs[pid] : GetProcessById(pid);
                if (proc is Process) result = proc.ProcessName;
            }

            return (result);
        }

        public string ProcessName(string title)
        {
            string result = string.Empty;

            if (!string.IsNullOrEmpty(title))
            {
                foreach (var proc in Process.GetProcesses())
                {
#if DEBUG
                    Debug.Print(title);
#endif
                    if (proc.SessionId == 0) continue;
                    var pname = proc.ProcessName;
                    if (pname.Equals("RuntimeBroker", StringComparison.CurrentCultureIgnoreCase)) continue;

                    var ptitle = proc.MainWindowTitle;
                    try
                    {
                        if (title.Equals(ptitle, StringComparison.InvariantCulture) ||
                            Regex.IsMatch(pname, $"{title}", RegexOptions.IgnoreCase) ||
                            (!string.IsNullOrEmpty(ptitle) && Regex.IsMatch(title, $"{ptitle.Trim('*')}", RegexOptions.IgnoreCase)) ||
                            Regex.IsMatch(ptitle, $"{title}", RegexOptions.IgnoreCase))
                        {
                            if (pname.Equals("ApplicationFrameHost", StringComparison.CurrentCultureIgnoreCase))
                                result = Path.GetFileNameWithoutExtension(UWP_AppName(proc.MainWindowHandle, (uint)proc.Id));
                            else
                                result = proc.ProcessName;
                            break;
                        }
                    }
                    catch (Exception) { }
                }
            }
            return (result);
        }

        #region Kill Process
        internal bool KillProcess(int pid)
        {
            bool result = true;
            try
            {
                if (pid > 0)
                {
                    Process ps = procs.ContainsKey((uint)pid) ? procs[(uint)pid] : GetProcessById(pid);
                    if (ps is Process) ps.Kill();
                }
                else result = false;
            }
            catch (ArgumentException) { }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Kill Failed!", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

            return (result);
        }

        internal bool KillProcess(uint pid)
        {
            bool result = true;
            try
            {
                if (pid > 0)
                {
                    Process ps = procs.ContainsKey(pid) ? procs[pid] : GetProcessById(pid);
                    if (ps is Process) ps.Kill();
                }
                else result = false;
            }
            catch (ArgumentException) { }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Kill Failed!", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

            return (result);
        }

        internal bool KillProcess(string processName)
        {
            bool result = true;

            try
            {
                //Process[] pslist = Process.GetProcesses();
                var pName = Path.GetFileNameWithoutExtension(processName);
                var pslist = GetProcessesByName(pName);
                foreach (var ps in pslist)
                {
                    try
                    {
                        if (ps.ProcessName.EndsWith(pName, StringComparison.CurrentCultureIgnoreCase))
                        {
                            ps.Kill();
                        }
                    }
                    catch (Exception) { }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

            return (result);
        }

        internal void KillProcess(string[] processNames)
        {
            foreach (var processName in processNames)
            {
                try
                {
                    var pid = AutoItX.ProcessExists(processName);
                    if (pid > 0)
                    {
                        KillProcess(pid);
                    }
                    //KillProcess(processName);
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }

        public void Kill(int pid)
        {
            KillProcess(pid);
        }

        public void Kill(uint pid)
        {
            KillProcess(pid);
        }

        public void Kill(string processName)
        {
            KillProcess(processName);
        }

        public void Kill(string[] processList)
        {
            KillProcess(processList);
        }
        #endregion

        #endregion

        #region Power routines
        #region Win32 Power On/Off Monitor
        private const int LCI_WM_SYSCommand = 274;
        private const int LCI_SC_MonitorPower = 61808;
        private const int LCI_Power_Off = 2;
        private const int LCI_Power_On = -1;

        //[DllImport("coredll.dll")]
        //private static extern IntPtr SendMessage(IntPtr hWnd, int Msg, IntPtr wParam, IntPtr lParam);

        [DllImport("user32.dll")]
        public static extern int SendMessage(int hWnd, uint Msg, int wParam, int lParam);

        internal static void Monitor(bool on, bool lockscreen=false)
        {
            var handle = AutoItX.WinGetHandle("[CLASS:Progman]");
            if (handle != IntPtr.Zero)
            {
                if (on)
                    SendMessage(handle.ToInt32(), LCI_WM_SYSCommand, LCI_SC_MonitorPower, LCI_Power_On);
                else
                    SendMessage(handle.ToInt32(), LCI_WM_SYSCommand, LCI_SC_MonitorPower, LCI_Power_Off);
            }
            if (lockscreen) AutoItX.Send("#l");
        }
        #endregion

        public void MonitorOn(bool lockscreen = false)
        {
            Monitor(true, lockscreen);
        }

        public void MonitorOff(bool lockscreen = false)
        {
            Monitor(false, lockscreen);
        }
        #endregion

        #region Media routines
        private bool InternalPlay = false;
        private NAudio.Wave.WaveOut waveOut = new NAudio.Wave.WaveOut();

        public void AppMute(MUTE_MODE mode, string app = default(string))
        {
            try
            {
                //AudioSessionManager asman = new AudioSessionManager();
                //NAudio.Mixer.num
                //if (NAudio.Mixer.Mixer.NumberOfDevices > 0)
                //{
                //    var mixer = new NAudio.Mixer.Mixer(0);
                //}

                if (string.IsNullOrEmpty(app))
                {
                    MMDevice maindev = new MMDeviceEnumerator().GetDefaultAudioEndpoint(DataFlow.Render, Role.Multimedia);
                    maindev.AudioEndpointVolume.Mute = true;
                }
                else
                {
                    //Instantiate an Enumerator to find audio devices
                    MMDeviceEnumerator MMDE = new MMDeviceEnumerator();
                    //Get all the devices, no matter what condition or status
                    MMDeviceCollection DevCol = MMDE.EnumerateAudioEndPoints(DataFlow.All, NAudio.CoreAudioApi.DeviceState.Active);
                    //Loop through all devices
                    foreach (MMDevice dev in DevCol)
                    {
                        try
                        {
#if DEBUG
                            Debug.Print(dev.FriendlyName);
#endif
                            var sessions = dev.AudioSessionManager.Sessions;
                            for (int i = 0; i < sessions.Count; i++)
                            {
                                var session = sessions[i];
                                var pid = session.GetProcessID;
                                if (pid == 0) continue;
                                var process = procs.ContainsKey(pid) ? procs[pid] : GetProcessById(pid);
                                if (process is Process)
                                {
                                    var title = process.MainWindowTitle;
                                    var pname = process.ProcessName;
                                    if (//session.State == NAudio.CoreAudioApi.Interfaces.AudioSessionState.AudioSessionStateActive &&
                                        !session.IsSystemSoundsSession &&
                                        (Regex.IsMatch(pname, app, RegexOptions.IgnoreCase) ||
                                        Regex.IsMatch(app, title, RegexOptions.IgnoreCase) ||
                                        Regex.IsMatch(title, app, RegexOptions.IgnoreCase)))
                                    {
                                        switch (mode)
                                        {
                                            case MUTE_MODE.Background:

                                                break;
                                            case MUTE_MODE.Mute:
                                                session.SimpleAudioVolume.Mute = true;
                                                break;
                                            case MUTE_MODE.UnMute:
                                                session.SimpleAudioVolume.Mute = false;
                                                break;
                                            case MUTE_MODE.Toggle:
                                                session.SimpleAudioVolume.Mute = !session.SimpleAudioVolume.Mute;
                                                break;
                                        }
                                    }
                                }
                            }
                        }
                        catch (Exception)
                        {
                            //Do something with exception when an audio endpoint could not be muted
                        }
                    }
                }
            }
            catch (Exception)
            {
                //When something happend that prevent us to iterate through the devices
            }
        }

        public void AppMute(string app = default(string))
        {
            AppMute(MUTE_MODE.Mute, app);
        }

        public void AppUnMute(string app = default(string))
        {
            AppMute(MUTE_MODE.UnMute, app);
        }

        public void AppToggleMute(string app = default(string))
        {
            AppMute(MUTE_MODE.Toggle, app);
        }

        public void AppBackgroundMute(string app = default(string))
        {
            AppMute(MUTE_MODE.Background, app);
        }

        public void Mute(MUTE_MODE mode, string device = default(string))
        {
            try
            {
                //AudioSessionManager asman = new AudioSessionManager();
                //NAudio.Mixer.num
                //if (NAudio.Mixer.Mixer.NumberOfDevices > 0)
                //{
                //    var mixer = new NAudio.Mixer.Mixer(0);
                //}

                if (string.IsNullOrEmpty(device))
                {
                    MMDevice maindev = new MMDeviceEnumerator().GetDefaultAudioEndpoint(DataFlow.All, Role.Multimedia);
                    maindev.AudioEndpointVolume.Mute = true;
                }
                else
                {
                    //Instantiate an Enumerator to find audio devices
                    MMDeviceEnumerator MMDE = new MMDeviceEnumerator();
                    //Get all the devices, no matter what condition or status
                    MMDeviceCollection DevCol = MMDE.EnumerateAudioEndPoints(DataFlow.All, NAudio.CoreAudioApi.DeviceState.Active);
                    //Loop through all devices
                    foreach (MMDevice dev in DevCol)
                    {
                        try
                        {
#if DEBUG
                            Debug.Print(dev.FriendlyName);
#endif
                            if (device.Equals("*") || dev.FriendlyName.Equals(device, StringComparison.CurrentCultureIgnoreCase))
                            {
                                //Mute it
                                switch (mode)
                                {
                                    case MUTE_MODE.Background:
                                        break;
                                    case MUTE_MODE.Mute:
                                        dev.AudioEndpointVolume.Mute = true;
                                        break;
                                    case MUTE_MODE.UnMute:
                                        dev.AudioEndpointVolume.Mute = false;
                                        break;
                                    case MUTE_MODE.Toggle:
                                        dev.AudioEndpointVolume.Mute = !dev.AudioEndpointVolume.Mute;
                                        break;
                                }
                            }
                        }
                        catch (Exception)
                        {
                            //Do something with exception when an audio endpoint could not be muted
                        }
                    }
                }
            }
            catch (Exception)
            {
                //When something happend that prevent us to iterate through the devices
            }
        }

        public void Mute(string device = default(string))
        {
            Mute(MUTE_MODE.Mute, device);
        }

        public void UnMute(string device = default(string))
        {
            Mute(MUTE_MODE.UnMute, device);
        }

        public void ToggleMute(string device = default(string))
        {
            Mute(MUTE_MODE.Toggle, device);
            //AutoItX.Send("{VOLUME_MUTE}");
            //AutoItX.Sleep(10);
        }

        public void BackgroundMute(string app = default(string))
        {
            Mute(MUTE_MODE.Background, app);
        }

        public bool Muted(string device = default(string))
        {
            bool result = false;

            try
            {
                if (string.IsNullOrEmpty(device))
                {
                    MMDevice maindev = new MMDeviceEnumerator().GetDefaultAudioEndpoint(DataFlow.All, Role.Multimedia);
                    if(maindev.AudioEndpointVolume.Mute) result = true;
                }
                else
                {
                    //Instantiate an Enumerator to find audio devices
                    MMDeviceEnumerator MMDE = new MMDeviceEnumerator();
                    //Get all the devices, no matter what condition or status
                    MMDeviceCollection DevCol = MMDE.EnumerateAudioEndPoints(DataFlow.All, NAudio.CoreAudioApi.DeviceState.Active);
                    //Loop through all devices
                    foreach (MMDevice dev in DevCol)
                    {
                        try
                        {
                            if (device.Equals("*") || dev.FriendlyName.Equals(device, StringComparison.CurrentCultureIgnoreCase))
                            {
                                if (dev.AudioEndpointVolume.Mute)
                                {
                                    result = true;
                                    break;
                                }
                            }
                        }
                        catch (Exception)
                        {
                            //Do something with exception when an audio endpoint could not be muted
                        }
                    }
                }
            }
            catch (Exception)
            {
                //When something happend that prevent us to iterate through the devices
            }

            return(result);
        }

        public bool AppMuted(string app = default(string))
        {
            bool result = false;

            try
            {
                if (string.IsNullOrEmpty(app))
                {
                    MMDevice maindev = new MMDeviceEnumerator().GetDefaultAudioEndpoint(DataFlow.Render, Role.Multimedia);
                    if(maindev.AudioEndpointVolume.Mute) result = true;
                }
                else
                {
                    //Instantiate an Enumerator to find audio devices
                    MMDeviceEnumerator MMDE = new MMDeviceEnumerator();
                    //Get all the devices, no matter what condition or status
                    MMDeviceCollection DevCol = MMDE.EnumerateAudioEndPoints(DataFlow.All, NAudio.CoreAudioApi.DeviceState.Active);
                    //Loop through all devices
                    foreach (MMDevice dev in DevCol)
                    {
                        try
                        {
                            var sessions = dev.AudioSessionManager.Sessions;
                            for (int i = 0; i < sessions.Count; i++)
                            {
                                var session = sessions[i];
                                var pid = session.GetProcessID;
                                if (pid == 0) continue;
                                var process = procs.ContainsKey(pid) ? procs[pid] : GetProcessById(pid);
                                if (process is Process)
                                {
                                    var title = process.MainWindowTitle;
                                    var pname = process.ProcessName;
                                    if (!session.IsSystemSoundsSession &&
                                        (Regex.IsMatch(pname, app, RegexOptions.IgnoreCase) ||
                                        Regex.IsMatch(app, title, RegexOptions.IgnoreCase) ||
                                        Regex.IsMatch(title, app, RegexOptions.IgnoreCase)))
                                    {
                                        if (session.SimpleAudioVolume.Mute)
                                        {
                                            result = true;
                                            break;
                                        }

                                    }
                                }
                            }
                        }
                        catch (Exception)
                        {
                            //Do something with exception when an audio endpoint could not be muted
                        }
                        if (result) break;
                    }
                }
            }
            catch (Exception)
            {
                //When something happend that prevent us to iterate through the devices
            }

            return (result);
        }

        private bool DeviceIsActive(MMDevice dev, string app = default(string))
        {
            bool result = false;

            try
            {
                var sessions = dev.AudioSessionManager.Sessions;
                for (int i = 0; i < sessions.Count; i++)
                {
                    var session = sessions[i];
                    var pid = session.GetProcessID;
                    if (pid <= 0) continue;
                    var process = procs.ContainsKey(pid) ? procs[pid] : GetProcessById(pid);
                    if (process is Process)
                    {
                        var title = process.MainWindowTitle;
                        var pname = process.ProcessName;
                        if (session.State == NAudio.CoreAudioApi.Interfaces.AudioSessionState.AudioSessionStateActive &&
                            !session.IsSystemSoundsSession && session.AudioMeterInformation.MasterPeakValue > 0 &&
                            (string.IsNullOrEmpty(app) || app.Equals("*") ||
                            pname.Equals(app, StringComparison.InvariantCulture) ||
                            Regex.IsMatch(pname, Path.GetFileNameWithoutExtension(app), RegexOptions.IgnoreCase) ||
                            Regex.IsMatch(title, app, RegexOptions.IgnoreCase)))
                        {
                            result = true;
                            break;
                        }
                    }
                }
            }
            catch (Exception)
            {
                //Do something with exception when an audio endpoint could not be muted
            }

            return (result);
        }

        private bool DeviceIsActive(MMDevice dev, string[] apps)
        {
            bool result = false;

            try
            {
                var sessions = dev.AudioSessionManager.Sessions;
                for (int i = 0; i < sessions.Count; i++)
                {
                    var session = sessions[i];
                    if (session.State == NAudio.CoreAudioApi.Interfaces.AudioSessionState.AudioSessionStateActive &&
                        !session.IsSystemSoundsSession)
                    {
                        if (!(apps is string[]) || apps.Length <= 0) result = true;
                        else
                        {
                            var pid = session.GetProcessID;
                            if (pid == 0) continue;
                            //var process = GetProcessById(pid);
                            var process = procs.ContainsKey(pid) ? procs[pid] : GetProcessById(pid);
                            if (process is Process)
                            {
                                var title = process.MainWindowTitle;
                                var pname = process.ProcessName;
                                foreach (var app in apps)
                                {
                                    if (string.IsNullOrEmpty(app) || app.Equals("*") ||
                                        Regex.IsMatch(pname, Path.GetFileNameWithoutExtension(app), RegexOptions.IgnoreCase) ||
                                        Regex.IsMatch(title, app, RegexOptions.IgnoreCase))
                                    {
                                        result = true;
                                        break;
                                    }
                                }
                            }
                        }
                    }
                    if (result) break;
                }
            }
            catch (Exception)
            {
                //Do something with exception when an audio endpoint could not be muted
            }

            return (result);
        }

        private bool DeviceIsActive(MMDevice dev, int pid = default(int))
        {
            bool result = false;

            try
            {
                var sessions = dev.AudioSessionManager.Sessions;
                for (int i = 0; i < sessions.Count; i++)
                {
                    var session = sessions[i];
                    var id = session.GetProcessID;
                    if (id <= 0) continue;
                    if (session.State == NAudio.CoreAudioApi.Interfaces.AudioSessionState.AudioSessionStateActive &&
                        !session.IsSystemSoundsSession && session.AudioMeterInformation.MasterPeakValue > 0 &&
                        pid > 0 && pid == id)
                    {
                        result = true;
                        break;
                    }
                }

            }
            catch (Exception)
            {
                //Do something with exception when an audio endpoint could not be muted
            }

            return (result);
        }

        private bool DeviceIsActive(MMDevice dev, int[] pids)
        {
            bool result = false;

            try
            {
                var sessions = dev.AudioSessionManager.Sessions;
                for (int i = 0; i < sessions.Count; i++)
                {
                    var session = sessions[i];
                    if (session.State == NAudio.CoreAudioApi.Interfaces.AudioSessionState.AudioSessionStateActive &&
                        !session.IsSystemSoundsSession)
                    {
                        if (!(pids is int[]) || pids.Length <= 0) result = true;
                        else
                        {
                            var pid = session.GetProcessID;
                            if (pid == 0) continue;
                            foreach (var id in pids)
                            {
                                if (id > 0 && id == pid)
                                {
                                    result = true;
                                    break;
                                }
                            }
                        }
                    }
                    if (result) break;
                }
            }
            catch (Exception)
            {
                //Do something with exception when an audio endpoint could not be muted
            }

            return (result);
        }

        private bool MediaIsActive(DataFlow mode, string app = default(string))
        {
            bool result = false;

            try
            {
                if (string.IsNullOrEmpty(app))
                {
                    MMDevice dev = new MMDeviceEnumerator().GetDefaultAudioEndpoint(mode, Role.Multimedia);
                    result = dev.State == NAudio.CoreAudioApi.DeviceState.Active;
                    //result = DeviceIsActive(dev, app);
                }
                else
                {
                    if (IsAdmin)
                        procs = Process.GetProcesses().ToDictionary(p => (uint)(p.Id), p => p);

                    //Instantiate an Enumerator to find audio devices
                    MMDeviceEnumerator MMDE = new MMDeviceEnumerator();
                    //Get all the devices, no matter what condition or status
                    MMDeviceCollection DevCol = MMDE.EnumerateAudioEndPoints(mode, NAudio.CoreAudioApi.DeviceState.Active);
                    //Loop through all devices
                    foreach (MMDevice dev in DevCol)
                    {
                        result = DeviceIsActive(dev, app);
                        if (result) break;
                    }
                }
            }
            catch (Exception)
            {
                //When something happend that prevent us to iterate through the devices
            }

            return (result);
        }

        private bool MediaIsActive(DataFlow mode, string[] apps)
        {
            bool result = false;

            try
            {
                if (!(apps is string[]) || apps.Length <= 0)
                {
                    MMDevice dev = new MMDeviceEnumerator().GetDefaultAudioEndpoint(mode, Role.Multimedia);
                    result = dev.State == NAudio.CoreAudioApi.DeviceState.Active;
                    //result = DeviceIsActive(dev, apps);
                }
                else
                {
                    if (IsAdmin)
                        procs = Process.GetProcesses().ToDictionary(p => (uint)(p.Id), p => p);
                    
                    //Instantiate an Enumerator to find audio devices
                    MMDeviceEnumerator MMDE = new MMDeviceEnumerator();
                    //Get all the devices, no matter what condition or status
                    MMDeviceCollection DevCol = MMDE.EnumerateAudioEndPoints(mode, NAudio.CoreAudioApi.DeviceState.Active);
                    //Loop through all devices
                    foreach (MMDevice dev in DevCol)
                    {
                        result = DeviceIsActive(dev, apps);
                        if (result) break;
                    }
                }
            }
            catch (Exception)
            {
                //When something happend that prevent us to iterate through the devices
            }

            return (result);
        }

        private bool MediaIsActive(DataFlow mode, int pid)
        {
            bool result = false;

            try
            {
                if (pid == 0)
                {
                    MMDevice dev = new MMDeviceEnumerator().GetDefaultAudioEndpoint(mode, Role.Multimedia);
                    result = dev.State == NAudio.CoreAudioApi.DeviceState.Active;
                    //result = DeviceIsActive(dev, app);
                }
                else
                {
                    //Instantiate an Enumerator to find audio devices
                    MMDeviceEnumerator MMDE = new MMDeviceEnumerator();
                    //Get all the devices, no matter what condition or status
                    MMDeviceCollection DevCol = MMDE.EnumerateAudioEndPoints(mode, NAudio.CoreAudioApi.DeviceState.Active);
                    //Loop through all devices
                    foreach (MMDevice dev in DevCol)
                    {
                        result = DeviceIsActive(dev, pid);
                        if (result) break;
                    }
                }
            }
            catch (Exception)
            {
                //When something happend that prevent us to iterate through the devices
            }

            return (result);
        }

        private bool MediaIsActive(DataFlow mode, int[] pids)
        {
            bool result = false;

            try
            {
                if (!(pids is int[]) || pids.Length <= 0)
                {
                    MMDevice dev = new MMDeviceEnumerator().GetDefaultAudioEndpoint(mode, Role.Multimedia);
                    result = dev.State == NAudio.CoreAudioApi.DeviceState.Active;
                    //result = DeviceIsActive(dev, apps);
                }
                else
                {
                    //Instantiate an Enumerator to find audio devices
                    MMDeviceEnumerator MMDE = new MMDeviceEnumerator();
                    //Get all the devices, no matter what condition or status
                    MMDeviceCollection DevCol = MMDE.EnumerateAudioEndPoints(mode, NAudio.CoreAudioApi.DeviceState.Active);
                    //Loop through all devices
                    foreach (MMDevice dev in DevCol)
                    {
                        result = DeviceIsActive(dev, pids);
                        if (result) break;
                    }
                }
            }
            catch (Exception)
            {
                //When something happend that prevent us to iterate through the devices
            }

            return (result);
        }

        public bool MediaIsOut(string app = default(string))
        {
            bool result = false;

            result = MediaIsActive(DataFlow.Render, app);

            return (result);
        }

        public bool MediaIsOut(string[] apps)
        {
            bool result = false;

            result = MediaIsActive(DataFlow.Render, apps);

            return (result);
        }

        public bool MediaIsOut(int pid)
        {
            bool result = false;

            result = MediaIsActive(DataFlow.Render, pid);

            return (result);
        }

        public bool MediaIsOut(int[] pids)
        {
            bool result = false;

            result = MediaIsActive(DataFlow.Render, pids);

            return (result);
        }

        public bool MediaIsIn(string app = default(string))
        {
            bool result = false;

            result = MediaIsActive(DataFlow.Capture, app);

            return (result);
        }

        public bool MediaIsIn(string[] apps)
        {
            bool result = false;

            result = MediaIsActive(DataFlow.Capture, apps);

            return (result);
        }

        public bool MediaIsIn(int pid)
        {
            bool result = false;

            result = MediaIsActive(DataFlow.Capture, pid);

            return (result);
        }

        public bool MediaIsIn(int[] pids)
        {
            bool result = false;

            result = MediaIsActive(DataFlow.Capture, pids);

            return (result);
        }

        public void MediaPlay(string media= default(string))
        {
            if (string.IsNullOrEmpty(media))
            {
                AutoItX.Send("{MEDIA_STOP}");
                AutoItX.Send("{MEDIA_PLAY_PAUSE}");
                InternalPlay = false;
            }
            else
            {
                if (waveOut.PlaybackState != NAudio.Wave.PlaybackState.Stopped) waveOut.Stop();
                if (File.Exists(media) && waveOut is NAudio.Wave.WaveOut)
                {
                    NAudio.Wave.WaveStream audio = new NAudio.Wave.AudioFileReader(media);
                    if (audio is NAudio.Wave.WaveStream)
                    {
                        waveOut.Init(audio);
                        waveOut.Play();
                        InternalPlay = true;
                    }
                }
            }
        }

        public void MediaTogglePause()
        {
            if (!InternalPlay)
                AutoItX.Send("{MEDIA_PLAY_PAUSE}");
            else
            {
                if (waveOut is NAudio.Wave.WaveOut)
                {
                    if (waveOut.PlaybackState == NAudio.Wave.PlaybackState.Playing) waveOut.Pause();
                }
            }
        }

        public void MediaStop()
        {
            if (!InternalPlay)
                AutoItX.Send("{MEDIA_STOP}");
            else
            {
                if(waveOut is NAudio.Wave.WaveOut)
                {
                    if (waveOut.PlaybackState != NAudio.Wave.PlaybackState.Stopped) waveOut.Stop();
                    InternalPlay = false;
                }
            }
        }

        public void MediaToggleMute()
        {
            AutoItX.Send("{VOLUME_MUTE}");
        }
        #endregion

        #region Misc
        public void Beep(string type = default(string))
        {
            if (string.IsNullOrEmpty(type))
                System.Media.SystemSounds.Beep.Play();
            else if (type.Equals("*"))
                System.Media.SystemSounds.Asterisk.Play();
            else if (type.Equals("!"))
                System.Media.SystemSounds.Exclamation.Play();
            else if (type.Equals("?"))
                System.Media.SystemSounds.Question.Play();
            else if (type.Equals("h", StringComparison.CurrentCultureIgnoreCase))
                System.Media.SystemSounds.Hand.Play();
            else
                System.Media.SystemSounds.Beep.Play();
        }

        public void Sleep(int ms)
        {
            AutoItX.Sleep(ms);
        }

        private System.Speech.Synthesis.SpeechSynthesizer synth = new System.Speech.Synthesis.SpeechSynthesizer() { };
        private string voice_default = string.Empty;
        public void Speak(string text, int vol=100, int rate=0)
        {
            List<string> lang_cn = new List<string>() { "zh-hans", "zh-cn", "zh" };
            List<string> lang_tw = new List<string>() { "zh-hant", "zh-tw" };
            List<string> lang_jp = new List<string>() { "ja-jp", "ja", "jp" };
            List<string> lang_en = new List<string>() { "en-us", "us", "en" };

            try
            {
                if(string.IsNullOrEmpty(voice_default)) voice_default = synth.Voice.Name;

                synth.SelectVoice(voice_default);
                //
                // 中文：[\u4e00-\u9fcc, \u3400-\u4db5, \u20000-\u2a6d6, \u2a700-\u2b734, \u2b740-\u2b81d, \uf900-\ufad9, \u2f800-\u2fa1d]
                // 日文：[\u0800-\u4e00] [\u3041-\u31ff]
                // 韩文：[\uac00-\ud7ff]
                //
                //var m_jp = Regex.Matches(text, @"([\u0800-\u4e00])", RegexOptions.Multiline | RegexOptions.IgnoreCase);
                //var m_zh = Regex.Matches(text, @"([\u4e00-\u9fbb])", RegexOptions.Multiline | RegexOptions.IgnoreCase);

                var lang = System.Globalization.CultureInfo.CurrentCulture.IetfLanguageTag.ToLower();

                if (Regex.Matches(text, @"[\u3041-\u31ff]", RegexOptions.Multiline).Count > 0)
                {
                    lang = "ja";
                }
                else if (Regex.Matches(text, @"[\u4e00-\u9fbb]", RegexOptions.Multiline).Count > 0)
                {
                    lang = "zh";
                }

                #region Detect language
                // Initialize a new instance of the SpeechSynthesizer.
                foreach (System.Speech.Synthesis.InstalledVoice voice in synth.GetInstalledVoices())
                {
                    System.Speech.Synthesis.VoiceInfo info = voice.VoiceInfo;
                    var vl = info.Culture.IetfLanguageTag;

                    if (lang_cn.Contains(vl.ToLower()) &&
                        lang.StartsWith("zh", StringComparison.CurrentCultureIgnoreCase) &&
                        voice.VoiceInfo.Name.ToLower().Contains("huihui"))
                    {
                        synth.SelectVoice(voice.VoiceInfo.Name);
                        break;
                    }
                    else if (lang_jp.Contains(vl.ToLower()) &&
                        lang.StartsWith("ja", StringComparison.CurrentCultureIgnoreCase) &&
                        voice.VoiceInfo.Name.ToLower().Contains("haruka"))
                    {
                        synth.SelectVoice(voice.VoiceInfo.Name);
                        break;
                    }
                    else if (lang_en.Contains(vl.ToLower()) &&
                        lang.StartsWith("en", StringComparison.CurrentCultureIgnoreCase) &&
                        voice.VoiceInfo.Name.ToLower().Contains("zira"))
                    {
                        synth.SelectVoice(voice.VoiceInfo.Name);
                        break;
                    }
                }
                #endregion

                // Synchronous
                //synth.Speak( text );
                // Asynchronous
                synth.Volume = Math.Min(100, Math.Max(0, vol));
                synth.Rate = Math.Min(10, Math.Max(-10, rate));
                synth.SpeakAsyncCancelAll();
                synth.Resume();
                synth.SpeakAsync(text);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1, MessageBoxOptions.ServiceNotification);
            }
        }

        private List<string> logger = new List<string>();
        public List<string> Logger { get { return(logger); } }
        public void Print(string text)
        {
            logger.Add(text);
        }
        #endregion
    }
}
