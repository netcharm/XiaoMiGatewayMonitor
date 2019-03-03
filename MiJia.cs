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
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows.Forms;

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
                        Devices[device.Name] = new DEVICE() { client = client, Info = device };
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
        private Dictionary<string, DEVICE> Devices = new Dictionary<string, DEVICE>();

        private async void DeviceStateChanged(object sender, StateChangedEventArgs e)
        {
            if (Devices.ContainsKey(e.Device.Name) && Devices[e.Device.Name] is DEVICE)
            {
                Devices[e.Device.Name].State = e.NewData;
                Devices[e.Device.Name].StateName = e.StateName;
                Devices[e.Device.Name].Info = e.Device;
                Devices[e.Device.Name].StateDuration = 0;
            }
            else
                Devices[e.Device.Name] = new DEVICE() { client = client, State = e.NewData, StateName = e.StateName, Info = e.Device };

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
            timerRefresh.Interval = 1000;
            timerRefresh.Tick += TimerRefresh_Tick;
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
                    Assembly.GetAssembly(typeof(System.Dynamic.DynamicObject)),  // System.Code
                    Assembly.GetAssembly(typeof(Microsoft.CSharp.RuntimeBinder.CSharpArgumentInfo)),  // Microsoft.CSharp
                    Assembly.GetAssembly(typeof(System.Dynamic.ExpandoObject))  // System.Dynamic                    
                });
            scriptOptions = scriptOptions.AddImports(new string[] {
                    "System",
                    "System.Dynamic",
                    "System.Globalization",
                    "System.IO",
                    "System.Math",
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

                    //result = await CSharpScript.RunAsync(ScriptContext, scriptOptions, globals);
                    if (!(script is Script)) Load();
                    result = await script.RunAsync(globals, cancelToken);

                    if (AutoReset) globals.Reset();

                    StringBuilder sb = new StringBuilder();
                    sb.AppendLine("".PaddingRight(72, '-'));// "--------------------------------");
                    foreach (var v in result.Variables)
                    {
                        if (v.Name.Equals("Device", StringComparison.CurrentCultureIgnoreCase)) continue;
                        sb.AppendLine($"{v.Name} = {v.Value}");
                        globals.vars[v.Name] = v.Value;
                    }
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
        public AqaraDevice Info { get; set; } = default(AqaraDevice);

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
                client.SendWriteCommand(Info, states);
            }
        }

        public void Reset()
        {
            State = string.Empty;
            StateName = string.Empty;
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

        #region MiJia Gateway/ZigBee Device
        internal Dictionary<string, DEVICE> device = new Dictionary<string, DEVICE>();
        public Dictionary<string, DEVICE> Device { get { return (device); } }

        public void Reset(string device = default(string))
        {
            foreach (var dev in Device)
            {
                if (string.IsNullOrEmpty(device) || device.Equals("*") || device.Equals(dev.Key, StringComparison.InvariantCulture))
                    dev.Value.Reset();
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
                //if (AutoItX.IsAdmin() == 1)
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
                    //if(p.MainModule.FileName)
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

        [DllImport("user32.dll", SetLastError = true)]
        public static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint lpdwProcessId);

        [DllImport("user32", SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        public static extern bool EnumChildWindows(IntPtr hWndParent, EnumWindowProc lpEnumFunc, IntPtr lParam);

        [DllImport("user32.dll")]
        static extern IntPtr SetWinEventHook(uint eventMin, uint eventMax, IntPtr hmodWinEventProc, WinEventDelegate lpfnWinEventProc, uint idProcess, uint idThread, uint dwFlags);

        [DllImport("kernel32.dll", SetLastError = true)]
        public static extern bool QueryFullProcessImageName([In]IntPtr hProcess, [In]int dwFlags, [Out]StringBuilder lpExeName, ref int lpdwSize);

        [DllImport("kernel32.dll", SetLastError = true)]
        public static extern IntPtr OpenProcess(UInt32 dwDesiredAccess, [MarshalAs(UnmanagedType.Bool)]Boolean bInheritHandle, Int32 dwProcessId);

        public delegate bool EnumWindowProc(IntPtr hWnd, IntPtr parameter);
        delegate void WinEventDelegate(IntPtr hWinEventHook, uint eventType, IntPtr hwnd, int idObject, int idChild, uint dwEventThread, uint dwmsEventTime);

        public const uint PROCESS_QUERY_INFORMATION = 0x400;
        public const uint PROCESS_VM_READ = 0x010;

        private const uint WINEVENT_OUTOFCONTEXT = 0;
        private const uint EVENT_SYSTEM_FOREGROUND = 3;

        private static string UWP_AppName(IntPtr hWnd, uint pID)
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

        private static bool EnumChildWindowsCallback(IntPtr hWnd, IntPtr lParam)
        {
            WINDOWINFO info = (WINDOWINFO)Marshal.PtrToStructure(lParam, typeof(WINDOWINFO));

            uint pID;
            GetWindowThreadProcessId(hWnd, out pID);

            if (pID != info.ownerpid) info.childpid = pID;

            Marshal.StructureToPtr(info, lParam, true);

            return true;
        }
        #endregion
        #endregion

        #region Process routines
        public void Affinity(int pid, int value = 0)
        {
            if (pid > 0 && AutoItX.IsAdmin() == 1)
            {
                Process proc = Process.GetProcessById(pid);
                proc.ProcessorAffinity = new IntPtr(value);
            }
        }

        public string ProcessName(int pid)
        {
            string result = string.Empty;

            if (pid > 0)//&& AutoItX.IsAdmin() == 1)
            {
                Process proc = Process.GetProcessById(pid);
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
                    //if (proc.MainWindowHandle == IntPtr.Zero) continue;
                    if (proc.SessionId == 0) continue;
                    var pname = proc.ProcessName;
                    if (pname.Equals("RuntimeBroker", StringComparison.CurrentCultureIgnoreCase)) continue;

                    var ptitle = proc.MainWindowTitle;
                    //if (string.IsNullOrEmpty(ptitle)) continue;
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
        internal static bool KillProcess(int pid)
        {
            bool result = true;
            try
            {
                if (pid > 0)
                {
                    //AutoItX.Sleep(10);
                    Process ps = Process.GetProcessById(pid);
                    if (ps is Process)
                        ps.Kill();
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

        internal static bool KillProcess(string processName)
        {
            bool result = true;

            try
            {
                //Process[] pslist = Process.GetProcesses();
                var pName = Path.GetFileNameWithoutExtension(processName);
                Process[] pslist = Process.GetProcessesByName(pName);
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
                foreach(var process in Process.GetProcesses())
                {
                    try
                    {
                        var pname = process.ProcessName;
                        var ptitle = process.MainWindowTitle;
                        if (Regex.IsMatch(pname, pName, RegexOptions.IgnoreCase) ||
                           Regex.IsMatch(ptitle, pName, RegexOptions.IgnoreCase))
                        {
                            process.Kill();
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

        internal static void KillProcess(string[] processNames)
        {
            foreach (var processName in processNames)
            {
                try
                {
                    var pid = AutoItX.ProcessExists(processName);
                    if (pid > 0)
                    {
                        KillProcess(pid);
                        //var ret = AutoItX.ProcessClose(processName);
                        //if (ret > 0) KillProcess(pid);
                        //if (ret > 0) KillProcess(processName);
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

        public void MonitorOn()
        {
            Monitor(true);
        }

        public void MonitorOff()
        {
            Monitor(false);
        }
        #endregion

        #region Media routines
        private bool InternalPlay = false;
        private NAudio.Wave.WaveOut waveOut = new NAudio.Wave.WaveOut();

        public void MuteApp(MUTE_MODE mode, string app = default(string))
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
                    MMDeviceCollection DevCol = MMDE.EnumerateAudioEndPoints(DataFlow.All, NAudio.CoreAudioApi.DeviceState.All);
                    //Loop through all devices
                    foreach (MMDevice dev in DevCol)
                    {
                        try
                        {
                            if (dev.State == NAudio.CoreAudioApi.DeviceState.Active)
                            {
                                //Show us the human understandable name of the device
#if DEBUG
                                Debug.Print(dev.FriendlyName);
#endif
                                var sessions = dev.AudioSessionManager.Sessions;
                                for (int i = 0; i < sessions.Count; i++)
                                {
                                    var session = sessions[i];
                                    var pid = session.GetProcessID;
                                    if (pid == 0) continue;
                                    var process = Process.GetProcessById((int)pid);
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

        public void MuteApp(string app = default(string))
        {
            MuteApp(MUTE_MODE.Mute, app);
        }

        public void UnMuteApp(string app = default(string))
        {
            MuteApp(MUTE_MODE.UnMute, app);
        }

        public void ToggleMuteApp(string app = default(string))
        {
            MuteApp(MUTE_MODE.Toggle, app);
        }

        public void BackgroundAppMute(string app = default(string))
        {
            MuteApp(MUTE_MODE.Background, app);
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
                    MMDevice maindev = new MMDeviceEnumerator().GetDefaultAudioEndpoint(DataFlow.Render, Role.Multimedia);
                    maindev.AudioEndpointVolume.Mute = true;
                }
                else
                {
                    //Instantiate an Enumerator to find audio devices
                    MMDeviceEnumerator MMDE = new MMDeviceEnumerator();
                    //Get all the devices, no matter what condition or status
                    MMDeviceCollection DevCol = MMDE.EnumerateAudioEndPoints(DataFlow.All, NAudio.CoreAudioApi.DeviceState.All);
                    //Loop through all devices
                    foreach (MMDevice dev in DevCol)
                    {
                        try
                        {
                            if (dev.State == NAudio.CoreAudioApi.DeviceState.Active)
                            {
                                //Show us the human understandable name of the device
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

        private bool DeviceIsActive(MMDevice dev, string app = default(string))
        {
            bool result = false;

            try
            {
                if (dev.State == NAudio.CoreAudioApi.DeviceState.Active)
                {
                    //Show us the human understandable name of the device
#if DEBUG
                    //Debug.Print(dev.FriendlyName);
#endif
                    var sessions = dev.AudioSessionManager.Sessions;
                    for (int i = 0; i < sessions.Count; i++)
                    {
                        var session = sessions[i];
                        var pid = session.GetProcessID;
                        if (pid == 0) continue;
                        var process = Process.GetProcessById((int)pid, Environment.MachineName);
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
                if (dev.State == NAudio.CoreAudioApi.DeviceState.Active)
                {
                    //Show us the human understandable name of the device
#if DEBUG
                    Debug.Print(dev.FriendlyName);
#endif
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
                                var process = Process.GetProcessById((int)pid, Environment.MachineName);
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
            }
            catch (Exception)
            {
                //Do something with exception when an audio endpoint could not be muted
            }

            return (result);
        }

        private bool DeviceIsActive(MMDevice dev, int app = default(int))
        {
            bool result = false;

            try
            {
                if (dev.State == NAudio.CoreAudioApi.DeviceState.Active)
                {
                    //Show us the human understandable name of the device
#if DEBUG
                    //Debug.Print(dev.FriendlyName);
#endif
                    var sessions = dev.AudioSessionManager.Sessions;
                    for (int i = 0; i < sessions.Count; i++)
                    {
                        var session = sessions[i];
                        var pid = session.GetProcessID;
                        if (pid == 0) continue;
                        if (session.State == NAudio.CoreAudioApi.Interfaces.AudioSessionState.AudioSessionStateActive &&
                            !session.IsSystemSoundsSession && session.AudioMeterInformation.MasterPeakValue > 0 &&
                            app > 0 && app == pid)
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

        private bool DeviceIsActive(MMDevice dev, int[] apps)
        {
            bool result = false;

            try
            {
                if (dev.State == NAudio.CoreAudioApi.DeviceState.Active)
                {
                    //Show us the human understandable name of the device
#if DEBUG
                    //Debug.Print(dev.FriendlyName);
#endif
                    var sessions = dev.AudioSessionManager.Sessions;
                    for (int i = 0; i < sessions.Count; i++)
                    {
                        var session = sessions[i];
                        if (session.State == NAudio.CoreAudioApi.Interfaces.AudioSessionState.AudioSessionStateActive &&
                            !session.IsSystemSoundsSession)
                        {
                            if (!(apps is int[]) || apps.Length <= 0) result = true;
                            else
                            {
                                var pid = session.GetProcessID;
                                if (pid == 0) continue;
                                foreach (var app in apps)
                                {
                                    if (app > 0 && app == pid)
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
            }
            catch (Exception)
            {
                //Do something with exception when an audio endpoint could not be muted
            }

            return (result);
        }

        public bool MediaIsActive(string app = default(string))
        {
            bool result = false;

            try
            {
                if (string.IsNullOrEmpty(app))
                {
                    MMDevice dev = new MMDeviceEnumerator().GetDefaultAudioEndpoint(DataFlow.Render, Role.Multimedia);
                    //if (dev.State == NAudio.CoreAudioApi.DeviceState.Active) result = true;
                    result = DeviceIsActive(dev, app);
                }
                else
                {
                    //Instantiate an Enumerator to find audio devices
                    MMDeviceEnumerator MMDE = new MMDeviceEnumerator();
                    //Get all the devices, no matter what condition or status
                    MMDeviceCollection DevCol = MMDE.EnumerateAudioEndPoints(DataFlow.All, NAudio.CoreAudioApi.DeviceState.All);
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

        public bool MediaIsActive(string[] apps)
        {
            bool result = false;

            try
            {
                if (!(apps is string[]) || apps.Length <= 0)
                {
                    MMDevice dev = new MMDeviceEnumerator().GetDefaultAudioEndpoint(DataFlow.Render, Role.Multimedia);
                    //if (dev.State == NAudio.CoreAudioApi.DeviceState.Active) result = true;
                    result = DeviceIsActive(dev, apps);
                }
                else
                {
                    //Instantiate an Enumerator to find audio devices
                    MMDeviceEnumerator MMDE = new MMDeviceEnumerator();
                    //Get all the devices, no matter what condition or status
                    MMDeviceCollection DevCol = MMDE.EnumerateAudioEndPoints(DataFlow.All, NAudio.CoreAudioApi.DeviceState.All);
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

        public bool MediaIsActive(int app)
        {
            bool result = false;

            try
            {
                if (app == 0)
                {
                    MMDevice dev = new MMDeviceEnumerator().GetDefaultAudioEndpoint(DataFlow.Render, Role.Multimedia);
                    //if (dev.State == NAudio.CoreAudioApi.DeviceState.Active) result = true;
                    result = DeviceIsActive(dev, app);
                }
                else
                {
                    //Instantiate an Enumerator to find audio devices
                    MMDeviceEnumerator MMDE = new MMDeviceEnumerator();
                    //Get all the devices, no matter what condition or status
                    MMDeviceCollection DevCol = MMDE.EnumerateAudioEndPoints(DataFlow.All, NAudio.CoreAudioApi.DeviceState.All);
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

        public bool MediaIsActive(int[] apps)
        {
            bool result = false;

            try
            {
                if (!(apps is int[]) || apps.Length <= 0)
                {
                    MMDevice dev = new MMDeviceEnumerator().GetDefaultAudioEndpoint(DataFlow.Render, Role.Multimedia);
                    //if (dev.State == NAudio.CoreAudioApi.DeviceState.Active) result = true;
                    result = DeviceIsActive(dev, apps);
                }
                else
                {
                    //Instantiate an Enumerator to find audio devices
                    MMDeviceEnumerator MMDE = new MMDeviceEnumerator();
                    //Get all the devices, no matter what condition or status
                    MMDeviceCollection DevCol = MMDE.EnumerateAudioEndPoints(DataFlow.All, NAudio.CoreAudioApi.DeviceState.All);
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

        public void Beep(string type= default(string))
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
        #endregion

        #region Misc
        public void Sleep(int ms)
        {
            AutoItX.Sleep(ms);
        }
        #endregion
    }
}
