using System;
using System.Collections.Generic;
using System.Collections.Specialized;
using System.Diagnostics;
using System.Drawing;
using System.Dynamic;
using System.IO;
using System.Linq;
using System.Linq.Expressions;
using System.Management;
using System.Net;
using System.Net.Sockets;
using System.Reflection;
using System.Runtime.InteropServices;
using System.ServiceProcess;
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

    public enum OpConditionMode { AND, OR, NOR, XOR };
    public class OpCondition<T>
    {
        public OpConditionMode Mode;
        public KeyValuePair<string, T> Param { get; set; }
    }

    public enum ActionMode { Close, Minimize, Maximize, Mute };
    public class OpAction<T>
    {
        public string Name { get; set; }
        public ActionMode Mode { get; set; }
        public IList<OpCondition<T>> Conditions { get; set; }
        public IList<string> Param { get; set; }
    }

    public static class Extensions
    {
        private static NLog.Logger log = NLog.LogManager.GetCurrentClassLogger();

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
            control.Invoke((MethodInvoker)delegate
            {
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

            if (device is AqaraDevice)
            {
                if (device.States.ContainsKey("state"))
                {

                }
            }

            return (result);
        }

        #endregion

        public static IEnumerable<T> Tail<T>(this IEnumerable<T> source, int N)
        {
            return source.Skip(Math.Max(0, source.Count() - N));
        }

        public static IEnumerable<T> Head<T>(this IEnumerable<T> source, int N)
        {
            return source.Take(Math.Max(0, Math.Min(N, source.Count())));
        }

        #region Process Helper
        public static void Kill(this Process process)
        {
            if (process is Process && process.Id > 0)
            {
                try
                {
                    var result = process.CloseMainWindow();
                    process.Close();
                    result = process.WaitForExit(5000);
                    if (!result && process.Id > 0)
                    {
                        process.Kill();
                        result = process.HasExited;
                    }
                }
                catch (Exception)
                {

                }
            }
        }

        public static void Restart(this Process process)
        {
            if (process is Process && process.Id > 0)
            {
                try
                {
                    var cmd = process.StartInfo;
                    Kill(process);
                    Process.Start(cmd);
                }
                catch (Exception)
                {

                }
            }
        }
        #endregion

        #region Service Helper
        public static void Start(this ServiceController service, double timeout = 30)
        {
            if (service is ServiceController &&
                (service.Status == ServiceControllerStatus.Stopped || service.Status == ServiceControllerStatus.StopPending))
            {
                try
                {
                    service.Start();
                    service.WaitForStatus(ServiceControllerStatus.Running, TimeSpan.FromSeconds(timeout));
                }
#if DEBUG
                catch (Exception ex)
#else
                catch (Exception)
#endif
                {
#if DEBUG
                    MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
#endif
                }
            }
        }

        public static void Start(this IEnumerable<ServiceController> services, double timeout = 30)
        {
            foreach (var service in services)
            {
                Start(service, timeout);
            }
        }

        public static void Start(this Dictionary<string, ServiceController> services, double timeout = 30)
        {
            foreach (var service in services)
            {
                Start(service.Value, timeout);
            }
        }

        public static void Stop(this ServiceController service, double timeout = 30)
        {
            if (service is ServiceController && service.CanStop &&
                service.Status != ServiceControllerStatus.Stopped && service.Status != ServiceControllerStatus.StopPending)
            {
                try
                {
                    service.Stop();
                    service.WaitForStatus(ServiceControllerStatus.Stopped, TimeSpan.FromSeconds(timeout));
                }
#if DEBUG
                catch (Exception ex)
#else
                catch (Exception)
#endif
                {
#if DEBUG
                    MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
#endif
                }
            }
        }

        public static void Stop(this IEnumerable<ServiceController> services, double timeout = 30)
        {
            foreach (var service in services)
            {
                Stop(service, timeout);
            }
        }

        public static void Stop(this Dictionary<string, ServiceController> services, double timeout = 30)
        {
            foreach (var service in services)
            {
                Stop(service.Value, timeout);
            }
        }

        public static void Restart(this ServiceController service, bool force = false, double timeout = 30)
        {
            if (service is ServiceController && service.CanStop)
            {
                try
                {
                    bool running = service.Status == ServiceControllerStatus.Running;
                    Stop(service, timeout);
                    if (running || force)
                        Start(service, timeout);
                }
#if DEBUG
                catch (Exception ex)
#else
                catch (Exception)
#endif
                {
#if DEBUG
                    MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
#endif
                }
            }
        }

        public static void Restart(this IEnumerable<ServiceController> services, bool force = false, double timeout = 30)
        {
            foreach (var service in services)
            {
                Restart(service, force, timeout);
            }
        }

        public static void Restart(this Dictionary<string, ServiceController> services, bool force = false, double timeout = 30)
        {
            foreach (var service in services)
            {
                Restart(service.Value, force, timeout);
            }
        }

        public static void Pause(this ServiceController service, double timeout = 30)
        {
            if (service is ServiceController && service.CanPauseAndContinue && service.Status == ServiceControllerStatus.Running)
            {
                try
                {
                    service.Pause();
                    service.WaitForStatus(ServiceControllerStatus.Paused, TimeSpan.FromSeconds(timeout));
                }
#if DEBUG
                catch (Exception ex)
#else
                catch (Exception)
#endif
                {
#if DEBUG
                    MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
#endif
                }
            }
        }

        public static void Pause(this IEnumerable<ServiceController> services, double timeout = 30)
        {
            foreach (var service in services)
            {
                Pause(service, timeout);
            }
        }

        public static void Pause(this Dictionary<string, ServiceController> services, double timeout = 30)
        {
            foreach (var service in services)
            {
                Pause(service.Value, timeout);
            }
        }

        public static void Continue(this ServiceController service, double timeout = 30)
        {
            if (service is ServiceController && service.CanPauseAndContinue &&
                (service.Status == ServiceControllerStatus.Paused || service.Status == ServiceControllerStatus.PausePending))
            {
                try
                {
                    service.Continue();
                    service.WaitForStatus(ServiceControllerStatus.Running, TimeSpan.FromSeconds(timeout));
                }
#if DEBUG
                catch (Exception ex)
#else
                catch (Exception)
#endif
                {
#if DEBUG
                    MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
#endif
                }
            }
        }

        public static void Continue(this IEnumerable<ServiceController> services, double timeout = 30)
        {
            foreach (var service in services)
            {
                Continue(service, timeout);
            }
        }

        public static void Continue(this Dictionary<string, ServiceController> services, double timeout = 30)
        {
            foreach (var service in services)
            {
                Continue(service.Value, timeout);
            }
        }

        #endregion

    }

    public class ScriptEngine
    {
        private static NLog.Logger log = NLog.LogManager.GetCurrentClassLogger();

        private static string APPFOLDER = Path.GetDirectoryName(Application.ExecutablePath);
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
                    if (timerRefresh.Enabled) timerRefresh.Stop();
                    timerRefresh.Interval = value;
                    timerRefresh.Start();
                }
            }
        }

        private async void TimerRefresh_Tick(object sender, EventArgs e)
        {
            if (sciptRunning || string.IsNullOrEmpty(scriptContext))
                return;
            else if (!(client is AqaraClient) || client.Gateways.Count() <= 0)
                return;

            try
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
            }
            catch (Exception) { }

            if (!Pausing) await RunScript();
        }

        private Task worker = null;
        private System.Threading.CancellationToken workerCancelToken = new System.Threading.CancellationToken();
        private AqaraConfig config = null;
        private AqaraClient client = null;
        private Dictionary<string, dynamic> Devices = new Dictionary<string, dynamic>();
        //private SortedDictionary<DateTime, StateChangedEventArgs> events = new SortedDictionary<DateTime, StateChangedEventArgs>();
        private Queue<KeyValuePair<DateTime, StateChangedEventArgs>> events = new Queue<KeyValuePair<DateTime, StateChangedEventArgs>>(100);

        private async void DeviceStateChanged(object sender, StateChangedEventArgs e)
        {
            //events.Add(DateTime.Now, e);
            if (events is Queue<KeyValuePair<DateTime, StateChangedEventArgs>>)
            {
                events.Enqueue(new KeyValuePair<DateTime, StateChangedEventArgs>(DateTime.Now, e));
                if (events.Count > 100) events.Dequeue();
            }
            if (Devices is Dictionary<string, dynamic>)
            {
                Devices[e.Device.Name] = e.Device;
                if (e.Device is AqaraDevice)
                {
                    Devices[e.Device.Name].NewStateName = e.StateName;
                    Devices[e.Device.Name].NewStateValue = e.NewData;
                    Devices[e.Device.Name].StateDuration = 0;
                }
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

            if (client is AqaraClient)
            {
                client.CancellationPending = true;
                AutoItX.Sleep(100);
                if (worker is Task) workerCancelToken.ThrowIfCancellationRequested();
                AutoItX.Sleep(100);
            }

            client = new AqaraClient(config);
            client.DeviceStateChanged += DeviceStateChanged;
            worker = Task.Run(() =>
            {
                client.DoWork(workerCancelToken);
            }, workerCancelToken);

            if (timerRefresh is Timer) timerRefresh.Stop();
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

        private string scriptFile = string.Empty;
        public string ScriptFile
        {
            get { return (scriptFile); }
            set
            {
                if (File.Exists(value))
                {
                    scriptFile = value;
                    ScriptContext = File.ReadAllText(scriptFile);
                }
            }
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
                            Assembly.GetAssembly(typeof(StateChangedEventArgs)),
                            Assembly.GetCallingAssembly(),
                            Assembly.GetEntryAssembly(),
                            Assembly.GetExecutingAssembly(),
                            Assembly.GetAssembly(typeof(System.Globalization.CultureInfo)),
                            Assembly.GetAssembly(typeof(Color)),
                            Assembly.GetAssembly(typeof(Math)),
                            Assembly.GetAssembly(typeof(Regex)),
                            Assembly.GetAssembly(typeof(DynamicObject)),  // System.Dynamic
                            Assembly.GetAssembly(typeof(ExpandoObject)), // System.Dynamic
                            Assembly.GetAssembly(typeof(Microsoft.CSharp.RuntimeBinder.CSharpArgumentInfo)),  // Microsoft.CSharp
                            Assembly.GetAssembly(typeof(Enumerable)), // Linq
                            Assembly.GetAssembly(typeof(DefaultExpression)), // Linq Expression
                            Assembly.GetAssembly(typeof(KnownFolders)), // KnownFolderPaths
                        });
            scriptOptions = scriptOptions.AddImports(new string[] {
                            "System",
                            "System.Collections.Generic",
                            "System.Collections.Specialized",
                            "System.Dynamic",
                            "System.Drawing",
                            "System.Globalization",
                            "System.IO",
                            "System.Linq",
                            "System.Linq.Expressions",
                            "System.Math",
                            "System.Text",
                            "System.Text.RegularExpressions",
                            "System.Windows.Forms",
                            "AutoIt",
                            "KnownFolderPaths",
                            "Newtonsoft.Json",
                            "Elton.Aqara",
                            "MiJia",
                        });

            if (File.Exists(scriptFile))
                Load(File.ReadAllText(scriptFile));

            return (scriptOptions);
        }

        public void Load(string context = "")
        {
            if (!string.IsNullOrEmpty(context)) scriptContext = context;
            script = CSharpScript.Create(scriptContext, scriptOptions, typeof(Globals), loader);
            script.Compile();
        }

        internal void Init(string basepath, string configFile, TextBox logger)
        {
            InitMiJiaGateway(basepath, configFile);
            scriptOptions = InitScriptEngine();
            if (logger is TextBox) Logger = logger;
        }

        internal async Task<ScriptState> RunScript(bool AutoReset = false, bool IsTest = false)
        {
            ScriptState result = null;
            if (sciptRunning || string.IsNullOrEmpty(scriptContext))
                return (result);
            else if (!(client is AqaraClient) || client.Gateways.Count() <= 0)
                return (result);

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
                    globals.Logger.Clear();

                    globals.isTest = IsTest;
                    globals.device = Devices;
                    globals.events = events.Tail(25).ToList();

                    if (!(script is Script)) Load();
                    if (script is Script)
                    {
                        result = await script.RunAsync(globals, cancelToken);
                    }
                    if (AutoReset) globals.Reset();
                    globals.vars.Clear();

                    StringBuilder sb = new StringBuilder();
                    if (globals.Logger.Count > 0)
                    {
                        sb.AppendLine("-- Print Out ".PaddingRight(72, '-'));
                        foreach (var line in globals.Logger)
                        {
                            sb.AppendLine(line);
                        }
                    }
                    if (result is ScriptState && result.Variables.Length > 0)
                    {
                        sb.AppendLine("-- Variables ".PaddingRight(72, '-'));
                        foreach (var v in result.Variables)
                        {
                            if (v.Name.Equals("Device", StringComparison.CurrentCultureIgnoreCase)) continue;
                            if (v.Name.Equals("EventLog", StringComparison.CurrentCultureIgnoreCase)) continue;
                            if (v.Name.StartsWith("_")) continue;

                            sb.AppendLine($"{v.Name} = {v.Value}");
                            globals.vars[v.Name] = v.Value;
                        }
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
        private static NLog.Logger log = NLog.LogManager.GetCurrentClassLogger();

        internal AqaraClient client = default(AqaraClient);
        public string State { get; internal set; } = string.Empty;
        public string StateName { get; internal set; } = string.Empty;
        public uint StateDuration { get; set; } = 0;
        public Dictionary<string, string> Properties { get; set; } = new Dictionary<string, string>();
        public dynamic Info { get; set; }

        public bool Open { get; }

        public void SetState(string key, string value)
        {
            if (client is AqaraClient)
            {
                List<KeyValuePair<string, dynamic>> states = new List<KeyValuePair<string, dynamic>>();
                KeyValuePair<string, dynamic> kv = new KeyValuePair<string, dynamic>(key, value);
                states.Add(kv);
                SetStates(states);
            }
        }

        public void SetStates(IEnumerable<KeyValuePair<string, dynamic>> states)
        {
            if (client is AqaraClient && states is IEnumerable<KeyValuePair<string, dynamic>>)
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
        private static NLog.Logger log = NLog.LogManager.GetCurrentClassLogger();

        public enum MUTE_MODE { Mute, UnMute, Toggle, Background }

        private class ProcInfo
        {
            uint PID { get; set; } = 0;
            uint Parent { get; set; } = 0;
            string Name { get; set; } = string.Empty;
            string Title { get; set; } = string.Empty;
            Process Info { get; set; } = default(Process);
        }

        Dictionary<uint, Process> procs = null;
        private ManagementEventWatcher _watcherStart;
        private ManagementEventWatcher _watcherStop;

        private ScreenPowerMgmt _screenMgmtPower;
        private PowerMgmt _PowerStatus = PowerMgmt.On;
        private void ScreenMgmtPower(object sender, ScreenPowerMgmtEventArgs e)
        {
            _PowerStatus = e.PowerStatus;
            if      (e.PowerStatus == PowerMgmt.StandBy) logger.Add("StandBy Event!");
            else if (e.PowerStatus == PowerMgmt.Off) logger.Add("Off Event!");
            else if (e.PowerStatus == PowerMgmt.On) logger.Add("On Event!");
        }

        private Hardcodet.Wpf.TaskbarNotification.TaskbarIcon tbi = null;

        public Globals()
        {
            if (IsAdmin)
            {
                _watcherStart = new ManagementEventWatcher("SELECT ProcessID, ProcessName FROM Win32_ProcessStartTrace");
                _watcherStart.EventArrived += WatcherProcessStarted;
                _watcherStop = new ManagementEventWatcher("SELECT * FROM Win32_ProcessStopTrace");
                _watcherStop.EventArrived += WatcherProcessStoped;
                _watcherStart.Start();
                _watcherStop.Start();
            }

            _screenMgmtPower = new ScreenPowerMgmt();
            _screenMgmtPower.ScreenPower += ScreenMgmtPower;

            procs = Process.GetProcesses().ToDictionary(p => (uint)p.Id, p => p);

            if (tbi == null)
            {
                var title = Regex.Replace(Title, $@"[{string.Join("|", Path.GetInvalidFileNameChars())}]", "_", RegexOptions.IgnoreCase);
                tbi = new Hardcodet.Wpf.TaskbarNotification.TaskbarIcon()
                {
                    Name = title,
                    Uid = title, 
                    Visibility = System.Windows.Visibility.Collapsed,
                    Icon = Icon.ExtractAssociatedIcon(Application.ExecutablePath),
                    SnapsToDevicePixels = true,
                    UseLayoutRounding = true
                };
            }
        }

        ~Globals()
        {
            if (IsAdmin)
            {
                if (_watcherStop is ManagementEventWatcher) _watcherStop.Stop();
                if (_watcherStart is ManagementEventWatcher) _watcherStart.Stop();
            }

            if (tbi is Hardcodet.Wpf.TaskbarNotification.TaskbarIcon)
            {
                tbi.Dispose();
            }
        }

        internal bool isTest = false;
        public bool IsTest { get { return (isTest); } }

        public bool Is32BitApp { get { return (!Environment.Is64BitProcess); } }
        public bool Is64BitApp { get { return (Environment.Is64BitProcess); } }
        public bool Is32BitOS { get { return (!Environment.Is64BitOperatingSystem); } }
        public bool Is64BitOS { get { return (Environment.Is64BitOperatingSystem); } }

        public bool IsAdmin { get; } = AutoItX.IsAdmin() == 1 ? true : false;

        #region Vars routines
        internal Dictionary<string, dynamic> vars = new Dictionary<string, dynamic>();
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

        public dynamic GetVar(string vn)
        {
            dynamic result = default(dynamic);
            try
            {
                if (vars.ContainsKey(vn)) result = vars[vn];
                else result = default(dynamic);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            return (result);
        }
        #endregion

        #region MiJia Gateway/ZigBee Device
        internal List<KeyValuePair<DateTime, StateChangedEventArgs>> events = new List<KeyValuePair<DateTime, StateChangedEventArgs>>();
        public List<KeyValuePair<DateTime, StateChangedEventArgs>> EventLog { get { return (events); } }

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
            //isTest = false;
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

        public void Minimize(IEnumerable<string> windows)
        {
            if (windows.Count() == 0)
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
                    Regex.IsMatch(up, title, RegexOptions.IgnoreCase) ||
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
                if (childs.Count() > 0)
                {
                    var pname = UWP_AppName(host, pID);
                    if (!result.Contains(pname))
                        result.Add(UWP_AppName(host, pID));
                }
            }

            return (result);
        }

        #endregion
        #endregion

        #region Process routines
        public IEnumerable<Process> AllProcess { get { return Process.GetProcesses(); } }

        private void WatcherProcessStarted(object sender, EventArrivedEventArgs e)
        {
            // add proc to proc dict
            var procinfo = e.NewEvent.Properties;
            var pid = Convert.ToUInt32(procinfo["ProcessID"].Value);
            var proc = GetProcessById(pid);
            procs[pid] = proc;
        }

        private void WatcherProcessStoped(object sender, EventArrivedEventArgs e)
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
            return (GetProcessById((uint)pid));
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

        public List<Process> GetProcessesByName(IEnumerable<string> names, bool regex = false)
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

        public List<Process> GetProcessesByTitle(IEnumerable<string> titles, bool regex = false)
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
            return (ProcessName((uint)pid));
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
            return (KillProcess((uint)pid));
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

        internal void KillProcess(IEnumerable<string> processNames)
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

        public void Kill(IEnumerable<string> processList)
        {
            KillProcess(processList);
        }

        public void RestartProcess(string processName, bool CurrentUserOnly = true)
        {
            if (string.IsNullOrEmpty(processName)) return;

            var processes = GetProcessesByName(processName);
            foreach (var proc in processes)
            {
                var proc_info = proc.StartInfo;
                if (CurrentUserOnly && proc_info.UserName.Equals(Process.GetCurrentProcess().StartInfo.UserName, StringComparison.CurrentCultureIgnoreCase))
                {
                    proc.Kill();
                    var proc_new = Process.Start(proc_info);
                }
                else if (IsAdmin)
                {
                    proc.Kill();
                    var proc_new = Process.Start(proc_info);
                }
            }
            Sleep(100);
            if (GetProcessesByName(processName).Count() <= 0)
            {
                Process.Start(processName);
            }
        }
        #endregion
        #endregion

        #region Service routines
        public Dictionary<string, ServiceController> Service { get { return (GetServices()); } }

        public Dictionary<string, ServiceController> GetServicesByName(string name, bool regex = false)
        {
            var result = new Dictionary<string, ServiceController>();

            IEnumerable<ServiceController> services = null;
            if (regex)
            {
                services = ServiceController.GetServices().Where(p => Regex.IsMatch(p.ServiceName, $"{name}", RegexOptions.IgnoreCase));
            }
            else
            {
                services = ServiceController.GetServices().Where(p => p.ServiceName.Equals(name, StringComparison.OrdinalIgnoreCase));
            }
            foreach (var service in services)
            {
                result.Add(service.ServiceName, service);
            }

            return (result);
        }

        public Dictionary<string, ServiceController> GetServices()
        {
            var result = new Dictionary<string, ServiceController>();

            var services = ServiceController.GetServices();
            foreach (var service in services)
            {
                result.Add(service.ServiceName, service);
            }

            return (result);
        }

        public Dictionary<string, ServiceController> GetServices(string name, bool regex = true)
        {
            return (GetServicesByName(name, regex));
        }

        public Dictionary<string, ServiceController> Services(string name, bool regex = true)
        {
            return (GetServicesByName(name, regex));
        }
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
        private static extern int SendMessage(int hWnd, uint Msg, int wParam, int lParam);

        [DllImport("user32.dll", CharSet = CharSet.Auto)]
        private static extern bool LockWorkStation();

        public bool LockScreen()
        {
            return (LockWorkStation());
        }

        internal void Monitor(bool on, bool lockscreen = false)
        {
            var handle = AutoItX.WinGetHandle("[CLASS:Progman]");
            if (handle != IntPtr.Zero)
            {
                if (on)
                    SendMessage(handle.ToInt32(), LCI_WM_SYSCommand, LCI_SC_MonitorPower, LCI_Power_On);
                else
                    SendMessage(handle.ToInt32(), LCI_WM_SYSCommand, LCI_SC_MonitorPower, LCI_Power_Off);
            }
            if (lockscreen) LockScreen();
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

        public bool IsMonitorOn { get { return (_PowerStatus == PowerMgmt.On || _PowerStatus == PowerMgmt.StandBy); } }
        public bool IsMonitorOff { get { return (_PowerStatus == PowerMgmt.Off); } }

        #endregion

        #region Media routines
        private bool InternalPlay = false;
        private NAudio.Wave.WaveOut waveOut = new NAudio.Wave.WaveOut();

        // Application Mute/UnMute/ToggleMute
        public void AppMute(MUTE_MODE mode, IEnumerable<string> apps)
        {
            if(apps is IEnumerable<string>)
            {
                foreach(var app in apps)
                {
                    AppMute(mode, app);
                }
            }
        }

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
                                        (!string.IsNullOrEmpty(title) && 
                                        (Regex.IsMatch(app, title, RegexOptions.IgnoreCase) ||
                                        Regex.IsMatch(title, app, RegexOptions.IgnoreCase)))))
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

        public void AppMute(IEnumerable<string> apps)
        {
            if (apps is IEnumerable<string>)
            {
                foreach (var app in apps)
                {
                    AppMute(app);
                }
            }
        }

        public void AppMute(string app = default(string))
        {
            AppMute(MUTE_MODE.Mute, app);
        }

        public void AppUnMute(IEnumerable<string> apps)
        {
            if (apps is IEnumerable<string>)
            {
                foreach (var app in apps)
                {
                    AppUnMute(app);
                }
            }
        }

        public void AppUnMute(string app = default(string))
        {
            AppMute(MUTE_MODE.UnMute, app);
        }

        public void AppToggleMute(IEnumerable<string> apps)
        {
            if (apps is IEnumerable<string>)
            {
                foreach (var app in apps)
                {
                    AppToggleMute(app);
                }
            }
        }

        public void AppToggleMute(string app = default(string))
        {
            AppMute(MUTE_MODE.Toggle, app);
        }

        public void AppBackgroundMute(IEnumerable<string> apps)
        {
            if (apps is IEnumerable<string>)
            {
                foreach (var app in apps)
                {
                    AppBackgroundMute(app);
                }
            }
        }

        public void AppBackgroundMute(string app = default(string))
        {
            AppMute(MUTE_MODE.Background, app);
        }

        // Device Mute/UnMute/ToggleMute
        public void Mute(MUTE_MODE mode, IEnumerable<string> devices)
        {
            if (devices is IEnumerable<string>)
            {
                foreach (var device in devices)
                {
                    Mute(mode, device);
                }
            }
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

        public void Mute(IEnumerable<string> devices)
        {
            if (devices is IEnumerable<string>)
            {
                foreach (var device in devices)
                {
                    Mute(device);
                }
            }
        }

        public void Mute(string device = default(string))
        {
            Mute(MUTE_MODE.Mute, device);
        }

        public void UnMute(IEnumerable<string> devices)
        {
            if (devices is IEnumerable<string>)
            {
                foreach (var device in devices)
                {
                    UnMute(device);
                }
            }
        }

        public void UnMute(string device = default(string))
        {
            Mute(MUTE_MODE.UnMute, device);
        }

        public void ToggleMute(IEnumerable<string> devices)
        {
            if (device is IEnumerable<string>)
            {
                foreach (var device in devices)
                {
                    ToggleMute(device);
                }
            }
        }

        public void ToggleMute(string device = default(string))
        {
            Mute(MUTE_MODE.Toggle, device);
            //AutoItX.Send("{VOLUME_MUTE}");
            //AutoItX.Sleep(10);
        }

        public void BackgroundMute(IEnumerable<string> devices)
        {
            if (device is IEnumerable<string>)
            {
                foreach (var device in devices)
                {
                    BackgroundMute(device);
                }
            }
        }

        public void BackgroundMute(string app = default(string))
        {
            Mute(MUTE_MODE.Background, app);
        }

        public bool Muted(IEnumerable<string> devices)
        {
            bool result = false;
            if (device is IEnumerable<string>)
            {
                foreach (var device in devices)
                {
                    if (Muted(device))
                    {
                        result = true;
                        break;
                    }
                }
            }
            return (result);
        }

        public bool Muted(string device = default(string))
        {
            bool result = false;

            try
            {
                if (string.IsNullOrEmpty(device))
                {
                    MMDevice maindev = new MMDeviceEnumerator().GetDefaultAudioEndpoint(DataFlow.All, Role.Multimedia);
                    if (maindev.AudioEndpointVolume.Mute) result = true;
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

            return (result);
        }

        public bool AppMuted(IEnumerable<string> apps)
        {
            bool result = false;
            if (apps is IEnumerable<string>)
            {
                foreach (var app in apps)
                {
                    if (AppMuted(app))
                    {
                        result = true;
                        break;
                    }
                }
            }
            return (result);
        }

        public bool AppMuted(string app = default(string))
        {
            bool result = false;

            try
            {
                if (string.IsNullOrEmpty(app))
                {
                    MMDevice maindev = new MMDeviceEnumerator().GetDefaultAudioEndpoint(DataFlow.Render, Role.Multimedia);
                    if (maindev.AudioEndpointVolume.Mute) result = true;
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
                                        (!string.IsNullOrEmpty(title) &&
                                        (Regex.IsMatch(app, title, RegexOptions.IgnoreCase) ||
                                        Regex.IsMatch(title, app, RegexOptions.IgnoreCase)))))
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

        private bool DeviceIsActive(MMDevice dev, IEnumerable<string> apps)
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
                        if (!(apps is string[]) || apps.Count() <= 0) result = true;
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

        private bool DeviceIsActive(MMDevice dev, IEnumerable<int> pids)
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
                        if (!(pids is int[]) || pids.Count() <= 0) result = true;
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

        private bool MediaIsActive(DataFlow mode, IEnumerable<string> apps)
        {
            bool result = false;

            try
            {
                if (!(apps is string[]) || apps.Count() <= 0)
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

        private bool MediaIsActive(DataFlow mode, IEnumerable<int> pids)
        {
            bool result = false;

            try
            {
                if (!(pids is int[]) || pids.Count() <= 0)
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

        public bool MediaIsOut(IEnumerable<string> apps)
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

        public bool MediaIsOut(IEnumerable<int> pids)
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

        public bool MediaIsIn(IEnumerable<string> apps)
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

        public bool MediaIsIn(IEnumerable<int> pids)
        {
            bool result = false;

            result = MediaIsActive(DataFlow.Capture, pids);

            return (result);
        }

        public void MediaPlay(string media = default(string))
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
                if (waveOut is NAudio.Wave.WaveOut)
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
        #region Get Input Idle Info
        [StructLayout(LayoutKind.Sequential)]
        struct LASTINPUTINFO
        {
            // 设置结构体块容量  
            [MarshalAs(UnmanagedType.U4)]
            public int cbSize;
            // 捕获的时间  
            [MarshalAs(UnmanagedType.U4)]
            public uint dwTime;
        }

        [DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]
        private static extern bool GetLastInputInfo(ref LASTINPUTINFO plii);
        #endregion

        //获取键盘和鼠标没有操作的时间  
        public static double GetLastInputTime()
        {
            uint idleTime = 0;
            LASTINPUTINFO lastInputInfo = new LASTINPUTINFO();
            lastInputInfo.cbSize = Marshal.SizeOf(lastInputInfo);
            lastInputInfo.dwTime = 0;

            uint envTicks = (uint)Environment.TickCount;

            if (GetLastInputInfo(ref lastInputInfo))
            {
                uint lastInputTick = lastInputInfo.dwTime;

                idleTime = envTicks - lastInputTick;
            }

            return ((idleTime > 0) ? (idleTime / 1000.0) : 0);

            //LASTINPUTINFO vLastInputInfo = new LASTINPUTINFO();
            //vLastInputInfo.cbSize = Marshal.SizeOf(vLastInputInfo);
            //// 捕获时间  
            //if (!GetLastInputInfo(ref vLastInputInfo))
            //    return 0;
            //else
            //    return Environment.TickCount - (long)vLastInputInfo.dwTime;
        }

        public double Idle
        {
            get { return (GetLastInputTime()); }
        }

        public bool IsIdle(double secs = 10)
        {
            return(secs >= 0 && Idle >= secs ? true : false);
        }

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
        public void Speak(string text, int vol = 100, int rate = 0)
        {
            List<string> lang_cn = new List<string>() { "zh-hans", "zh-cn", "zh" };
            List<string> lang_tw = new List<string>() { "zh-hant", "zh-tw" };
            List<string> lang_jp = new List<string>() { "ja-jp", "ja", "jp" };
            List<string> lang_en = new List<string>() { "en-us", "us", "en" };

            try
            {
                if (string.IsNullOrEmpty(voice_default)) voice_default = synth.Voice.Name;

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

        public void Say(string text, int vol = 100, int rate = 0)
        {
            Speak(text, vol, rate);
        }

        private List<string> logger = new List<string>();
        public List<string> Logger { get { return (logger); } }
        public void Print(string text)
        {
            logger.Add(text);
        }

        public void Print(dynamic content)
        {
            logger.Add($"{content}");
        }

        public void Log(dynamic content)
        {
            logger.Add($"{content}");
            log.Info($"{content}");
        }

        public string Title
        {
            get
            {
                var entryAssembly = Assembly.GetEntryAssembly();
                var applicationTitle = ((AssemblyTitleAttribute)entryAssembly.GetCustomAttributes(typeof(AssemblyTitleAttribute), false)[0]).Title;
                if (string.IsNullOrWhiteSpace(applicationTitle))
                {
                    applicationTitle = entryAssembly.GetName().Name;
                }
                return (applicationTitle);
            }
        }

        //private var appbar = new Hardcodet.Wpf.TaskbarNotification.Interop.AppBarInfo();
        public void Toast(string content, string title = "", MessageBoxIcon icon = MessageBoxIcon.None)
        {
            if (tbi is Hardcodet.Wpf.TaskbarNotification.TaskbarIcon)
            {
                var ballon_icon = Hardcodet.Wpf.TaskbarNotification.BalloonIcon.None;
                switch(icon)
                {
                    case MessageBoxIcon.Information:
                        ballon_icon = Hardcodet.Wpf.TaskbarNotification.BalloonIcon.Info;
                        break;
                    case MessageBoxIcon.Warning:
                        ballon_icon = Hardcodet.Wpf.TaskbarNotification.BalloonIcon.Warning;
                        break;
                    case MessageBoxIcon.Error:
                        ballon_icon = Hardcodet.Wpf.TaskbarNotification.BalloonIcon.Error;
                        break;
                    default:
                        ballon_icon = Hardcodet.Wpf.TaskbarNotification.BalloonIcon.None;
                        break;
                }
                if (string.IsNullOrEmpty(title)) title = Title;
                tbi.Visibility = System.Windows.Visibility.Visible;
                tbi.ShowBalloonTip(title, content, ballon_icon);
                tbi.Visibility = System.Windows.Visibility.Collapsed;
            }
        }

        public void Notify(string content, string title = "", MessageBoxIcon icon = MessageBoxIcon.None)
        {
            Toast(content, title, icon);
        }
        #endregion
    }

    #region Power Management Helper
    //
    // https://stackoverflow.com/questions/2208595/c-sharp-how-to-get-the-events-when-the-screen-display-goes-to-power-off-or-on
    //

    public enum PowerMgmt
    {
        StandBy,
        Off,
        On
    };

    public class ScreenPowerMgmtEventArgs
    {
        private PowerMgmt _PowerStatus;

        public ScreenPowerMgmtEventArgs(PowerMgmt powerStat)
        {
            _PowerStatus = powerStat;
        }

        public PowerMgmt PowerStatus
        {
            get { return this._PowerStatus; }
        }
    }

    public class ScreenPowerMgmt
    {
        public delegate void ScreenPowerMgmtEventHandler(object sender, ScreenPowerMgmtEventArgs e);
        public event ScreenPowerMgmtEventHandler ScreenPower;

        private void OnScreenPowerMgmtEvent(ScreenPowerMgmtEventArgs args)
        {
            ScreenPower?.Invoke(this, args);
        }

        public void SwitchMonitorOff()
        {
            /* The code to switch off */
            OnScreenPowerMgmtEvent(new ScreenPowerMgmtEventArgs(PowerMgmt.Off));
        }

        public void SwitchMonitorOn()
        {
            /* The code to switch on */
            OnScreenPowerMgmtEvent(new ScreenPowerMgmtEventArgs(PowerMgmt.On));
        }

        public void SwitchMonitorStandby()
        {
            /* The code to switch standby */
            OnScreenPowerMgmtEvent(new ScreenPowerMgmtEventArgs(PowerMgmt.StandBy));
        }
    }
    #endregion

    public static class ReflectionExtensions
    {
        public static bool IsPublic(this object obj)
        {
            // Set the flags so that private and public fields from instances will be found
            var bindingFlags = BindingFlags.NonPublic | BindingFlags.Public | BindingFlags.Static | BindingFlags.Instance;
            var field = obj.GetType().GetField("IsPublic", bindingFlags);

            if (field == null) return (false);
            else return (true);
        }

        public static bool IsPrivate(this object obj)
        {
            // Set the flags so that private and public fields from instances will be found
            var bindingFlags = BindingFlags.NonPublic | BindingFlags.Public | BindingFlags.Static | BindingFlags.Instance;
            var field = obj.GetType().GetField("IsPrivate", bindingFlags);
            if (field == null) return (false);
            else return (true);
        }

        public static bool IsStatic(this object obj)
        {
            // Set the flags so that private and public fields from instances will be found
            var bindingFlags = BindingFlags.NonPublic | BindingFlags.Public | BindingFlags.Static | BindingFlags.Instance;
            var field = obj.GetType().GetField("IsStatic", bindingFlags);
            if (field == null) return (false);
            else return (true);
        }

        public static bool IsSeal(this object obj)
        {
            // Set the flags so that private and public fields from instances will be found
            var bindingFlags = BindingFlags.NonPublic | BindingFlags.Public | BindingFlags.Static | BindingFlags.Instance;
            var field = obj.GetType().GetField("IsSeal", bindingFlags);
            if (field == null) return (false);
            else return (true);
        }

        public static T GetFieldValue<T>(this object obj, string name)
        {
            // Set the flags so that private and public fields from instances will be found
            var bindingFlags = BindingFlags.Public | BindingFlags.NonPublic | BindingFlags.Instance;
            var field = obj.GetType().GetField(name, bindingFlags);
            return (T)field?.GetValue(obj);
        }

        public static T GetPublicFieldValue<T>(this object obj, string name)
        {
            // Set the flags so that private and public fields from instances will be found
            var bindingFlags = BindingFlags.Public | BindingFlags.Instance;
            var field = obj.GetType().GetField(name, bindingFlags);
            return (T)field?.GetValue(obj);
        }

        public static T GetPrivateFieldValue<T>(this object obj, string name)
        {
            // Set the flags so that private and public fields from instances will be found
            var bindingFlags = BindingFlags.NonPublic | BindingFlags.Instance;
            var field = obj.GetType().GetField(name, bindingFlags);
            return (T)field?.GetValue(obj);
        }
    }
}
