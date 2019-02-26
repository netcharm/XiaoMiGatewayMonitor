using AutoIt;
using Elton.Aqara;
using KnownFolderPaths;
using Microsoft.CodeAnalysis.Scripting;
using Microsoft.CodeAnalysis.Scripting.CSharp;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace MiJia
{
    public partial class MainForm : Form
    {
        internal string APPFOLDER = Path.GetDirectoryName(Application.ExecutablePath);
        internal string SKYFOLDER = Path.Combine(KnownFolders.GetPath(KnownFolder.SkyDrive), @"ApplicationData\ConnectedHome");
        internal string DOCFOLDER = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.CommonDocuments), @"Elton\ConnectedHome\");
        internal string USERNAME = Environment.UserName;

        //static public Dictionary<string, KeyValuePair<string, string>> DEVICE_STATES = new Dictionary<string, KeyValuePair<string, string>>();
        public Dictionary<string, string> DEVICE_STATES = new Dictionary<string, string>();
        public Dictionary<string, DEVICE> Devices = new Dictionary<string, DEVICE>();

        internal AqaraClient client = null;
        internal Timer timerRefresh = null;
        internal bool DOOR_CLOSE = false;
        internal bool MOTION = false;

        private Gateway gw = new Gateway();

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
        const int LCI_WM_SYSCommand = 274;
        const int LCI_SC_MonitorPower = 61808;
        const int LCI_Power_Off = 2;
        const int LCI_Power_On = -1;

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

        internal string scriptContext = string.Empty;
        internal ScriptOptions scriptOptions = ScriptOptions.Default;
        //internal Script script;

        internal ScriptOptions InitScriptEngine()
        {
            scriptOptions = ScriptOptions.Default;
            scriptOptions = scriptOptions.AddSearchPaths(APPFOLDER);
            //options = options.AddReferences(AppDomain.CurrentDomain.GetAssemblies());
            scriptOptions = scriptOptions.AddReferences(new Assembly[] {
                    Assembly.GetAssembly(typeof(Path)),
                    Assembly.GetAssembly(typeof(AutoItX)),
                    Assembly.GetAssembly(typeof(JsonConvert)),
                    Assembly.GetAssembly(typeof(AqaraDevice)),
                    Assembly.GetAssembly(GetType()),
                    Assembly.GetAssembly(typeof(System.Dynamic.DynamicObject)),  // System.Code
                    Assembly.GetAssembly(typeof(Microsoft.CSharp.RuntimeBinder.CSharpArgumentInfo)),  // Microsoft.CSharp
                    Assembly.GetAssembly(typeof(System.Dynamic.ExpandoObject))  // System.Dynamic                    
                });
            scriptOptions = scriptOptions.AddNamespaces(new string[] {
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
                scriptContext = File.ReadAllText(sf);

            return (scriptOptions);
        }

        internal ScriptState RunScript()
        {
            ScriptState result = null;

            var gateway = client.Gateways.Values.FirstOrDefault();
            if (gateway == null)
            {
                edResult.Text = string.Empty;
            }
            else
            {
                foreach (var device in gateway.Devices.Values)
                {
                    if (!DEVICE_STATES.ContainsKey(device.Name)) DEVICE_STATES[device.Name] = string.Empty;
                    if (!Devices.ContainsKey(device.Name))
                        Devices[device.Name] = new DEVICE() { State = string.Empty, Info = device };
                }

                try
                {
                    var globals = new Globals()
                    {
                        Device = Devices,
                        DEVICE_STATE = DEVICE_STATES,
                        DEVICE_LIST = gateway.Devices,
                    };
                    result = CSharpScript.Run(scriptContext, scriptOptions, globals);
                }
                catch (Exception ex)
                {
                    edResult.Text += ex.Message;
                }
            }

            return (result);
        }

        public MainForm()
        {
            InitializeComponent();
        }

        private void MainForm_Load(object sender, EventArgs e)
        {
            Icon = Icon.ExtractAssociatedIcon(Application.ExecutablePath);

            var basepath = APPFOLDER;
            if (Directory.Exists(APPFOLDER))
                basepath = APPFOLDER;
            else if (Directory.Exists(SKYFOLDER))
                basepath = SKYFOLDER;
            else if(Directory.Exists(DOCFOLDER))
                basepath = DOCFOLDER;

            var configFile = Path.Combine(basepath, "config", "aqara.json");
            AqaraConfig config = null;
            if(File.Exists(Path.Combine(basepath, "aqara.json")))
            {
                config = AqaraConfig.Parse(File.ReadAllText(configFile));
            }
            else if (File.Exists(configFile))
            {
                config = AqaraConfig.Parse(File.ReadAllText(configFile));
            }
            client = new AqaraClient(config);
            Task.Run(() => { client.DoWork(null); });
            client.DeviceStateChanged += DeviceStateChanged;

            timerRefresh = new Timer();
            timerRefresh.Interval = 1000;
            timerRefresh.Tick += TimerRefresh_Tick;
            timerRefresh.Start();

            if (USERNAME.StartsWith("netch", StringComparison.CurrentCultureIgnoreCase)) btnTest.Visible = true;
            else btnTest.Visible = false;

            InitScriptEngine();
        }

        private void DeviceStateChanged(object sender, StateChangedEventArgs e)
        {
            DEVICE_STATES[e.Device.Name] = e.NewData;
            if (Devices.ContainsKey(e.Device.Name) && Devices[e.Device.Name] is DEVICE)
            {
                Devices[e.Device.Name].State = e.NewData;
                Devices[e.Device.Name].Info = e.Device;
            }
            else
                Devices[e.Device.Name] = new DEVICE() { State = e.NewData, Info = e.Device };

            if (e.Device.Name.Equals("走道-人体传感器", StringComparison.CurrentCulture))
            {
                if (e.StateName.Equals("status", StringComparison.CurrentCultureIgnoreCase) && 
                    e.NewData.Equals("motion", StringComparison.CurrentCultureIgnoreCase))
                {
                    MOTION = true;
                }
            }
            else if(e.Device.Name.Equals("书房-门", StringComparison.CurrentCulture))
            {
                if (e.StateName.Equals("status", StringComparison.CurrentCultureIgnoreCase) &&
                    e.NewData.Equals("close", StringComparison.CurrentCultureIgnoreCase))
                {
                    DOOR_CLOSE = true;
                }
            }
        }

        private void TimerRefresh_Tick(object sender, EventArgs e)
        {
            var gateway = client.Gateways.Values.FirstOrDefault();
            if (gateway == null)
            {
                edResult.Text = string.Empty;
            }
            else
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
                edResult.Text = sb.ToString();


            }
            RunScript();

            //if (MOTION && USERNAME.StartsWith("netch", StringComparison.CurrentCultureIgnoreCase))
            //{
            //    KillProcess(new string[] { "mpc-be64.exe", "zPlayer UWP.exe", "mangameeya.exe", "comicsviewer.exe", });
            //    //Task.Run(() => { KillProcess(new string[] { "mpc-be64.exe", "zPlayer UWP.exe", "mangameeya.exe", "comicsviewer.exe", }); });
            //    MOTION = false;
            //}
            //if (DOOR_CLOSE)
            //{
            //    Monitor(false);
            //    DOOR_CLOSE = false;
            //}
        }

        private void btnTest_Click(object sender, EventArgs e)
        {
            RunScript();
        }

        private void btnReloadScript_Click(object sender, EventArgs e)
        {
            //InitScriptEngine();
            var sf = Path.Combine(APPFOLDER, "actions.csx");
            if (File.Exists(sf))
                scriptContext = File.ReadAllText(sf);
        }

        private void btnEditScript_Click(object sender, EventArgs e)
        {
            var sf = Path.Combine(APPFOLDER, "actions.csx");
            if (File.Exists(sf))
            {
                var ret = AutoItX.RunWait($"notepad2 /s cs {sf}", APPFOLDER);
                if (ret == 0) scriptContext = File.ReadAllText(sf);
                else MessageBox.Show("notepad2 run failed!");
            }
        }
    }

    public class DEVICE
    {
        public string State { get; set; } = string.Empty;
        public AqaraDevice Info { get; set; }

        public void Reset()
        {
            State = string.Empty;
        }
    }

    public class Globals
    {
        public Dictionary<string, DEVICE> Device = new Dictionary<string, DEVICE>();
        public Dictionary<string, string> DEVICE_STATE = new Dictionary<string, string>();
        public Dictionary<string, AqaraDevice> DEVICE_LIST = new Dictionary<string, AqaraDevice>();

        public void Reset()
        {
            foreach(var device in Device)
            {
                device.Value.Reset();
            }
        }

        public void Kill(string[] processList)
        {
            MainForm.KillProcess(processList);
        }

        public void Monitor(bool on)
        {
            MainForm.Monitor(on);
        }

        public void Mute()
        {
            AutoItX.Send("{VOLUME_MUTE}");
        }
    }

}
