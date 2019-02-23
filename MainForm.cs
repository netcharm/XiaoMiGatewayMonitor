using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Elton.Aqara;
using KnownFolderPaths;
using System.IO;
using AutoIt;
using System.Runtime.InteropServices;

namespace MiJia
{
    public partial class MainForm : Form
    {
        string APPFOLDER = Path.GetDirectoryName(Application.ExecutablePath);
        string SKYFOLDER = Path.Combine(KnownFolders.GetPath(KnownFolder.SkyDrive), @"ApplicationData\ConnectedHome");
        string DOCFOLDER = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.CommonDocuments), @"Elton\ConnectedHome\");
        string USERNAME = Environment.UserName;

        AqaraClient client = null;
        Timer timerRefresh = null;
        bool DOOR_CLOSE = false;
        bool MOTION = false;

        private Gateway gw = new Gateway();

        #region Kill Process
        internal bool KillProcess(string processName)
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

        internal bool KillProcess(int pid)
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

        internal void KillProcess(string[] processNames)
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

        internal void Monitor(bool on)
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
        }

        private void DeviceStateChanged(object sender, StateChangedEventArgs e)
        {
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

            if (MOTION && USERNAME.StartsWith("netch", StringComparison.CurrentCultureIgnoreCase))
            {
                KillProcess(new string[] { "mpc-be64.exe", "zPlayer UWP.exe", "mangameeya.exe", "comicsviewer.exe", });
                //Task.Run(() => { KillProcess(new string[] { "mpc-be64.exe", "zPlayer UWP.exe", "mangameeya.exe", "comicsviewer.exe", }); });
                MOTION = false;
            }
            if (DOOR_CLOSE)
            {
                Monitor(false);
                DOOR_CLOSE = false;
            }
        }

        private async void btnTest_Click(object sender, EventArgs e)
        {
            var devices = await gw.GetDevice();
            edResult.Text += string.Join(Environment.NewLine, devices);

            foreach (var deviceId in devices)
            {
                var deviceInfo = await gw.GetDevice(deviceId);
                edResult.Text += string.Join(Environment.NewLine, deviceInfo);
            }
        }

    }
}
