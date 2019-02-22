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

namespace MiJia
{
    public partial class MainForm : Form
    {
        string APPFOLDER = Path.GetDirectoryName(Application.ExecutablePath);
        string SKYFOLDER = Path.Combine(KnownFolders.GetPath(KnownFolder.SkyDrive), @"ApplicationData\ConnectedHome");
        string DOCFOLDER = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.CommonDocuments), @"Elton\ConnectedHome\");

        AqaraClient client = null;
        Timer timerRefresh = null;

        private Gateway gw = new Gateway();

        internal bool KillProcess(string processName)
        {
            bool result = true;

            //Process[] pslist = Process.GetProcesses();
            Process[] pslist = Process.GetProcessesByName(Path.GetFileNameWithoutExtension(processName));
            foreach (var ps in pslist)
            {
                if (ps.ProcessName.EndsWith(processName, StringComparison.CurrentCultureIgnoreCase))
                {
                    ps.Kill();
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

            timerRefresh = new Timer();
            timerRefresh.Interval = 1000;
            timerRefresh.Tick += TimerRefresh_Tick;
            timerRefresh.Start();
        }

        private void TimerRefresh_Tick(object sender, EventArgs e)
        {
            var gateway = client.Gateways.Values.FirstOrDefault();
            if (gateway == null)
            {

            }
            else
            {
                var sb = new StringBuilder();
                sb.AppendLine(gateway.EndPoint?.ToString());
                sb.AppendLine(gateway.LatestTimestamp.ToString());
                sb.AppendLine(gateway.Token);

                bool c0 = false;
                bool c1 = false;

                foreach (var device in gateway.Devices.Values)
                {
                    sb.AppendLine($"{device.Name}");
                    foreach (var pair in device.States)
                    {
                        sb.Append($"{pair.Key} = {pair.Value.Value} ");
                        if(device.Name.Equals("书房-门", StringComparison.CurrentCulture) && 
                            pair.Key.Equals("status", StringComparison.CurrentCultureIgnoreCase) &&
                            pair.Value.Value.Equals("open", StringComparison.CurrentCultureIgnoreCase))
                        {
                            c0 = true;
                        }
                        else if (device.Name.Equals("走道-人体传感器", StringComparison.CurrentCulture) &&
                            pair.Key.Equals("status", StringComparison.CurrentCultureIgnoreCase) &&
                            pair.Value.Value.Equals("motion", StringComparison.CurrentCultureIgnoreCase))
                        {
                            c1 = true;
                        }
                    }
                    sb.AppendLine();
                }
                edResult.Text = sb.ToString();
                if (c0 && c1) KillProcess("mpc-be64");
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
            //KillProcess("mpc-be64.exe");
        }

    }
}
