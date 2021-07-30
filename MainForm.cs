using AutoIt;
using KnownFolderPaths;
using System;
using System.Drawing;
using System.IO;
using System.Windows.Forms;

namespace MiJia
{
    public partial class MainForm : Form
    {
        internal static string APPFOLDER = Path.GetDirectoryName(Application.ExecutablePath);
        internal string SKYFOLDER = Path.Combine(KnownFolders.GetPath(KnownFolder.SkyDrive), @"ApplicationData\ConnectedHome");
        internal string DOCFOLDER = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.CommonDocuments), @"Elton\ConnectedHome\");
        internal string SCRIPT_FILE = Path.Combine(APPFOLDER, "actions.csx");
        internal string USERNAME = Environment.UserName;

        private Gateway gw = new Gateway();

        internal ScriptEngine engine = new ScriptEngine();

        private void SetHidden(bool hide = true)
        {
            if (hide)
            {
                ShowInTaskbar = false;
                WindowState = FormWindowState.Minimized;
                Hide();
            }
            else
            {
                ShowInTaskbar = true;
                WindowState = FormWindowState.Normal;
                Show();
            }
        }

        private Action<string, string, MessageBoxIcon> NotifyAction { get; set; } = null;
        private void InitNotifyIcon()
        {
            notifyIcon.Icon = Icon;
            notifyIcon.Text = Text;
            notifyIcon.BalloonTipTitle = Text;
            notifyIcon.BalloonTipIcon = ToolTipIcon.Info;

            NotifyAction = new Action<string, string, MessageBoxIcon>((content, title, icon) =>
            {
                if (notifyIcon is NotifyIcon)
                {
                    if (string.IsNullOrEmpty(title)) title = Text;
                    ToolTipIcon ballon_icon = ToolTipIcon.None;
                    switch (icon)
                    {
                        case MessageBoxIcon.Information:
                            ballon_icon = ToolTipIcon.Info;
                            break;
                        case MessageBoxIcon.Warning:
                            ballon_icon = ToolTipIcon.Warning;
                            break;
                        case MessageBoxIcon.Error:
                            ballon_icon = ToolTipIcon.Error;
                            break;
                        default:
                            ballon_icon = ToolTipIcon.None;
                            break;
                    }
                    notifyIcon.ShowBalloonTip(5, title, content, ballon_icon);
                }
            });
        }

        public MainForm()
        {
            InitializeComponent();
        }

        private void MainForm_Load(object sender, EventArgs e)
        {
            Icon = Icon.ExtractAssociatedIcon(Application.ExecutablePath);
            InitNotifyIcon();

            var basepath = APPFOLDER;
            if (Directory.Exists(APPFOLDER))
                basepath = APPFOLDER;
            else if (Directory.Exists(SKYFOLDER))
                basepath = SKYFOLDER;
            else if (Directory.Exists(DOCFOLDER))
                basepath = DOCFOLDER;

            var opts = Environment.CommandLine.Split();
            if (opts.Length > 1)
            {
                if (!string.IsNullOrEmpty(opts[1]) && File.Exists(opts[1])) SCRIPT_FILE = opts[1];
            }

            var configFile = Path.Combine(basepath, "config", "aqara.json");
            if (engine is ScriptEngine)
            {
                engine.NotificationAction = NotifyAction;
                engine.Init(basepath, configFile, edResult);
                engine.ScriptFile = SCRIPT_FILE;
            }
#if DEBUG
            if (USERNAME.StartsWith("netch", StringComparison.CurrentCultureIgnoreCase)) btnTest.Visible = true;
            else btnTest.Visible = false;
            ContextMenuStrip = contextNotify;
#else
            chkOnTop.Checked = true;
            //SetHide(true);
            this.WindowState = FormWindowState.Minimized;
#endif
        }

        private void MainForm_FormClosing(object sender, FormClosingEventArgs e)
        {
#if !DEBUG
            var ret = MessageBox.Show(this, "Exit?", "Question", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button1);
            if (ret != DialogResult.Yes)
            {
                e.Cancel = true;
            }
#endif
        }

        private void MainForm_Resize(object sender, EventArgs e)
        {
            if (sender == this)
            {
                switch(WindowState)
                {
                    case FormWindowState.Normal:
                        SetHidden(false);
                        break;
                    case FormWindowState.Maximized:
                        SetHidden(false);
                        break;
                    case FormWindowState.Minimized:
                        SetHidden(true);
                        break;
                    default:
                        break;
                }
            }
        }

        private void notifyIcon_DoubleClick(object sender, EventArgs e)
        {
            if (e is MouseEventArgs)
            {
                var me = e as MouseEventArgs;
                if (me.Button == MouseButtons.Left && sender == notifyIcon)
                {
                    SetHidden(false);
                }
            }
        }

        private async void btnTest_Click(object sender, EventArgs e)
        {
            if (engine is ScriptEngine)
            {
                var ret = await engine.RunScript(true, true);
            }
        }

        private void btnReloadScript_Click(object sender, EventArgs e)
        {
            //InitScriptEngine();
            if (File.Exists(SCRIPT_FILE) && engine is ScriptEngine)
                engine.ScriptContext = File.ReadAllText(SCRIPT_FILE);
        }

        private void btnEditScript_Click(object sender, EventArgs e)
        {
            if (File.Exists(SCRIPT_FILE))
            {
#if DEBUG
                var ret = AutoItX.Run($"notepad2 /s cs {SCRIPT_FILE}", APPFOLDER);
                //var ret = AutoItX.RunWait($"notepad2 /s cs {sf}", APPFOLDER);
                //if (ret == 0 && engine is ScriptEngine) engine.ScriptContext = File.ReadAllText(sf);
                //else MessageBox.Show("notepad2 run failed!");
#else
                AutoItX.Run($"notepad2 /s cs {SCRIPT_FILE}", APPFOLDER);
#endif
            }
        }

        private void chkOnTop_CheckedChanged(object sender, EventArgs e)
        {
            if (sender == chkOnTop)
            {
                TopMost = chkOnTop.Checked;
                tsmiOnTop.Checked = chkOnTop.Checked;
            }
            else if (sender == tsmiOnTop)
            {
                chkOnTop.Checked = tsmiOnTop.Checked;
            }
        }

        private void chkPause_CheckStateChanged(object sender, EventArgs e)
        {
            if (sender == chkPause)
            {
                if (engine is ScriptEngine) engine.Pausing = chkPause.Checked;
                tsmiPause.Checked = chkPause.Checked;
            }
            else if (sender == tsmiPause)
            {
                chkPause.Checked = tsmiPause.Checked;
            }
        }

        private void tsmiShowForm_Click(object sender, EventArgs e)
        {
            if (!Visible) SetHidden(false);
        }

        private void tsmiResetGateway_Click(object sender, EventArgs e)
        {
            var basepath = APPFOLDER;
            if (Directory.Exists(APPFOLDER))
                basepath = APPFOLDER;
            else if (Directory.Exists(SKYFOLDER))
                basepath = SKYFOLDER;
            else if (Directory.Exists(DOCFOLDER))
                basepath = DOCFOLDER;

            var configFile = Path.Combine(basepath, "config", "aqara.json");
            engine.InitMiJiaGateway(basepath, configFile);
        }

        private void tsmiExit_Click(object sender, EventArgs e)
        {
            //Environment.Exit(0);
            //Application.Exit();
            Close();
        }
    }

}
