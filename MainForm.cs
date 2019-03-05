﻿using AutoIt;
using KnownFolderPaths;
using System;
using System.Drawing;
using System.IO;
using System.Windows.Forms;

namespace MiJia
{
    public partial class MainForm : Form
    {
        internal string APPFOLDER = Path.GetDirectoryName(Application.ExecutablePath);
        internal string SKYFOLDER = Path.Combine(KnownFolders.GetPath(KnownFolder.SkyDrive), @"ApplicationData\ConnectedHome");
        internal string DOCFOLDER = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.CommonDocuments), @"Elton\ConnectedHome\");
        internal string USERNAME = Environment.UserName;

        private Gateway gw = new Gateway();

        internal ScriptEngine engine = new ScriptEngine();

        private void SetHide(bool hide = true)
        {
            if (hide)
            {
                this.ShowInTaskbar = false;
                this.WindowState = FormWindowState.Minimized;
                this.Hide();
            }
            else
            {
                this.ShowInTaskbar = true;
                this.Show();
                this.WindowState = FormWindowState.Normal;
            }
            //notifyIcon.Visible = hide;
        }

        public MainForm()
        {
            InitializeComponent();
        }

        private void MainForm_Load(object sender, EventArgs e)
        {
            Icon = Icon.ExtractAssociatedIcon(Application.ExecutablePath);
            notifyIcon.Icon = Icon;
            notifyIcon.Text = this.Text;
            notifyIcon.BalloonTipIcon = ToolTipIcon.Info;

            var basepath = APPFOLDER;
            if (Directory.Exists(APPFOLDER))
                basepath = APPFOLDER;
            else if (Directory.Exists(SKYFOLDER))
                basepath = SKYFOLDER;
            else if(Directory.Exists(DOCFOLDER))
                basepath = DOCFOLDER;

            var configFile = Path.Combine(basepath, "config", "aqara.json");
            if(engine is ScriptEngine)
                engine.Init(basepath, configFile, edResult);
#if DEBUG
            if (USERNAME.StartsWith("netch", StringComparison.CurrentCultureIgnoreCase)) btnTest.Visible = true;
            else btnTest.Visible = false;
            ContextMenuStrip = contextNotify;
#else
            //SetHide(true);
#endif
        }

        private void MainForm_Resize(object sender, EventArgs e)
        {
            if (FormWindowState.Minimized == this.WindowState)
            {
                SetHide(true);
            }

            else if (FormWindowState.Normal == this.WindowState)
            {
                SetHide(false);
            }
        }

        private void notifyIcon_DoubleClick(object sender, EventArgs e)
        {
            SetHide(false);
        }

        private void tsmiShowForm_Click(object sender, EventArgs e)
        {
            if(!Visible) SetHide(false);
        }

        private void tsmiExit_Click(object sender, EventArgs e)
        {
            Environment.Exit(0);
            //Application.Exit();
            //Close();
        }

        private async void btnTest_Click(object sender, EventArgs e)
        {
            if ( engine is ScriptEngine)
            {
                var ret = await engine.RunScript(true, true);
            }                
        }

        private void btnReloadScript_Click(object sender, EventArgs e)
        {
            //InitScriptEngine();
            var sf = Path.Combine(APPFOLDER, "actions.csx");
            if (File.Exists(sf) && engine is ScriptEngine)
                engine.ScriptContext = File.ReadAllText(sf);
        }

        private void btnEditScript_Click(object sender, EventArgs e)
        {
            var sf = Path.Combine(APPFOLDER, "actions.csx");
            if (File.Exists(sf))
            {
#if !DEBUG
                AutoItX.Run($"notepad2 /s cs {sf}", APPFOLDER);
#else
                var ret = AutoItX.Run($"notepad2 /s cs {sf}", APPFOLDER);
                //var ret = AutoItX.RunWait($"notepad2 /s cs {sf}", APPFOLDER);
                //if (ret == 0 && engine is ScriptEngine) engine.ScriptContext = File.ReadAllText(sf);
                //else MessageBox.Show("notepad2 run failed!");
#endif
            }
        }

        private void chkPause_CheckStateChanged(object sender, EventArgs e)
        {
            if (sender == chkPause)
            {
                if (engine is ScriptEngine) engine.Pausing = chkPause.Checked;
                tsmiPause.Checked = chkPause.Checked;
            }
            else if(sender == tsmiPause)
            {
                chkPause.Checked = tsmiPause.Checked;
            }
        }

        private void chkOnTop_CheckedChanged(object sender, EventArgs e)
        {
            if(sender == chkOnTop)
            {
                TopMost = chkOnTop.Checked;
                tsmiOnTop.Checked = chkOnTop.Checked;
            }
            else if(sender == tsmiOnTop)
            {
                chkOnTop.Checked = tsmiOnTop.Checked;
            }
        }

    }

}
