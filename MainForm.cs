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
        internal string APPFOLDER = Path.GetDirectoryName(Application.ExecutablePath);
        internal string SKYFOLDER = Path.Combine(KnownFolders.GetPath(KnownFolder.SkyDrive), @"ApplicationData\ConnectedHome");
        internal string DOCFOLDER = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.CommonDocuments), @"Elton\ConnectedHome\");
        internal string USERNAME = Environment.UserName;

        private Gateway gw = new Gateway();

        internal ScriptEngine engine = new ScriptEngine();

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
            if(engine is ScriptEngine)
                engine.Init(basepath, configFile, edResult);
#if DEBUG
            if (USERNAME.StartsWith("netch", StringComparison.CurrentCultureIgnoreCase)) btnTest.Visible = true;
            else btnTest.Visible = false;
#endif
        }

        private async void btnTest_Click(object sender, EventArgs e)
        {
            if ( engine is ScriptEngine)
                await engine.RunScript(true);
        }

        private void btnReloadScript_Click(object sender, EventArgs e)
        {
            //InitScriptEngine();
            var sf = Path.Combine(APPFOLDER, "actions.csx");
            if (File.Exists(sf) && engine is ScriptEngine)
                engine.Load(File.ReadAllText(sf));
                //engine.ScriptContext = File.ReadAllText(sf);
        }

        private void btnEditScript_Click(object sender, EventArgs e)
        {
            var sf = Path.Combine(APPFOLDER, "actions.csx");
            if (File.Exists(sf))
            {
#if !DEBUG
                AutoItX.Run($"notepad2 /s cs {sf}", APPFOLDER);
#else
                var ret = AutoItX.RunWait($"notepad2 /s cs {sf}", APPFOLDER);
                if (ret == 0 && engine is ScriptEngine) engine.Load(File.ReadAllText(sf));
                else MessageBox.Show("notepad2 run failed!");
#endif
            }
        }
    }

}
