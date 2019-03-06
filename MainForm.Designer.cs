namespace MiJia
{
    partial class MainForm
    {
        /// <summary>
        /// 必需的设计器变量。
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// 清理所有正在使用的资源。
        /// </summary>
        /// <param name="disposing">如果应释放托管资源，为 true；否则为 false。</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows 窗体设计器生成的代码

        /// <summary>
        /// 设计器支持所需的方法 - 不要修改
        /// 使用代码编辑器修改此方法的内容。
        /// </summary>
        private void InitializeComponent()
        {
            this.components = new System.ComponentModel.Container();
            this.btnTest = new System.Windows.Forms.Button();
            this.edResult = new System.Windows.Forms.TextBox();
            this.btnReloadScript = new System.Windows.Forms.Button();
            this.btnEditScript = new System.Windows.Forms.Button();
            this.chkPause = new System.Windows.Forms.CheckBox();
            this.chkOnTop = new System.Windows.Forms.CheckBox();
            this.notifyIcon = new System.Windows.Forms.NotifyIcon(this.components);
            this.contextNotify = new System.Windows.Forms.ContextMenuStrip(this.components);
            this.tsmiOnTop = new System.Windows.Forms.ToolStripMenuItem();
            this.tsmiShowForm = new System.Windows.Forms.ToolStripMenuItem();
            this.toolStripMenuItem1 = new System.Windows.Forms.ToolStripSeparator();
            this.tsmiReload = new System.Windows.Forms.ToolStripMenuItem();
            this.tsmiEdit = new System.Windows.Forms.ToolStripMenuItem();
            this.tsmiTest = new System.Windows.Forms.ToolStripMenuItem();
            this.tsmiPause = new System.Windows.Forms.ToolStripMenuItem();
            this.toolStripMenuItem2 = new System.Windows.Forms.ToolStripSeparator();
            this.tsmiExit = new System.Windows.Forms.ToolStripMenuItem();
            this.contextNotify.SuspendLayout();
            this.SuspendLayout();
            // 
            // btnTest
            // 
            this.btnTest.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.btnTest.Location = new System.Drawing.Point(12, 426);
            this.btnTest.Name = "btnTest";
            this.btnTest.Size = new System.Drawing.Size(75, 23);
            this.btnTest.TabIndex = 0;
            this.btnTest.Text = "Test Script";
            this.btnTest.UseVisualStyleBackColor = true;
            this.btnTest.Click += new System.EventHandler(this.btnTest_Click);
            // 
            // edResult
            // 
            this.edResult.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.edResult.Location = new System.Drawing.Point(12, 12);
            this.edResult.Multiline = true;
            this.edResult.Name = "edResult";
            this.edResult.ReadOnly = true;
            this.edResult.ScrollBars = System.Windows.Forms.ScrollBars.Both;
            this.edResult.Size = new System.Drawing.Size(460, 399);
            this.edResult.TabIndex = 1;
            this.edResult.WordWrap = false;
            // 
            // btnReloadScript
            // 
            this.btnReloadScript.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.btnReloadScript.Location = new System.Drawing.Point(397, 426);
            this.btnReloadScript.Name = "btnReloadScript";
            this.btnReloadScript.Size = new System.Drawing.Size(75, 23);
            this.btnReloadScript.TabIndex = 2;
            this.btnReloadScript.Text = "Reload Script";
            this.btnReloadScript.UseVisualStyleBackColor = true;
            this.btnReloadScript.Click += new System.EventHandler(this.btnReloadScript_Click);
            // 
            // btnEditScript
            // 
            this.btnEditScript.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.btnEditScript.Location = new System.Drawing.Point(316, 426);
            this.btnEditScript.Name = "btnEditScript";
            this.btnEditScript.Size = new System.Drawing.Size(75, 23);
            this.btnEditScript.TabIndex = 3;
            this.btnEditScript.Text = "Edit Script";
            this.btnEditScript.UseVisualStyleBackColor = true;
            this.btnEditScript.Click += new System.EventHandler(this.btnEditScript_Click);
            // 
            // chkPause
            // 
            this.chkPause.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.chkPause.AutoEllipsis = true;
            this.chkPause.Location = new System.Drawing.Point(238, 425);
            this.chkPause.Name = "chkPause";
            this.chkPause.Size = new System.Drawing.Size(72, 23);
            this.chkPause.TabIndex = 4;
            this.chkPause.Text = "Pausing";
            this.chkPause.UseVisualStyleBackColor = true;
            this.chkPause.CheckStateChanged += new System.EventHandler(this.chkPause_CheckStateChanged);
            // 
            // chkOnTop
            // 
            this.chkOnTop.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.chkOnTop.AutoEllipsis = true;
            this.chkOnTop.Location = new System.Drawing.Point(110, 425);
            this.chkOnTop.Name = "chkOnTop";
            this.chkOnTop.Size = new System.Drawing.Size(111, 23);
            this.chkOnTop.TabIndex = 5;
            this.chkOnTop.Text = "Always On Top";
            this.chkOnTop.UseVisualStyleBackColor = true;
            this.chkOnTop.CheckedChanged += new System.EventHandler(this.chkOnTop_CheckedChanged);
            // 
            // notifyIcon
            // 
            this.notifyIcon.ContextMenuStrip = this.contextNotify;
            this.notifyIcon.Text = "notifyIcon1";
            this.notifyIcon.Visible = true;
            this.notifyIcon.DoubleClick += new System.EventHandler(this.notifyIcon_DoubleClick);
            // 
            // contextNotify
            // 
            this.contextNotify.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.tsmiOnTop,
            this.tsmiShowForm,
            this.toolStripMenuItem1,
            this.tsmiReload,
            this.tsmiEdit,
            this.tsmiTest,
            this.tsmiPause,
            this.toolStripMenuItem2,
            this.tsmiExit});
            this.contextNotify.Name = "contextNotify";
            this.contextNotify.Size = new System.Drawing.Size(164, 170);
            // 
            // tsmiOnTop
            // 
            this.tsmiOnTop.CheckOnClick = true;
            this.tsmiOnTop.Name = "tsmiOnTop";
            this.tsmiOnTop.Size = new System.Drawing.Size(163, 22);
            this.tsmiOnTop.Text = "Always On Top";
            this.tsmiOnTop.Click += new System.EventHandler(this.chkOnTop_CheckedChanged);
            // 
            // tsmiShowForm
            // 
            this.tsmiShowForm.Name = "tsmiShowForm";
            this.tsmiShowForm.Size = new System.Drawing.Size(163, 22);
            this.tsmiShowForm.Text = "Show Window";
            this.tsmiShowForm.Click += new System.EventHandler(this.tsmiShowForm_Click);
            // 
            // toolStripMenuItem1
            // 
            this.toolStripMenuItem1.Name = "toolStripMenuItem1";
            this.toolStripMenuItem1.Size = new System.Drawing.Size(160, 6);
            // 
            // tsmiReload
            // 
            this.tsmiReload.Name = "tsmiReload";
            this.tsmiReload.Size = new System.Drawing.Size(163, 22);
            this.tsmiReload.Text = "Reload Script";
            this.tsmiReload.Click += new System.EventHandler(this.btnReloadScript_Click);
            // 
            // tsmiEdit
            // 
            this.tsmiEdit.Name = "tsmiEdit";
            this.tsmiEdit.Size = new System.Drawing.Size(163, 22);
            this.tsmiEdit.Text = "Edit Script";
            this.tsmiEdit.Click += new System.EventHandler(this.btnEditScript_Click);
            // 
            // tsmiTest
            // 
            this.tsmiTest.Name = "tsmiTest";
            this.tsmiTest.Size = new System.Drawing.Size(163, 22);
            this.tsmiTest.Text = "Test Script";
            this.tsmiTest.DoubleClick += new System.EventHandler(this.btnTest_Click);
            // 
            // tsmiPause
            // 
            this.tsmiPause.CheckOnClick = true;
            this.tsmiPause.Name = "tsmiPause";
            this.tsmiPause.Size = new System.Drawing.Size(163, 22);
            this.tsmiPause.Text = "Pausing Script";
            this.tsmiPause.Click += new System.EventHandler(this.chkPause_CheckStateChanged);
            // 
            // toolStripMenuItem2
            // 
            this.toolStripMenuItem2.Name = "toolStripMenuItem2";
            this.toolStripMenuItem2.Size = new System.Drawing.Size(160, 6);
            // 
            // tsmiExit
            // 
            this.tsmiExit.Name = "tsmiExit";
            this.tsmiExit.Size = new System.Drawing.Size(163, 22);
            this.tsmiExit.Text = "Exit";
            this.tsmiExit.Click += new System.EventHandler(this.tsmiExit_Click);
            // 
            // MainForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(484, 461);
            this.Controls.Add(this.chkOnTop);
            this.Controls.Add(this.chkPause);
            this.Controls.Add(this.btnEditScript);
            this.Controls.Add(this.btnReloadScript);
            this.Controls.Add(this.edResult);
            this.Controls.Add(this.btnTest);
            this.MinimumSize = new System.Drawing.Size(400, 300);
            this.Name = "MainForm";
            this.Text = "MiJia Gateway Monitor";
            this.FormClosing += new System.Windows.Forms.FormClosingEventHandler(this.MainForm_FormClosing);
            this.Load += new System.EventHandler(this.MainForm_Load);
            this.Resize += new System.EventHandler(this.MainForm_Resize);
            this.contextNotify.ResumeLayout(false);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button btnTest;
        private System.Windows.Forms.TextBox edResult;
        private System.Windows.Forms.Button btnReloadScript;
        private System.Windows.Forms.Button btnEditScript;
        private System.Windows.Forms.CheckBox chkPause;
        private System.Windows.Forms.CheckBox chkOnTop;
        private System.Windows.Forms.NotifyIcon notifyIcon;
        private System.Windows.Forms.ContextMenuStrip contextNotify;
        private System.Windows.Forms.ToolStripMenuItem tsmiExit;
        private System.Windows.Forms.ToolStripMenuItem tsmiShowForm;
        private System.Windows.Forms.ToolStripSeparator toolStripMenuItem1;
        private System.Windows.Forms.ToolStripMenuItem tsmiOnTop;
        private System.Windows.Forms.ToolStripMenuItem tsmiReload;
        private System.Windows.Forms.ToolStripMenuItem tsmiEdit;
        private System.Windows.Forms.ToolStripMenuItem tsmiTest;
        private System.Windows.Forms.ToolStripMenuItem tsmiPause;
        private System.Windows.Forms.ToolStripSeparator toolStripMenuItem2;
    }
}

