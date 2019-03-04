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
            this.btnTest = new System.Windows.Forms.Button();
            this.edResult = new System.Windows.Forms.TextBox();
            this.btnReloadScript = new System.Windows.Forms.Button();
            this.btnEditScript = new System.Windows.Forms.Button();
            this.chkPause = new System.Windows.Forms.CheckBox();
            this.chkOnTop = new System.Windows.Forms.CheckBox();
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
            this.DoubleBuffered = true;
            this.MinimumSize = new System.Drawing.Size(400, 300);
            this.Name = "MainForm";
            this.Text = "MiJia Gateway Monitor";
            this.Load += new System.EventHandler(this.MainForm_Load);
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
    }
}

