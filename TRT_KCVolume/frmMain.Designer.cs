namespace TRT_KCVolume
{
    partial class frmMain
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.components = new System.ComponentModel.Container();
            this.BtnConfig = new System.Windows.Forms.Button();
            this.btnStatus = new System.Windows.Forms.Button();
            this.statusStrip1 = new System.Windows.Forms.StatusStrip();
            this.toolStripStatusLabel1 = new System.Windows.Forms.ToolStripStatusLabel();
            this.lbFile = new System.Windows.Forms.ListBox();
            this.btnUpdateVolume = new System.Windows.Forms.Button();
            this.timer1 = new System.Windows.Forms.Timer(this.components);
            this.statusStrip1.SuspendLayout();
            this.SuspendLayout();
            // 
            // BtnConfig
            // 
            this.BtnConfig.Location = new System.Drawing.Point(18, 1);
            this.BtnConfig.Name = "BtnConfig";
            this.BtnConfig.Size = new System.Drawing.Size(75, 32);
            this.BtnConfig.TabIndex = 15;
            this.BtnConfig.Text = "Config";
            this.BtnConfig.UseVisualStyleBackColor = true;
            this.BtnConfig.Click += new System.EventHandler(this.BtnConfig_Click);
            // 
            // btnStatus
            // 
            this.btnStatus.Location = new System.Drawing.Point(18, 207);
            this.btnStatus.Name = "btnStatus";
            this.btnStatus.Size = new System.Drawing.Size(24, 23);
            this.btnStatus.TabIndex = 14;
            this.btnStatus.UseVisualStyleBackColor = true;
            // 
            // statusStrip1
            // 
            this.statusStrip1.ImageScalingSize = new System.Drawing.Size(20, 20);
            this.statusStrip1.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.toolStripStatusLabel1});
            this.statusStrip1.Location = new System.Drawing.Point(0, 238);
            this.statusStrip1.Name = "statusStrip1";
            this.statusStrip1.Size = new System.Drawing.Size(1012, 26);
            this.statusStrip1.TabIndex = 13;
            this.statusStrip1.Text = "statusStrip1";
            // 
            // toolStripStatusLabel1
            // 
            this.toolStripStatusLabel1.Name = "toolStripStatusLabel1";
            this.toolStripStatusLabel1.Size = new System.Drawing.Size(151, 20);
            this.toolStripStatusLabel1.Text = "toolStripStatusLabel1";
            // 
            // lbFile
            // 
            this.lbFile.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.lbFile.FormattingEnabled = true;
            this.lbFile.ItemHeight = 16;
            this.lbFile.Location = new System.Drawing.Point(18, 37);
            this.lbFile.Margin = new System.Windows.Forms.Padding(4);
            this.lbFile.Name = "lbFile";
            this.lbFile.Size = new System.Drawing.Size(982, 164);
            this.lbFile.TabIndex = 12;
            // 
            // btnUpdateVolume
            // 
            this.btnUpdateVolume.BackColor = System.Drawing.Color.Lime;
            this.btnUpdateVolume.Location = new System.Drawing.Point(877, 208);
            this.btnUpdateVolume.Name = "btnUpdateVolume";
            this.btnUpdateVolume.Size = new System.Drawing.Size(123, 27);
            this.btnUpdateVolume.TabIndex = 11;
            this.btnUpdateVolume.Text = "Update Volume";
            this.btnUpdateVolume.UseVisualStyleBackColor = false;
            this.btnUpdateVolume.Click += new System.EventHandler(this.btnUpdateVolume_Click);
            // 
            // timer1
            // 
            this.timer1.Enabled = true;
            this.timer1.Interval = 5000;
            this.timer1.Tick += new System.EventHandler(this.timer1_Tick);
            // 
            // frmMain
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 16F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1012, 264);
            this.Controls.Add(this.BtnConfig);
            this.Controls.Add(this.btnStatus);
            this.Controls.Add(this.statusStrip1);
            this.Controls.Add(this.lbFile);
            this.Controls.Add(this.btnUpdateVolume);
            this.Name = "frmMain";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "KC Production Volume";
            this.Load += new System.EventHandler(this.Main_Load);
            this.statusStrip1.ResumeLayout(false);
            this.statusStrip1.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button BtnConfig;
        private System.Windows.Forms.Button btnStatus;
        private System.Windows.Forms.StatusStrip statusStrip1;
        private System.Windows.Forms.ToolStripStatusLabel toolStripStatusLabel1;
        private System.Windows.Forms.ListBox lbFile;
        private System.Windows.Forms.Button btnUpdateVolume;
        private System.Windows.Forms.Timer timer1;
    }
}

