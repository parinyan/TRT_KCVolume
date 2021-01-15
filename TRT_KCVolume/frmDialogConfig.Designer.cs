namespace TRT_KCVolume
{
    partial class frmDialogConfig
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
            this.label8 = new System.Windows.Forms.Label();
            this.label6 = new System.Windows.Forms.Label();
            this.dtpShiftBStart = new System.Windows.Forms.DateTimePicker();
            this.dtpShiftAStart = new System.Windows.Forms.DateTimePicker();
            this.label10 = new System.Windows.Forms.Label();
            this.label11 = new System.Windows.Forms.Label();
            this.label9 = new System.Windows.Forms.Label();
            this.label4 = new System.Windows.Forms.Label();
            this.txtDBServer = new System.Windows.Forms.TextBox();
            this.txtDBUser = new System.Windows.Forms.TextBox();
            this.label3 = new System.Windows.Forms.Label();
            this.label1 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.label5 = new System.Windows.Forms.Label();
            this.txtDBPassword = new System.Windows.Forms.TextBox();
            this.txtHost = new System.Windows.Forms.TextBox();
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.label7 = new System.Windows.Forms.Label();
            this.txtDBName = new System.Windows.Forms.TextBox();
            this.txtEmail = new System.Windows.Forms.TextBox();
            this.txtSender = new System.Windows.Forms.TextBox();
            this.txtPassword = new System.Windows.Forms.TextBox();
            this.txtPort = new System.Windows.Forms.TextBox();
            this.btnSave = new System.Windows.Forms.Button();
            this.groupBox2 = new System.Windows.Forms.GroupBox();
            this.dtpTimeSch = new System.Windows.Forms.DateTimePicker();
            this.label12 = new System.Windows.Forms.Label();
            this.dtpShiftBEnd = new System.Windows.Forms.DateTimePicker();
            this.dtpShiftAEnd = new System.Windows.Forms.DateTimePicker();
            this.timer1 = new System.Windows.Forms.Timer(this.components);
            this.groupBox3 = new System.Windows.Forms.GroupBox();
            this.label14 = new System.Windows.Forms.Label();
            this.label15 = new System.Windows.Forms.Label();
            this.txtFolderSource = new System.Windows.Forms.TextBox();
            this.txtFolderExport = new System.Windows.Forms.TextBox();
            this.groupBox1.SuspendLayout();
            this.groupBox2.SuspendLayout();
            this.groupBox3.SuspendLayout();
            this.SuspendLayout();
            // 
            // label8
            // 
            this.label8.AutoSize = true;
            this.label8.Font = new System.Drawing.Font("Microsoft Sans Serif", 10.2F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(222)));
            this.label8.Location = new System.Drawing.Point(68, 136);
            this.label8.Name = "label8";
            this.label8.Size = new System.Drawing.Size(118, 20);
            this.label8.TabIndex = 9;
            this.label8.Text = "DBPassword :";
            // 
            // label6
            // 
            this.label6.AutoSize = true;
            this.label6.Font = new System.Drawing.Font("Microsoft Sans Serif", 10.2F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(222)));
            this.label6.Location = new System.Drawing.Point(98, 67);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(88, 20);
            this.label6.TabIndex = 7;
            this.label6.Text = "DBName :";
            // 
            // dtpShiftBStart
            // 
            this.dtpShiftBStart.CustomFormat = "HH:mm";
            this.dtpShiftBStart.Format = System.Windows.Forms.DateTimePickerFormat.Custom;
            this.dtpShiftBStart.Location = new System.Drawing.Point(205, 245);
            this.dtpShiftBStart.Name = "dtpShiftBStart";
            this.dtpShiftBStart.ShowUpDown = true;
            this.dtpShiftBStart.Size = new System.Drawing.Size(99, 22);
            this.dtpShiftBStart.TabIndex = 12;
            this.dtpShiftBStart.Value = new System.DateTime(2021, 1, 2, 20, 0, 0, 0);
            // 
            // dtpShiftAStart
            // 
            this.dtpShiftAStart.CustomFormat = "HH:mm";
            this.dtpShiftAStart.Format = System.Windows.Forms.DateTimePickerFormat.Custom;
            this.dtpShiftAStart.Location = new System.Drawing.Point(205, 210);
            this.dtpShiftAStart.Name = "dtpShiftAStart";
            this.dtpShiftAStart.ShowUpDown = true;
            this.dtpShiftAStart.Size = new System.Drawing.Size(99, 22);
            this.dtpShiftAStart.TabIndex = 11;
            this.dtpShiftAStart.Value = new System.DateTime(2021, 1, 2, 8, 0, 0, 0);
            // 
            // label10
            // 
            this.label10.AutoSize = true;
            this.label10.Font = new System.Drawing.Font("Microsoft Sans Serif", 10.2F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(222)));
            this.label10.Location = new System.Drawing.Point(120, 210);
            this.label10.Name = "label10";
            this.label10.Size = new System.Drawing.Size(69, 20);
            this.label10.TabIndex = 10;
            this.label10.Text = "Shift A :";
            // 
            // label11
            // 
            this.label11.AutoSize = true;
            this.label11.Font = new System.Drawing.Font("Microsoft Sans Serif", 10.2F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(222)));
            this.label11.Location = new System.Drawing.Point(120, 247);
            this.label11.Name = "label11";
            this.label11.Size = new System.Drawing.Size(70, 20);
            this.label11.TabIndex = 9;
            this.label11.Text = "Shift B :";
            // 
            // label9
            // 
            this.label9.AutoSize = true;
            this.label9.Font = new System.Drawing.Font("Microsoft Sans Serif", 10.2F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(222)));
            this.label9.Location = new System.Drawing.Point(136, 136);
            this.label9.Name = "label9";
            this.label9.Size = new System.Drawing.Size(50, 20);
            this.label9.TabIndex = 8;
            this.label9.Text = "Port :";
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Font = new System.Drawing.Font("Microsoft Sans Serif", 10.2F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(222)));
            this.label4.Location = new System.Drawing.Point(93, 100);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(93, 20);
            this.label4.TabIndex = 7;
            this.label4.Text = "Password :";
            // 
            // txtDBServer
            // 
            this.txtDBServer.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(222)));
            this.txtDBServer.Location = new System.Drawing.Point(205, 24);
            this.txtDBServer.Name = "txtDBServer";
            this.txtDBServer.Size = new System.Drawing.Size(307, 30);
            this.txtDBServer.TabIndex = 0;
            // 
            // txtDBUser
            // 
            this.txtDBUser.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(222)));
            this.txtDBUser.Location = new System.Drawing.Point(205, 96);
            this.txtDBUser.Name = "txtDBUser";
            this.txtDBUser.Size = new System.Drawing.Size(307, 30);
            this.txtDBUser.TabIndex = 2;
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Font = new System.Drawing.Font("Microsoft Sans Serif", 10.2F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(222)));
            this.label3.Location = new System.Drawing.Point(125, 64);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(61, 20);
            this.label3.TabIndex = 6;
            this.label3.Text = "Email :";
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("Microsoft Sans Serif", 10.2F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(222)));
            this.label1.Location = new System.Drawing.Point(58, 172);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(128, 20);
            this.label1.TabIndex = 5;
            this.label1.Text = "SenderDisplay :";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Font = new System.Drawing.Font("Microsoft Sans Serif", 10.2F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(222)));
            this.label2.Location = new System.Drawing.Point(131, 28);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(55, 20);
            this.label2.TabIndex = 5;
            this.label2.Text = "Host :";
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Font = new System.Drawing.Font("Microsoft Sans Serif", 10.2F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(222)));
            this.label5.Location = new System.Drawing.Point(93, 31);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(93, 20);
            this.label5.TabIndex = 6;
            this.label5.Text = "DBServer :";
            // 
            // txtDBPassword
            // 
            this.txtDBPassword.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(222)));
            this.txtDBPassword.Location = new System.Drawing.Point(205, 132);
            this.txtDBPassword.Name = "txtDBPassword";
            this.txtDBPassword.Size = new System.Drawing.Size(307, 30);
            this.txtDBPassword.TabIndex = 3;
            // 
            // txtHost
            // 
            this.txtHost.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(222)));
            this.txtHost.Location = new System.Drawing.Point(205, 21);
            this.txtHost.Name = "txtHost";
            this.txtHost.Size = new System.Drawing.Size(307, 30);
            this.txtHost.TabIndex = 0;
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.label8);
            this.groupBox1.Controls.Add(this.label7);
            this.groupBox1.Controls.Add(this.label6);
            this.groupBox1.Controls.Add(this.label5);
            this.groupBox1.Controls.Add(this.txtDBServer);
            this.groupBox1.Controls.Add(this.txtDBName);
            this.groupBox1.Controls.Add(this.txtDBUser);
            this.groupBox1.Controls.Add(this.txtDBPassword);
            this.groupBox1.Location = new System.Drawing.Point(16, 12);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(518, 175);
            this.groupBox1.TabIndex = 12;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "DATABASE";
            // 
            // label7
            // 
            this.label7.AutoSize = true;
            this.label7.Font = new System.Drawing.Font("Microsoft Sans Serif", 10.2F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(222)));
            this.label7.Location = new System.Drawing.Point(106, 103);
            this.label7.Name = "label7";
            this.label7.Size = new System.Drawing.Size(80, 20);
            this.label7.TabIndex = 8;
            this.label7.Text = "DBUser :";
            // 
            // txtDBName
            // 
            this.txtDBName.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(222)));
            this.txtDBName.Location = new System.Drawing.Point(205, 60);
            this.txtDBName.Name = "txtDBName";
            this.txtDBName.Size = new System.Drawing.Size(307, 30);
            this.txtDBName.TabIndex = 1;
            // 
            // txtEmail
            // 
            this.txtEmail.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(222)));
            this.txtEmail.Location = new System.Drawing.Point(205, 57);
            this.txtEmail.Name = "txtEmail";
            this.txtEmail.Size = new System.Drawing.Size(307, 30);
            this.txtEmail.TabIndex = 1;
            // 
            // txtSender
            // 
            this.txtSender.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(222)));
            this.txtSender.Location = new System.Drawing.Point(205, 165);
            this.txtSender.Name = "txtSender";
            this.txtSender.Size = new System.Drawing.Size(307, 30);
            this.txtSender.TabIndex = 4;
            // 
            // txtPassword
            // 
            this.txtPassword.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(222)));
            this.txtPassword.Location = new System.Drawing.Point(205, 93);
            this.txtPassword.Name = "txtPassword";
            this.txtPassword.Size = new System.Drawing.Size(307, 30);
            this.txtPassword.TabIndex = 2;
            // 
            // txtPort
            // 
            this.txtPort.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(222)));
            this.txtPort.Location = new System.Drawing.Point(205, 129);
            this.txtPort.Name = "txtPort";
            this.txtPort.Size = new System.Drawing.Size(307, 30);
            this.txtPort.TabIndex = 3;
            // 
            // btnSave
            // 
            this.btnSave.BackColor = System.Drawing.Color.DodgerBlue;
            this.btnSave.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(222)));
            this.btnSave.Location = new System.Drawing.Point(464, 645);
            this.btnSave.Name = "btnSave";
            this.btnSave.Size = new System.Drawing.Size(75, 44);
            this.btnSave.TabIndex = 14;
            this.btnSave.Text = "SAVE";
            this.btnSave.UseVisualStyleBackColor = false;
            this.btnSave.Click += new System.EventHandler(this.btnSave_Click);
            // 
            // groupBox2
            // 
            this.groupBox2.Controls.Add(this.dtpTimeSch);
            this.groupBox2.Controls.Add(this.label12);
            this.groupBox2.Controls.Add(this.dtpShiftBEnd);
            this.groupBox2.Controls.Add(this.dtpShiftAEnd);
            this.groupBox2.Controls.Add(this.dtpShiftBStart);
            this.groupBox2.Controls.Add(this.dtpShiftAStart);
            this.groupBox2.Controls.Add(this.label10);
            this.groupBox2.Controls.Add(this.label11);
            this.groupBox2.Controls.Add(this.label9);
            this.groupBox2.Controls.Add(this.label4);
            this.groupBox2.Controls.Add(this.label3);
            this.groupBox2.Controls.Add(this.label1);
            this.groupBox2.Controls.Add(this.txtHost);
            this.groupBox2.Controls.Add(this.label2);
            this.groupBox2.Controls.Add(this.txtEmail);
            this.groupBox2.Controls.Add(this.txtSender);
            this.groupBox2.Controls.Add(this.txtPassword);
            this.groupBox2.Controls.Add(this.txtPort);
            this.groupBox2.Location = new System.Drawing.Point(16, 193);
            this.groupBox2.Name = "groupBox2";
            this.groupBox2.Size = new System.Drawing.Size(518, 321);
            this.groupBox2.TabIndex = 13;
            this.groupBox2.TabStop = false;
            this.groupBox2.Text = "EMAIL";
            // 
            // dtpTimeSch
            // 
            this.dtpTimeSch.CustomFormat = "HH:mm";
            this.dtpTimeSch.Format = System.Windows.Forms.DateTimePickerFormat.Custom;
            this.dtpTimeSch.Location = new System.Drawing.Point(204, 279);
            this.dtpTimeSch.Name = "dtpTimeSch";
            this.dtpTimeSch.ShowUpDown = true;
            this.dtpTimeSch.Size = new System.Drawing.Size(99, 22);
            this.dtpTimeSch.TabIndex = 16;
            this.dtpTimeSch.Value = new System.DateTime(2021, 1, 2, 8, 0, 0, 0);
            // 
            // label12
            // 
            this.label12.AutoSize = true;
            this.label12.Font = new System.Drawing.Font("Microsoft Sans Serif", 10.2F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(222)));
            this.label12.Location = new System.Drawing.Point(100, 281);
            this.label12.Name = "label12";
            this.label12.Size = new System.Drawing.Size(90, 20);
            this.label12.TabIndex = 15;
            this.label12.Text = "Time Sch :";
            // 
            // dtpShiftBEnd
            // 
            this.dtpShiftBEnd.CustomFormat = "HH:mm";
            this.dtpShiftBEnd.Format = System.Windows.Forms.DateTimePickerFormat.Custom;
            this.dtpShiftBEnd.Location = new System.Drawing.Point(310, 245);
            this.dtpShiftBEnd.Name = "dtpShiftBEnd";
            this.dtpShiftBEnd.ShowUpDown = true;
            this.dtpShiftBEnd.Size = new System.Drawing.Size(99, 22);
            this.dtpShiftBEnd.TabIndex = 14;
            this.dtpShiftBEnd.Value = new System.DateTime(2021, 1, 2, 8, 0, 0, 0);
            // 
            // dtpShiftAEnd
            // 
            this.dtpShiftAEnd.CustomFormat = "HH:mm";
            this.dtpShiftAEnd.Format = System.Windows.Forms.DateTimePickerFormat.Custom;
            this.dtpShiftAEnd.Location = new System.Drawing.Point(310, 210);
            this.dtpShiftAEnd.Name = "dtpShiftAEnd";
            this.dtpShiftAEnd.ShowUpDown = true;
            this.dtpShiftAEnd.Size = new System.Drawing.Size(99, 22);
            this.dtpShiftAEnd.TabIndex = 13;
            this.dtpShiftAEnd.Value = new System.DateTime(2021, 1, 2, 20, 0, 0, 0);
            // 
            // groupBox3
            // 
            this.groupBox3.Controls.Add(this.label14);
            this.groupBox3.Controls.Add(this.label15);
            this.groupBox3.Controls.Add(this.txtFolderSource);
            this.groupBox3.Controls.Add(this.txtFolderExport);
            this.groupBox3.Location = new System.Drawing.Point(16, 526);
            this.groupBox3.Name = "groupBox3";
            this.groupBox3.Size = new System.Drawing.Size(518, 112);
            this.groupBox3.TabIndex = 15;
            this.groupBox3.TabStop = false;
            this.groupBox3.Text = "EXCEL";
            // 
            // label14
            // 
            this.label14.AutoSize = true;
            this.label14.Font = new System.Drawing.Font("Microsoft Sans Serif", 10.2F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(222)));
            this.label14.Location = new System.Drawing.Point(71, 66);
            this.label14.Name = "label14";
            this.label14.Size = new System.Drawing.Size(119, 20);
            this.label14.TabIndex = 7;
            this.label14.Text = "Folder Export :";
            // 
            // label15
            // 
            this.label15.AutoSize = true;
            this.label15.Font = new System.Drawing.Font("Microsoft Sans Serif", 10.2F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(222)));
            this.label15.Location = new System.Drawing.Point(65, 30);
            this.label15.Name = "label15";
            this.label15.Size = new System.Drawing.Size(124, 20);
            this.label15.TabIndex = 6;
            this.label15.Text = "Folder Source :";
            // 
            // txtFolderSource
            // 
            this.txtFolderSource.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(222)));
            this.txtFolderSource.Location = new System.Drawing.Point(205, 24);
            this.txtFolderSource.Name = "txtFolderSource";
            this.txtFolderSource.Size = new System.Drawing.Size(307, 30);
            this.txtFolderSource.TabIndex = 0;
            // 
            // txtFolderExport
            // 
            this.txtFolderExport.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(222)));
            this.txtFolderExport.Location = new System.Drawing.Point(205, 60);
            this.txtFolderExport.Name = "txtFolderExport";
            this.txtFolderExport.Size = new System.Drawing.Size(307, 30);
            this.txtFolderExport.TabIndex = 1;
            // 
            // frmDialogConfig
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 16F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(551, 695);
            this.Controls.Add(this.groupBox3);
            this.Controls.Add(this.groupBox1);
            this.Controls.Add(this.btnSave);
            this.Controls.Add(this.groupBox2);
            this.Name = "frmDialogConfig";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Dialog Config";
            this.Load += new System.EventHandler(this.DialogConfig_Load);
            this.groupBox1.ResumeLayout(false);
            this.groupBox1.PerformLayout();
            this.groupBox2.ResumeLayout(false);
            this.groupBox2.PerformLayout();
            this.groupBox3.ResumeLayout(false);
            this.groupBox3.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Label label8;
        private System.Windows.Forms.Label label6;
        private System.Windows.Forms.DateTimePicker dtpShiftBStart;
        private System.Windows.Forms.DateTimePicker dtpShiftAStart;
        private System.Windows.Forms.Label label10;
        private System.Windows.Forms.Label label11;
        private System.Windows.Forms.Label label9;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.TextBox txtDBServer;
        private System.Windows.Forms.TextBox txtDBUser;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.TextBox txtDBPassword;
        private System.Windows.Forms.TextBox txtHost;
        private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.Label label7;
        private System.Windows.Forms.TextBox txtDBName;
        private System.Windows.Forms.TextBox txtEmail;
        private System.Windows.Forms.TextBox txtSender;
        private System.Windows.Forms.TextBox txtPassword;
        private System.Windows.Forms.TextBox txtPort;
        private System.Windows.Forms.Button btnSave;
        private System.Windows.Forms.GroupBox groupBox2;
        private System.Windows.Forms.DateTimePicker dtpShiftBEnd;
        private System.Windows.Forms.DateTimePicker dtpShiftAEnd;
        private System.Windows.Forms.Timer timer1;
        private System.Windows.Forms.GroupBox groupBox3;
        private System.Windows.Forms.Label label14;
        private System.Windows.Forms.Label label15;
        private System.Windows.Forms.TextBox txtFolderSource;
        private System.Windows.Forms.TextBox txtFolderExport;
        private System.Windows.Forms.DateTimePicker dtpTimeSch;
        private System.Windows.Forms.Label label12;
    }
}