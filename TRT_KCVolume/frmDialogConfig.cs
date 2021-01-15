using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace TRT_KCVolume
{
    public partial class frmDialogConfig : Form
    {
        frmMain frmMain;
        public frmDialogConfig(frmMain _frmMain)
        {
            InitializeComponent();
            frmMain = _frmMain;
        }

        private void btnSave_Click(object sender, EventArgs e)
        {
            try
            {
                Properties.Settings.Default.DBServer = txtDBServer.Text;
                Properties.Settings.Default.DBName = txtDBName.Text;
                Properties.Settings.Default.DBUser = txtDBUser.Text;
                Properties.Settings.Default.DBPassword = txtDBPassword.Text;
                Properties.Settings.Default.Host = txtHost.Text;
                Properties.Settings.Default.Email = txtEmail.Text;
                Properties.Settings.Default.Password = txtPassword.Text;
                Properties.Settings.Default.Port = txtPort.Text;
                Properties.Settings.Default.SenderDisplay = txtSender.Text;
                Properties.Settings.Default.ShiftAStart = dtpShiftAStart.Value.ToString("HH:mm");
                Properties.Settings.Default.ShiftAEnd = dtpShiftAEnd.Value.ToString("HH:mm");
                Properties.Settings.Default.ShiftBStart = dtpShiftBStart.Value.ToString("HH:mm");
                Properties.Settings.Default.ShiftBEnd = dtpShiftBEnd.Value.ToString("HH:mm");
                Properties.Settings.Default.PathExcelTemplateSource = txtFolderSource.Text;
                Properties.Settings.Default.PathExcelExport = txtFolderExport.Text;
                Properties.Settings.Default.TimeSch = dtpTimeSch.Value.ToString("HH:mm");
                Properties.Settings.Default.TimeSchStatus = false;
                frmMain.Setting();
                MessageBox.Show("Successfully saved", "Information", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString(), "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }           
        }

        private void DialogConfig_Load(object sender, EventArgs e)
        {
            txtDBServer.Text = Properties.Settings.Default.DBServer;
            txtDBName.Text = Properties.Settings.Default.DBName;
            txtDBUser.Text = Properties.Settings.Default.DBUser;
            txtDBPassword.Text = Properties.Settings.Default.DBPassword;
            txtHost.Text = Properties.Settings.Default.Host;
            txtEmail.Text = Properties.Settings.Default.Email;
            txtPassword.Text  = Properties.Settings.Default.Password;
            txtPort.Text = Properties.Settings.Default.Port;
            txtSender.Text = Properties.Settings.Default.SenderDisplay;
            dtpShiftAStart.Value = DateTime.ParseExact(Properties.Settings.Default.ShiftAStart, "HH:mm", CultureInfo.InvariantCulture);
            dtpShiftAEnd.Value = DateTime.ParseExact(Properties.Settings.Default.ShiftAEnd, "HH:mm", CultureInfo.InvariantCulture);
            dtpShiftBStart.Value = DateTime.ParseExact(Properties.Settings.Default.ShiftBStart, "HH:mm", CultureInfo.InvariantCulture);
            dtpShiftBEnd.Value = DateTime.ParseExact(Properties.Settings.Default.ShiftBEnd, "HH:mm", CultureInfo.InvariantCulture);
            txtFolderSource.Text = Properties.Settings.Default.PathExcelTemplateSource;
            txtFolderExport.Text = Properties.Settings.Default.PathExcelExport;
            dtpTimeSch.Value = DateTime.ParseExact(Properties.Settings.Default.TimeSch, "HH:mm", CultureInfo.InvariantCulture);
        }
    }
}
