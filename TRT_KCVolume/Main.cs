using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using System.Diagnostics;
using System.Globalization;
using Microsoft.Office.Interop.Excel;
using Application = System.Windows.Forms.Application;
using System.Net.Http;
using Newtonsoft.Json;

namespace TRT_KCVolume
{
    public partial class Main : Form
    {
        string DBServer,
                DBName,
                DBUser,
                DBPassword,
                Host,
                Email,
                Password,
                Port,
                SenderDisplay,
                ShiftAStart,
                ShiftAEnd,
                ShiftBStart,
                ShiftBEnd,
                PathExcelTemplateSource,
                PathExcelExport;

        private void btnUpdateVolume_Click(object sender, EventArgs e)
        {
            Process();
        }

        private void BtnConfig_Click(object sender, EventArgs e)
        {
            new DialogConfig().Show();
        }

        public Main()
        {
            InitializeComponent();
        }

        private void timer1_Tick(object sender, EventArgs e)
        {

        }

        public void Setting()
        {
            DBServer = Properties.Settings.Default.DBServer;
            DBName = Properties.Settings.Default.DBName;
            DBUser = Properties.Settings.Default.DBUser;
            DBPassword = Properties.Settings.Default.DBPassword;
            Host = Properties.Settings.Default.Host;
            Email = Properties.Settings.Default.Email;
            Password = Properties.Settings.Default.Password;
            Port = Properties.Settings.Default.Port;
            SenderDisplay = Properties.Settings.Default.SenderDisplay;
            ShiftAStart = Properties.Settings.Default.ShiftAStart;
            ShiftAEnd = Properties.Settings.Default.ShiftAEnd;
            ShiftBStart = Properties.Settings.Default.ShiftBStart;
            ShiftBEnd = Properties.Settings.Default.ShiftBEnd;
            PathExcelTemplateSource = Properties.Settings.Default.PathExcelTemplateSource;
            PathExcelExport = Properties.Settings.Default.PathExcelExport;
        }

        private void Main_Load(object sender, EventArgs e)
        {
            Setting();
            toolStripStatusLabel1.Text = "Version 1.1 Start Time : " + DateTime.Now.ToString();
        }

        public class ProductionData
        {
            public string PartNo { get; set; }
            public string Qty { get; set; }

        }

        public class Machine
        {
            public string MachineNo { get; set; }
            
            public List<ProductionData> data { get; set; }
        }

        public class Production
        { 
            public List<Machine> Machine { get; set; }
            public string Shift { get; set; }
        }

        public class ApiTRT
        {
            public List<Production> Production { get; set; }
            public string Month { get; set; }
            public string Day { get; set; }
        }



        private void Process()
        {
            string strShift, strMonth, strFileName;
            string[] arrPathExcelTemplateSource = Directory.GetDirectories(PathExcelTemplateSource);
            //Source
            foreach (string Source in arrPathExcelTemplateSource)
            {
                //Shift
                strShift = Path.GetFileName(Source);
                string[] arrShift = Directory.GetDirectories(Source);
                foreach (string Shift in arrShift)
                {
                    //Files
                    strMonth = Path.GetFileName(Shift);
                    string MonthCurrent = DateTime.Now.ToString("MMM", CultureInfo.InvariantCulture);
                    if (MonthCurrent == strMonth)
                    {
                        string[] arrFiles = Directory.GetFiles(Shift);
                        foreach (string FilesSource in arrFiles)
                        {
                            strFileName = Path.GetFileNameWithoutExtension(FilesSource);
                            string strFileType = Path.GetExtension(FilesSource);
                            if ((strFileType == ".xlsx" || strFileType == ".xls") && strFileName.IndexOf("~$") < 0)
                            {
                                ExportExcel(PathExcelExport, FilesSource, strShift, strMonth, strFileName, strFileType);
                            }
                        }
                    }
                }
            }
        }

        private void ExportExcel(string PathExcelExport, string FileSource, string strShift, string strMonth, string strFileName, string strFileType)
        {
            //try
            //{
                int DaysInMonth = DateTime.DaysInMonth(DateTime.Now.Year, DateTime.Now.Month);
                int DayCurrent = DateTime.Now.Day;
                Excel.Application xlApp = new Excel.Application();
                Excel.Workbook xlWorkBook = xlApp.Workbooks.Open(FileSource);
                Excel.Worksheet xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);
                
                if (xlWorkBook.ReadOnly)
                {
                    lbFile.Items.Insert(0, "Update Error : " + strFileName + strFileType + " Source Is Open!! " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"));
                    log_file("Update Error : " + strFileName + strFileType + " Source Is Open!! " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"));
                    KillProcessExcel();
                    return;
                }

                HttpClient clint = new HttpClient();
                clint.BaseAddress = new Uri("http://localhost/");
                HttpResponseMessage response = clint.GetAsync("api_trt.php").Result;
                var fileJsonString = response.Content.ReadAsStringAsync().Result.ToString();
                var result = JsonConvert.DeserializeObject<ApiTRT>(fileJsonString);

                var FilterShift = result.Production.Where(x => x.Shift == strShift).ToList();
                foreach (var Shift in FilterShift)
                {
                    var FilterMachine = Shift.Machine.Where(x => x.MachineNo == strFileName).ToList();
                    foreach (var Machine in FilterMachine)
                    {
                        foreach (var data in Machine.data)
                        {
                            Range rng = xlWorkSheet.get_Range("D28", "D319");
                            Range findRng = rng.Find(data.PartNo);

                            if (findRng is null)
                            {
                                for (int i = 1; i <= 300; i++)
                                {
                                    var PartNo = (xlWorkSheet.Cells[27+i, 4] as Range).Value;
                                    if (PartNo == null )
                                    {
                                        xlWorkSheet.Cells[27 + i, 4] = data.PartNo;
                                        xlWorkSheet.Cells[27 + i, 9 + DayCurrent] = data.Qty;
                                        for (int l = 1; l <= DaysInMonth + 8; l++)
                                        {
                                            xlWorkSheet.Cells[27 + i, 1 + l].Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Red);
                                        }
                                        break;
                                    }
                                }
                            }
                            else
                            {
                                xlWorkSheet.Cells[findRng.Row, 9 + DayCurrent] = data.Qty;
                            }
                            //excel set row hide
                            for (int i = 1; i <= 300; i++)
                            {
                                try
                                {
                                    var STDManhour = (xlWorkSheet.Cells[27 + i, 5] as Range).Value.ToString();
                                    if (STDManhour == "0")
                                    {
                                        xlWorkSheet.Rows[27 + i].Hidden = true;
                                    }
                                }
                                catch (Exception)
                                {
                                    xlWorkSheet.Rows[27 + i].Hidden = true;
                                }
                                
                            }

                            //excel set this day
                            for (int i = 8; i <= 546; i++)
                            {
                                xlWorkSheet.Cells[i, 9 + DayCurrent].Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.BlueViolet);
                            }

                            //insert db 
                        }
                    }
                }

                xlApp.DisplayAlerts = false;

                fnCreateDirectory(PathExcelExport + @"\" + strShift);
                fnCreateDirectory(PathExcelExport + @"\" + strShift + @"\" + strMonth);
                string PathExport = PathExcelExport + @"\" + strShift + @"\" + strMonth + @"\" + strFileName + strFileType;

                try
                {
                    Excel.Application xlAppExport = new Excel.Application();
                    Excel.Workbook xlWorkBookExport = xlAppExport.Workbooks.Open(PathExport);

                    if (xlWorkBookExport.ReadOnly)
                    {
                        lbFile.Items.Insert(0, "Update Error : " + strFileName + strFileType + " Destination Is Open!! " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"));
                        log_file("Update Error : " + strFileName + strFileType + " Destination Is Open!! " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"));
                        KillProcessExcel();
                        return;
                    }

                    xlWorkBookExport.Close();
                    xlAppExport.Quit();
                }
                catch (Exception)
                {
                }
               

                xlWorkSheet.SaveAs(PathExport);

                xlWorkBook.Close();
                xlApp.Quit();

                KillProcessExcel();

                lbFile.Items.Insert(0, "Update Success : " + strFileName + strFileType + " " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"));
                log_file("Update Success : " + strFileName + strFileType + " " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"));
            //}
            //catch (Exception ex)
            //{

            //    lbFile.Items.Insert(0, "Volume Error : " + strFileName + strFileType + " " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"));
            //    log_file("Volume Error : " + strFileName + strFileType + " " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"));
            //}

        }

        

        private void fnCreateDirectory(string Path)
        {
            bool exists = Directory.Exists(Path);
            if (!exists)
            {
                Directory.CreateDirectory(Path);
            }
        }

        private void log_file(string text)
        {
            StringBuilder sb = new StringBuilder();
            sb.Append(text);
            File.AppendAllText(Path.GetDirectoryName(Application.ExecutablePath) + "\\log.txt", sb.ToString() + Environment.NewLine);
            sb.Clear();
        }

        private void KillProcessExcel()
        {
            Process[] process2 = System.Diagnostics.Process.GetProcessesByName("Excel");
            foreach (Process p in process2)
            {
                try
                {
                    if (p.MainWindowTitle == "")
                    {
                        p.Kill();
                    }
                }
                catch { }
            }
        }
    }
}
