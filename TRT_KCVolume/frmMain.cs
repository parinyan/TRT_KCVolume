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
using Newtonsoft.Json.Linq;
using ChoETL;
using Microsoft.AspNetCore.Mvc.ViewFeatures;
using Microsoft.Build.Tasks.Deployment.Bootstrapper;

namespace TRT_KCVolume
{
    public partial class frmMain : Form
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
                PathExcelExport,
                TimeSch
            ;
        bool TimeSchStatus,
            Status = false;
        private void btnUpdateVolume_Click(object sender, EventArgs e)
        {
            Process();
        }

        private void BtnConfig_Click(object sender, EventArgs e)
        {
            frmDialogConfig frmDialogConfig = new frmDialogConfig(this);
            frmDialogConfig.Show();
        }

        public frmMain()
        {
            InitializeComponent();
        }

        private void timer1_Tick(object sender, EventArgs e)
        {
            if (Status)
            {
                btnStatus.BackColor = Color.Lime;
                Status = false;
            }
            else
            {
                btnStatus.BackColor = Color.Gray;
                Status = true;
            }
            try
            {
                if (TimeSch == DateTime.Now.ToString("HH:mm") && TimeSchStatus == false)
                {
                    //Process();
                }
            }
            catch (Exception ex)
            {
                lbFile.Items.Insert(0, ex.ToString());
                log_file(ex.ToString());
            }
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
            TimeSch = Properties.Settings.Default.TimeSch;
            TimeSchStatus = Properties.Settings.Default.TimeSchStatus;
        }

        private void Main_Load(object sender, EventArgs e)
        {
            Setting();
            toolStripStatusLabel1.Text = "Version 1.1 Start Time : " + DateTime.Now.ToString();

            HttpClient clint = new HttpClient();
            clint.BaseAddress = new Uri("http://localhost/");
            HttpResponseMessage response = clint.GetAsync("api_trt.php").Result;
            var csv = response.Content.ReadAsStringAsync().Result.ToString();
            csv = csv.Replace("|", ",");
            csv = csv.Replace('"', ' ');

            StringBuilder sb = new StringBuilder();
            using (var p = ChoCSVReader.LoadText(csv)
                .WithFirstLineHeader()
                )
            {
                using (var w = new ChoJSONWriter(sb))
                    w.Write(p);
            }
            var json = DeserializeToList<TRTData>(sb.ToString());
        }

        public static List<string> InvalidJsonElements;
        public static IList<T> DeserializeToList<T>(string jsonString)
        {
            InvalidJsonElements = null;
            var array = JArray.Parse(jsonString);
            IList<T> objectsList = new List<T>();

            foreach (var item in array)
            {
                try
                {
                    // CorrectElements  
                    objectsList.Add(item.ToObject<T>());
                }
                catch (Exception ex)
                {
                    InvalidJsonElements = InvalidJsonElements ?? new List<string>();
                    InvalidJsonElements.Add(item.ToString());
                }
            }

            return objectsList;
        }

        public class TRTData
        {
            public string TRPart { get; set; }
            public string WC { get; set; }

            public string Line { get; set; }
            public string BeforeUpdate { get; set; }
            public string UpdateQty { get; set; }
            public string AfterUpdate { get; set; }
            public string ProductionTime { get; set; }
            public string ImportDate { get; set; }
            public string FileName { get; set; }
            public string SEQ { get; set; }

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
                            //New Part No
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
                                    else
                                    {
                                        xlWorkSheet.Rows[27 + i].Hidden = false;
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
                        xlWorkBookExport.Close();
                        xlAppExport.Quit();
                        return;
                    }
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
