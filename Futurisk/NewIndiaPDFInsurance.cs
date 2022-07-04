using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using Aspose.Pdf;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System.Text;
using System.Threading.Tasks;
using Smartreader_DLL;
using System.Windows.Forms;
using Spire.Xls;
using System.Data.SqlClient;
using System.Configuration;
using Syncfusion.XlsIO;
using System.Globalization;
using System.Text.RegularExpressions;
using IWorkbook = NPOI.SS.UserModel.IWorkbook;
using NPOI.HSSF.UserModel;

namespace Futurisk
{
    public partial class NewIndiaPDFInsurance : Form
    {
        static string Dir, Filepath, RDate, Result, FormatResult;
        static string Filename, Filenamewithext, Filewithext, name, TranID, BatchID, NoRecord;

        private void btnCancel_Click(object sender, EventArgs e)
        {
            Filepath = "";
            Dir = "";
            Filename = "";
            txtfile.Text = "";
            linkLabel2.Text = "";
            btnBrowse.Enabled = false;
            btnCancel.Enabled = false;
            btnConvert.Enabled = false;
            linkLabel2.Enabled = false;
            txtfile.Text = "Select pdf document";
            txtfile.ForeColor = System.Drawing.Color.Gray;
            lblmsg1.Text = "";
            lblSuccMsg.Text = "";
            DDInsurance.SelectedValue = 0;
            DDLocation.SelectedValue = 0;
            DDsales.SelectedValue = 0;
            DDService.SelectedValue = 0;
            DDSupport.SelectedValue = 0;
            DDMonth.SelectedIndex = 0;
            DDInsurance.Enabled = true;
            DDLocation.Enabled = true;
            DDsales.Enabled = true;
            DDService.Enabled = true;
            DDSupport.Enabled = true;
            DDMonth.Enabled = true;
            DDMonth.Select();
        }

        private void DDService_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (DDService.SelectedValue.ToString() != "0")
            {
                btnBrowse.Enabled = true;
            }
            else
            {
                btnBrowse.Enabled = false;
            }
        }

        private void DDLocation_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (DDLocation.SelectedValue.ToString() == "05")
            {
                DDsales.SelectedValue = "S74";
                DDService.SelectedValue = "R57";
            }
            else
            {
                DDsales.SelectedValue = "0";
                DDService.SelectedValue = "0";
            }
        }

        private string strconn = ConfigurationManager.ConnectionStrings["IDP"].ToString();

        public NewIndiaPDFInsurance()
        {
            InitializeComponent();
            lblUser.Text = LoginInfo.UserID;
            BindDDInsurance();
            BindDDSales();
            BindDDService();
            BindDDLocation();
            BindDDSupport();
            DDMonth.Select();
            TimeUpdater(); 
            if (Fileinfo.ReportId == "NIP1")
            {
                lblHeader.Text = "NIAC (New India Assurance Company Ltd.) - Report Id: NIP1";
            }
            else if (Fileinfo.ReportId == "NIP2")
            {
                lblHeader.Text = "NIAC (New India Assurance Company Ltd.) - Report Id: NIP2";
            }
            else if (Fileinfo.ReportId == "NIP3")
            {
                lblHeader.Text = "NIAC (New India Assurance Company Ltd.) - Report Id: NIP3";
            }
        }
        async void TimeUpdater()
        {
            while (true)
            {
                lblTimer.Text = DateTime.Now.ToString("dd-MM-yyyy hh:mm:ss tt");
                await Task.Delay(1000);
            }
        }

        private void btnConvert_Click(object sender, EventArgs e)
        {
            try
            {
                btnCancel.Enabled = false;
                var confirmResult = MessageBox.Show("Have you got the right document?", "Confirm",
                                     MessageBoxButtons.YesNo);
                if (confirmResult == DialogResult.Yes)
                {
                    DateTime a = DateTime.Now;
                    lblmsg.Text = "Please wait.......";
                    lblmsg.Refresh();
                    btnBrowse.Enabled = false;
                    btnConvert.Enabled = false;
                    linkLabel2.Enabled = false;
                    DDInsurance.Enabled = false;
                    DDLocation.Enabled = false;
                    DDsales.Enabled = false;
                    DDService.Enabled = false;
                    DDSupport.Enabled = false;
                    DDMonth.Enabled = false;

                    DDInsurance.SelectionLength = 0;
                    DDLocation.SelectionLength = 0;
                    DDsales.SelectionLength = 0;
                    DDService.SelectionLength = 0;
                    DDSupport.SelectionLength = 0;
                    DDMonth.SelectionLength = 0;

                    string pathToPdf = Filepath;
                    name = Path.ChangeExtension(pathToPdf, "_"+DateTime.Now.AddDays(0).ToString("ddMMyyyyhhmmss")+".xls");

                    // Here we have our PDF and Excel docs as byte arrays
                    byte[] pdf = File.ReadAllBytes(pathToPdf);
                    byte[] xls = null;

                    // Convert PDF document to Excel workbook in memory
                    SautinSoft.PdfFocus f = new SautinSoft.PdfFocus();

                    System.Globalization.CultureInfo ci = new System.Globalization.CultureInfo("en-US");
                    ci.NumberFormat.NumberDecimalSeparator = ",";
                    ci.NumberFormat.NumberGroupSeparator = ".";
                    f.ExcelOptions.CultureInfo = ci;
                    f.ExcelOptions.SingleSheet = true;
                    f.ExcelOptions.ConvertNonTabularDataToSpreadsheet = false;
                    f.OpenPdf(pdf);

                    SQLProcs sql = new SQLProcs();
                    DataSet ds1 = new DataSet();
                    ds1 = sql.SQLExecuteDataset("SP_NewIndiaPDFTransaction",
                                  new SqlParameter { ParameterName = "@Imode", Value = 7 },
                                  new SqlParameter { ParameterName = "@Filename", Value = Fileinfo.Filename }
                         );
                    if (ds1 != null && ds1.Tables.Count > 0 && ds1.Tables[0].Rows.Count > 0)
                    {
                        Result = ds1.Tables[0].Rows[0]["Result"].ToString();
                    }
                    else
                    {
                        Result = "Not Exists";
                    }
                    if (Result == "Not Exists")
                    {
                        if (f.PageCount > 0)
                        {
                            xls = f.ToExcel();

                            //Save Excel workbook to a file in order to show it
                            if (xls != null)
                            {
                                File.WriteAllBytes(name, xls);
                                //System.Diagnostics.Process.Start(pathToExcel);
                            }
                        }
                        if (xls != null)
                        {
                            if (Fileinfo.ReportId == "NIP2")
                            {
                                IWorkbook workbook = OpenWorkBook(name);

                                ISheet sheet1 = workbook.GetSheet("AllPages");

                                DeleteRows(sheet1);

                                SaveWorkBook(workbook, name);
                            }

                            DataSet ds = new DataSet();
                            ds = sql.SQLExecuteDataset("SP_NewIndiaPDFTransaction",
                                         new SqlParameter { ParameterName = "@Imode", Value = 2 }
                                );
                            if (ds != null && ds.Tables.Count > 0 && ds.Tables[0].Rows.Count > 0)
                            {
                                TranID = ds.Tables[0].Rows[0]["TranID"].ToString();
                            }
                            else
                            {
                                TranID = "0";
                            }

                            //RDate = dateTimePicker1.Value.ToString("dd-MM-yyyy");

                            string Insurance = DDInsurance.Text.ToString();
                            string Salesby = DDsales.Text.ToString();
                            string Serviceby = DDService.Text.ToString();
                            string location = DDLocation.Text.ToString();
                            string Support = DDSupport.Text.ToString();
                            string Rmonth = DDMonth.Text.ToString();

                            RDate = DateTime.Now.AddDays(0).ToString("dd-MM-yyyy");
                            Microsoft.Office.Interop.Excel.Application oExcel = new Microsoft.Office.Interop.Excel.Application();
                            Microsoft.Office.Interop.Excel.Workbook WB = oExcel.Workbooks.Open(name);

                            //Microsoft.Office.Interop.Excel.Worksheet sheet = WB.Sheets["AllPages"];
                            Microsoft.Office.Interop.Excel.Worksheet sheet = WB.ActiveSheet;
                            if (Fileinfo.ReportId == "NIP1")
                            {
                                //InsertTransaction1(WB, TranID, RDate, Insurance, Salesby, Serviceby, location, Support, Rmonth);
                                FormatResult = NewIndiaPDFInsurancedll.InsertTransaction1(WB, TranID, RDate, Insurance, Salesby, Serviceby, location, Support, Rmonth, strconn);
                            }
                            else if (Fileinfo.ReportId == "NIP2")
                            {
                                //InsertTransaction2(WB, TranID, RDate, Insurance, Salesby, Serviceby, location, Support, Rmonth);
                                FormatResult = NewIndiaPDFInsurancedll.InsertTransaction2(WB, TranID, RDate, Insurance, Salesby, Serviceby, location, Support, Rmonth, strconn);
                            }
                            else if (Fileinfo.ReportId == "NIP3")
                            {
                                //InsertTransaction3(WB, TranID, RDate, Insurance, Salesby, Serviceby, location, Support, Rmonth);
                               FormatResult = NewIndiaPDFInsurancedll.InsertTransaction3(WB, TranID, RDate, Insurance, Salesby, Serviceby, location, Support, Rmonth, strconn);
                            }
                            if (FormatResult == "OK")
                            {
                                GetBachid_NoRecord();
                                DateTime b = DateTime.Now;
                                TimeSpan diff = b - a;
                                var Sec = String.Format("{0}", diff.Seconds);
                                lblmsg.Text = "";
                                lblSuccMsg.Text = "          SmartRead Done in " + Sec + " Seconds.\n" +
                                                     "       Batch ID: " + BatchID + " ,Number of records: " + NoRecord;
                                linkLabel2.Enabled = true;
                                linkLabel2.Text = "Click here to Edit the records if needed.";
                                oExcel.Workbooks.Close();
                                DDInsurance.SelectionLength = 0;
                                DDLocation.SelectionLength = 0;
                                DDsales.SelectionLength = 0;
                                DDService.SelectionLength = 0;
                                DDSupport.SelectionLength = 0;
                                DDMonth.SelectionLength = 0;
                                var confirmExportResult = MessageBox.Show("Data is now in database. Do you wish to get it in Excel format for your checking?", "Confirm",
                                                MessageBoxButtons.YesNo);
                                if (confirmExportResult == DialogResult.Yes)
                                {
                                    ExcelExport();
                                }
                                else
                                {
                                    lblmsg1.ForeColor = System.Drawing.Color.DarkGreen;
                                    lblmsg1.Text = "You can check the data through another Menu Option.";
                                }
                            }
                            else
                            {
                                lblmsg.Text = "";
                                lblmsg.Refresh();
                                InsertException("Wrong file format");
                                lblmsg1.Text = "                SmartRead data extraction failed." +
                                               "\n Possible reasons:Wrong document,Column mismatch or empty rows" +
                                               "\n                    Please check the source file.";
                                lblmsg1.ForeColor = System.Drawing.Color.Red;
                                btnCancel.Enabled = true;
                            }

                            DDInsurance.SelectionLength = 0;
                            DDLocation.SelectionLength = 0;
                            DDsales.SelectionLength = 0;
                            DDService.SelectionLength = 0;
                            DDSupport.SelectionLength = 0;
                            DDMonth.SelectionLength = 0;
                            //txtfile.Text = "Select pdf file";
                            //btnBrowse.Enabled = true;
                            btnCancel.Enabled = true;
                            //linkLabel2.Enabled = true;
                            //DDInsurance.SelectedValue = 0;
                            //DDLocation.SelectedValue = 0;
                            //DDsales.SelectedValue = 0;
                            //DDService.SelectedValue = 0;
                            //DDSupport.SelectedValue = 0;
                        }
                        else
                        {
                            lblmsg.Text = "";
                            lblmsg.Refresh();
                            MessageBox.Show("Please check the PDF. The pdf may contain image mask.", "Warning!");
                            btnCancel.Enabled = true; btnBrowse.Enabled = true;
                            DDInsurance.SelectionLength = 0;
                            DDLocation.SelectionLength = 0;
                            DDsales.SelectionLength = 0;
                            DDService.SelectionLength = 0;
                            DDSupport.SelectionLength = 0;
                            DDMonth.SelectionLength = 0;
                        }
                    }
                    else
                    {
                        lblmsg.Text = "";
                        lblmsg.Refresh();
                        MessageBox.Show("Given file's BDS is already exists in the database.", "Warning!");
                        btnCancel.Enabled = true; btnBrowse.Enabled = true;
                        DDInsurance.SelectionLength = 0;
                        DDLocation.SelectionLength = 0;
                        DDsales.SelectionLength = 0;
                        DDService.SelectionLength = 0;
                        DDSupport.SelectionLength = 0;
                        DDMonth.SelectionLength = 0;
                    }
                }
                else
                {
                    btnCancel.Enabled = true; 
                }
            }
            catch (Exception ex)
            {
                lblmsg.Text = "";
                lblmsg.Refresh(); 
                InsertException(ex.Message);
                lblmsg1.Text = "                SmartRead data extraction failed." +
                               "\n Possible reasons:Wrong document,Column mismatch or empty rows" +
                               "\n                    Please check the source file.";
                lblmsg1.ForeColor = System.Drawing.Color.Red;
                btnCancel.Enabled = true;
                DDInsurance.SelectionLength = 0;
                DDLocation.SelectionLength = 0;
                DDsales.SelectionLength = 0;
                DDService.SelectionLength = 0;
                DDSupport.SelectionLength = 0;
                DDMonth.SelectionLength = 0;
            }
        }

        public void BindDDInsurance()
        {
            DataRow dr;
            //string com = "select Code,InsurerCode + ','+ UPPER(LEFT(Description, 1)) + LOWER(RIGHT(Description, LEN(Description) - 1)) as Description from tblBRInsurancelkup where GroupBy = 'UN' and Code != '' order by Description asc";
            //string com = "select Code,InsurerCode + ',' + Code +' '+ UPPER(LEFT(Description, 1)) + LOWER(RIGHT(Description, LEN(Description) - 1)) as Description from tblBRInsurancelkup where GroupBy = 'UN' and Code != '' order by Description asc";
            string com = "select Code,Code +' '+ InsurerCode + ',' + UPPER(LEFT(Description, 1)) + LOWER(RIGHT(Description, LEN(Description) - 1)) as Description from tblBRInsurancelkup where GroupBy = 'NE' order by Description asc";
            SqlDataAdapter adpt = new SqlDataAdapter(com, strconn);
            DataTable dt = new DataTable();
            adpt.Fill(dt);
            dr = dt.NewRow();
            dr.ItemArray = new object[] { 0, "" };
            dt.Rows.InsertAt(dr, 0);

            DDInsurance.ValueMember = "Code";
            DDInsurance.DisplayMember = "Description";
            DDInsurance.DataSource = dt;
        }

        private void kryptonButton2_Click(object sender, EventArgs e)
        {
            Home obj = new Home();
            obj.Show();
            this.Close();
        }

        private void btnLogout_Click(object sender, EventArgs e)
        {
            Login obj = new Login();
            obj.Show();
            this.Close();
        }

        private void btnEdit_Click(object sender, EventArgs e)
        {
            string promptValue = ShowDialog("Batch");
            if (promptValue != "")
            {
                Fileinfo.Insurer = "NIAC,New India Assurance Company Ltd.";
                Fileinfo.BatchId = promptValue.Substring(0, promptValue.IndexOf(","));
                Fileinfo.Filename = promptValue.Substring(promptValue.IndexOf(",") + 1);
                EditForm obj = new EditForm();
                obj.Show();
                obj.WindowState = FormWindowState.Normal;
            }
        }
        public string ShowDialog(string caption)
        {
            var promptValue = "";
            Form prompt = new Form()
            {
                Width = 600,
                Height = 200,
                FormBorderStyle = FormBorderStyle.FixedDialog,
                BackColor = System.Drawing.Color.White,
                Text = caption,
                StartPosition = FormStartPosition.CenterScreen,
                MinimizeBox = false,
                MaximizeBox = false
            };
            Label textLabel = new Label() { Left = 50, Top = 30, Text = "Batch" };
            ComboBox CB = new ComboBox() { Left = 100, Top = 30, Width = 450 };
            Label lblErmsg = new Label() { Left = 50, Top = 50, Width = 300 };
            Button confirmation = new Button() { Text = "Ok", Left = 220, Width = 80, Top = 100, DialogResult = DialogResult.OK, Enabled = false };
            Button confirmation1 = new Button() { Text = "Cancel", Left = 320, Width = 80, Top = 100, DialogResult = DialogResult.Cancel };
            textLabel.Font = new Font("Verdana", 11);
            CB.Font = new Font("Verdana", 9);
            lblErmsg.Font = new Font("Verdana", 9);
            lblErmsg.ForeColor = System.Drawing.Color.Red;
            confirmation.Font = new Font("Verdana", 9);
            confirmation1.Font = new Font("Verdana", 9);
            DataRow dr;
            string com = "select distinct(Inv_No) as No,Inv_No+','+[Filename] as Name from BDSMaster where InsurerCode = 'NIAC' and ReportCode = '" + Fileinfo.ReportId+"'";
            SqlDataAdapter adpt = new SqlDataAdapter(com, strconn);
            DataTable dt = new DataTable();
            adpt.Fill(dt);
            dr = dt.NewRow();
            dr.ItemArray = new object[] { 0, "" };
            dt.Rows.InsertAt(dr, 0);

            CB.ValueMember = "No";
            CB.DisplayMember = "Name";
            CB.DataSource = dt;
            CB.AutoCompleteMode = AutoCompleteMode.Suggest;
            CB.AutoCompleteSource = AutoCompleteSource.ListItems;

            //confirmation.Click += (sender, e) => { prompt.Close(); };
            confirmation1.Click += (sender, e) => { prompt.Close(); };
            prompt.Controls.Add(CB);
            prompt.Controls.Add(confirmation);
            prompt.Controls.Add(confirmation1);
            prompt.Controls.Add(textLabel);
            prompt.AcceptButton = confirmation;

            CB.SelectedIndexChanged += (sender, e) =>
            {
                if (CB.SelectedValue.ToString() != "0")
                {
                    confirmation.Enabled = true;
                }
                else
                {
                    confirmation.Enabled = false;
                }
            };
            //if(confirmation.DialogResult == DialogResult.OK && CB.SelectedValue.ToString() != "0")
            if (prompt.ShowDialog() == DialogResult.OK && CB.SelectedValue.ToString() != "0")
            {
                // promptValue = CB.SelectedValue.ToString();
                promptValue = CB.Text;
            }
            else if (prompt.ShowDialog() == DialogResult.Cancel)
            {
                promptValue = "";
                prompt.Close();
            }
            return promptValue;
        }

        private void linkLabel2_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            Fileinfo.Insurer = "NIAC,New India Assurance Company Ltd.";
            Fileinfo.BatchId = BatchID;
            EditForm obj = new EditForm();
            obj.Show();
            obj.WindowState = FormWindowState.Normal;
        }

        public void BindDDSales()
        {
            DataRow dr;
            string com = "select Code,Description from tblSalesByLkup order by Description asc";
            SqlDataAdapter adpt = new SqlDataAdapter(com, strconn);
            DataTable dt = new DataTable();
            adpt.Fill(dt);
            dr = dt.NewRow();
            dr.ItemArray = new object[] { 0, "" };
            dt.Rows.InsertAt(dr, 0);

            DDsales.ValueMember = "Code";
            DDsales.DisplayMember = "Description";
            DDsales.DataSource = dt;
        }
        public void BindDDService()
        {
            DataRow dr;
            string com = "select Code,Description from tblServicedByLkup order by Description asc";
            SqlDataAdapter adpt = new SqlDataAdapter(com, strconn);
            DataTable dt = new DataTable();
            adpt.Fill(dt);
            dr = dt.NewRow();
            dr.ItemArray = new object[] { 0, "" };
            dt.Rows.InsertAt(dr, 0);

            DDService.ValueMember = "Code";
            DDService.DisplayMember = "Description";
            DDService.DataSource = dt;
        }

        private void kryptonButton3_Click(object sender, EventArgs e)
        {

            if (Fileinfo.ReportId == "NIP1") //New India Assurance Company Ltd.
            {
                NewIndiaTemplate obj = new NewIndiaTemplate();
                obj.Show();
            }
            if (Fileinfo.ReportId == "NIP2") //New India Assurance Company Ltd.
            {
                NewIndiaPDFSample2 obj = new NewIndiaPDFSample2();
                obj.Show();
            }
            if (Fileinfo.ReportId == "NIP3") //New India Assurance Company Ltd.
            {
                NewIndiaPDFSample3 obj = new NewIndiaPDFSample3();
                obj.Show();
            }
        }

        public void BindDDLocation()
        {
            DataRow dr;
            string com = "select CodeValue,ShortDescription from tblLookup where Codeid = 'OL' order by ShortDescription asc";
            SqlDataAdapter adpt = new SqlDataAdapter(com, strconn);
            DataTable dt = new DataTable();
            adpt.Fill(dt);
            dr = dt.NewRow();
            dr.ItemArray = new object[] { 0, "" };
            dt.Rows.InsertAt(dr, 0);

            DDLocation.ValueMember = "CodeValue";
            DDLocation.DisplayMember = "ShortDescription";
            DDLocation.DataSource = dt;
        }
        public void BindDDSupport()
        {
            DataRow dr;
            string com = "select CodeValue,ShortDescription from tblLookup where Codeid = 'BST' order by ShortDescription asc";
            SqlDataAdapter adpt = new SqlDataAdapter(com, strconn);
            DataTable dt = new DataTable();
            adpt.Fill(dt);
            dr = dt.NewRow();
            dr.ItemArray = new object[] { 0, "" };
            dt.Rows.InsertAt(dr, 0);

            DDSupport.ValueMember = "CodeValue";
            DDSupport.DisplayMember = "ShortDescription";
            DDSupport.DataSource = dt;
        }
        private void btnBrowse_Click(object sender, EventArgs e)
        {
            lblSuccMsg.Text = "";

            if (DDInsurance.SelectedValue.ToString() != "0" && DDLocation.SelectedValue.ToString() != "0" && DDsales.SelectedValue.ToString() != "0" && DDService.SelectedValue.ToString() != "0" && (DDMonth.SelectedIndex != -1 && DDMonth.SelectedIndex != 0))
            {
                btnConvert.Enabled = true;
                btnCancel.Enabled = true;
                DialogResult dr = openFileDialog1.ShowDialog();

                if (dr.Equals(DialogResult.OK))
                {
                    //lblFile1.Text = openFileDialog1.FileName;

                    btnBrowse.Enabled = true;
                    btnCancel.Enabled = true;
                    btnConvert.Enabled = true;
                    Dir = System.IO.Path.GetDirectoryName(openFileDialog1.FileName);
                    Filename = System.IO.Path.GetFileNameWithoutExtension(openFileDialog1.FileName);
                    Filepath = openFileDialog1.FileName;
                    Filewithext = System.IO.Path.GetFileName(openFileDialog1.FileName);

                    txtfile.Text = Filewithext;
                    txtfile.ForeColor = System.Drawing.Color.Black;
                    Fileinfo.Filename = Filewithext;
                    Fileinfo.TName = "NewIndiaTransaction";
                    btnBrowse.Enabled = false;
                }
                else
                {
                    btnBrowse.Enabled = true;
                    btnCancel.Enabled = true;
                    btnConvert.Enabled = false;
                }
            }
            else
            {
                if (DDInsurance.SelectedValue.ToString() == "0")
                {
                    MessageBox.Show("Please select Insurance");
                }
                else if (DDLocation.SelectedValue.ToString() == "0")
                {
                    MessageBox.Show("Please select Office Location");
                }
                else if (DDsales.SelectedValue.ToString() == "0")
                {
                    MessageBox.Show("Please select Sales Generated By");
                }
                else if (DDService.SelectedValue.ToString() == "0")
                {
                    MessageBox.Show("Please select Serviced By");
                }
                else if (DDMonth.SelectedIndex == -1 || DDMonth.SelectedIndex == 0)
                {
                    MessageBox.Show("Please select Report Month");
                }
                btnConvert.Enabled = false;
                btnCancel.Enabled = false;
                btnBrowse.Enabled = true;
            }
        }
        private static NPOI.SS.UserModel.IWorkbook OpenWorkBook(string workBookName)
        {
            using (FileStream file = new FileStream(workBookName, FileMode.Open, FileAccess.Read))
            {
                //return new XSSFWorkbook(file);
                return new HSSFWorkbook(file);
            }
        }
        private static void SaveWorkBook(IWorkbook workbook, string workBookName)
        {
            string newFileName = System.IO.Path.ChangeExtension(workBookName, "new.xls");
            using (FileStream file = new FileStream(newFileName, FileMode.Create, FileAccess.Write))
            {
                workbook.Write(file);
            }

            string backupFileName = System.IO.Path.ChangeExtension(workBookName, "bak.xls");
            File.Replace(newFileName, workBookName, backupFileName);
        }
        public static void DeleteRows(ISheet sheet)
        {
            var num = sheet.LastRowNum;
            int termsList1 = -1, termsList2 = -1;
            for (int rowIndex = sheet.LastRowNum; rowIndex >= 0; rowIndex--)
            {
                IRow row = sheet.GetRow(rowIndex);
                if (row == null) continue;
                ICell cell = row.GetCell(0);
                if (cell != null && cell.StringCellValue.Contains("Grand Total"))
                {
                    if (rowIndex != sheet.LastRowNum)
                    {
                        sheet.ShiftRows(row.RowNum + 1, sheet.LastRowNum, -1);
                    }
                }
                ICell cell2 = row.GetCell(4);
                if (cell2 != null && cell2.CellType.ToString() != "Numeric" && cell2.StringCellValue.Contains("CGST Input"))
                {
                    termsList1 = rowIndex;
                }
                termsList2 = sheet.LastRowNum - 1;

                if (termsList1 != -1 && termsList2 != -1)
                {

                    for (int list = termsList2; list >= termsList1; list--)
                    {
                        IRow row1 = sheet.GetRow(list);
                        sheet.ShiftRows(row1.RowNum + 1, sheet.LastRowNum, -1);
                    }
                    termsList2 = -1;
                    termsList1 = -1;
                }
            }
        }
        public void GetBachid_NoRecord()
        {
            try
            {
                SQLProcs sql = new SQLProcs();
                DataSet ds = new DataSet();

                ds = sql.SQLExecuteDataset("SP_NewIndiaPDFTransaction",
                     new SqlParameter { ParameterName = "@Imode", Value = 4 },
                     new SqlParameter { ParameterName = "@TranID", Value = TranID }
                     );

                if (ds != null && ds.Tables.Count > 0 && ds.Tables[0].Rows.Count > 0)
                {
                    BatchID = ds.Tables[0].Rows[0]["BatchID"].ToString();
                }
                else
                {
                    BatchID = "";
                }

                DataSet dsR = new DataSet();
                dsR = sql.SQLExecuteDataset("SP_NewIndiaPDFTransaction",
                     new SqlParameter { ParameterName = "@Imode", Value = 5 },
                     new SqlParameter { ParameterName = "@BatchID", Value = BatchID },
                     new SqlParameter { ParameterName = "@Filename", Value = Fileinfo.Filename },
                     new SqlParameter { ParameterName = "@version", Value = LoginInfo.version },
                     new SqlParameter { ParameterName = "@UserId", Value = LoginInfo.UserID }
                     );
                if (dsR != null && dsR.Tables.Count > 0 && dsR.Tables[0].Rows.Count > 0)
                {
                    NoRecord = dsR.Tables[0].Rows[0]["NoRecord"].ToString();
                }
                else
                {
                    NoRecord = "";
                }
            }
            catch (Exception ex)
            {

                InsertException(ex.Message);
            }
        }
        public void InsertException(string exception)
        {
            SQLProcs sql = new SQLProcs();
            sql.ExecuteSQLNonQuery("SP_Login",
                     new SqlParameter { ParameterName = "@Imode", Value = 10 },
                     new SqlParameter { ParameterName = "@Exception", Value = exception },
                     new SqlParameter { ParameterName = "@InsurerCode", Value = Fileinfo.InsurerCode },
                     new SqlParameter { ParameterName = "@ReportCode", Value = Fileinfo.ReportId },
                     new SqlParameter { ParameterName = "@UserId", Value = LoginInfo.UserID }
                     );
        }
        public void ExcelExport()
        {
            try
            {
                string pathUser = Environment.GetFolderPath(Environment.SpecialFolder.UserProfile);
                string pathDownload = Path.Combine(pathUser, "Downloads\\");

                SQLProcs sql = new SQLProcs();
                DataSet ResultsTable = new DataSet();

                ResultsTable = sql.SQLExecuteDataset("SP_NewIndiaPDFTransaction",
               new SqlParameter { ParameterName = "@Imode", Value = 3 },
               new SqlParameter { ParameterName = "@TranID", Value = TranID },
               new SqlParameter { ParameterName = "@DocName", Value = "New India Assurance Company Ltd." }
               );

                string date = DateTime.Now.ToString();
                date = date.Replace("/", "_").Replace(":", "").Replace(" ", "").Replace("AM", "").Replace("PM", "");

                string FileName = Fileinfo.ReportId+"_" + date + ".xlsx";

                using (ClosedXML.Excel.XLWorkbook wb = new ClosedXML.Excel.XLWorkbook())
                {
                    for (int i = 0; i < ResultsTable.Tables.Count; i++)
                    {
                        wb.Worksheets.Add(ResultsTable.Tables[i], ResultsTable.Tables[i].TableName);
                    }
                    wb.Style.Alignment.Horizontal = ClosedXML.Excel.XLAlignmentHorizontalValues.Center;
                    wb.Style.Font.Bold = true;
                    //if (saveDlg.ShowDialog() == DialogResult.OK)
                    //{
                    //string path = saveDlg.FileName;
                    string path = pathDownload + "\\" + FileName;
                    wb.SaveAs(path);

                    sql.SQLExecuteDataset("SP_NewIndiaPDFTransaction",
                    new SqlParameter { ParameterName = "@Imode", Value = 6 },
                    new SqlParameter { ParameterName = "@BatchID", Value = BatchID },
                    new SqlParameter { ParameterName = "@version", Value = LoginInfo.version },
                    new SqlParameter { ParameterName = "@UserId", Value = LoginInfo.UserID }
                    );

                    lblmsg1.ForeColor = System.Drawing.Color.Green;
                    lblmsg1.Text = "SmartRead data downloaded as XLSX file for your verification.\n     (File Name:" + FileName + ")";
                }
                //return result;
            }
            catch (Exception ex)
            {
                //return ex.Message;
                lblSuccMsg.Text = "";
                InsertException(ex.Message);
                lblmsg1.Text = "Data export failed.";
                lblmsg1.ForeColor = System.Drawing.Color.Red;
                linkLabel2.Text = "";
                linkLabel2.Enabled = false;
            }
        }

        public static void InsertTransaction1(Microsoft.Office.Interop.Excel.Workbook WB, string TranID, string RDate, string Insurance, string Salesby, string Serviceby, string location, string Support, string Rmonth)
        {
            SQLProcs sql = new SQLProcs();
            Microsoft.Office.Interop.Excel.Worksheet wks = (Microsoft.Office.Interop.Excel.Worksheet)WB.Worksheets[2];
            Microsoft.Office.Interop.Excel.Range lastCell = wks.Cells.SpecialCells(Microsoft.Office.Interop.Excel.XlCellType.xlCellTypeLastCell, Type.Missing);
            int lastrow = lastCell.Row; string PolicyNo = ""; int J = 0;
            string Checknull = ((Microsoft.Office.Interop.Excel.Range)wks.Cells[1, 13]).Value;
            string CheckFormat = ((Microsoft.Office.Interop.Excel.Range)wks.Cells[1, 3]).Value;
            FormatResult = "Not OK";
            if (Checknull != null && Checknull.Contains("Insured Name"))
            {
                J = 1;
            }
            if (CheckFormat != null)
            {
                CheckFormat = CheckFormat.Replace("\n", "").TrimStart();
                if (CheckFormat.Contains("Office Code") || CheckFormat.Contains("OfficeCode"))
                {
                    FormatResult = "OK";
                }
                else
                {
                    FormatResult = "Not OK";
                }
            }
            if (FormatResult == "OK")
            {
                for (int i = 2; i <= lastrow; i++)
                {
                    var Terrorism = ""; var Policy_Endorsement = ""; var InsuredType = ""; var Policy_Type = ""; var New_Renewal = ""; var Revenue_Amt = "";

                    string InsuredName = ((Microsoft.Office.Interop.Excel.Range)wks.Cells[i, 14 - J]).Value;
                    InsuredType = ((Microsoft.Office.Interop.Excel.Range)wks.Cells[i, 16 - J]).Value;
                    if (InsuredType != null && InsuredType != "" && InsuredType != " ")
                    {
                        InsuredName = InsuredName.Replace("\n", " ").TrimStart();
                        string Pno = Convert.ToString(((Microsoft.Office.Interop.Excel.Range)wks.Cells[i, 4]).Value);
                        //var Endorsementno = Convert.ToString(((Microsoft.Office.Interop.Excel.Range)wks.Cells[i, 15]).Value);
                        if (Pno != null && Pno != "" && Pno != " ")
                        {
                            PolicyNo = Pno.Replace("\n", "").TrimStart();
                        }
                        var Endo_Effective_Date = ((Microsoft.Office.Interop.Excel.Range)wks.Cells[i, 18 - J]).Value;
                        var Effective_Date = ((Microsoft.Office.Interop.Excel.Range)wks.Cells[i, 18 - J]).Value;
                        var END_Date = ((Microsoft.Office.Interop.Excel.Range)wks.Cells[i, 19 - J]).Value;
                        int ENdolen = Convert.ToString(Endo_Effective_Date).Length;
                        if (ENdolen > 11)
                        {
                            Endo_Effective_Date = Endo_Effective_Date.ToString("dd/MM/yyyy");
                        }
                        int Efflen = Convert.ToString(Effective_Date).Length;
                        if (Efflen > 11)
                        {
                            Effective_Date = Effective_Date.ToString("dd/MM/yyyy");
                        }
                        int ENDlen = Convert.ToString(END_Date).Length;
                        if (ENDlen > 11)
                        {
                            END_Date = END_Date.ToString("dd/MM/yyyy");
                        }
                        var Premium_Amt = Convert.ToString(((Microsoft.Office.Interop.Excel.Range)wks.Cells[i, 21 - J]).Value).Replace(",", "").Replace("(", "").Replace(")", "").Replace(".", "").TrimStart();
                        //Terrorism = Convert.ToString(((Microsoft.Office.Interop.Excel.Range)wks.Cells[i, 23]).Value).Replace(",", "").Replace("(", "").Replace(")", "").TrimStart();
                        Revenue_Amt = Convert.ToString(((Microsoft.Office.Interop.Excel.Range)wks.Cells[i, 22 - J]).Value).Replace(",", "").Replace("(", "").Replace(")", "").Replace(".", "").TrimStart();
                        var Revenue_Pcnt = Convert.ToString(((Microsoft.Office.Interop.Excel.Range)wks.Cells[i, 20 - J]).Text);
                        Revenue_Pcnt = Revenue_Pcnt.Replace("\n", "").Replace("%", "").Replace(",", "").TrimStart();
                        PolicyNo = Regex.Replace(PolicyNo, "[^a-zA-Z0-9_.]+", "", RegexOptions.Compiled);
                        string endno = Convert.ToString(((Microsoft.Office.Interop.Excel.Range)wks.Cells[i, 8]).Value).Replace("\n", "").TrimStart();
                        if (endno != null && endno != "" && endno != " " && endno != ":")
                        {
                            Policy_Endorsement = "Endorsement";
                        }
                        else
                        {
                            Policy_Endorsement = "Policy";
                        }
                        Policy_Type = Convert.ToString(((Microsoft.Office.Interop.Excel.Range)wks.Cells[i, 12]).Value);
                        //if (InsuredName.ToUpper().Contains("LIMITED") || InsuredName.ToUpper().Contains("LTD"))
                        if (InsuredType == "Organizational")
                        {
                            InsuredType = "Corporate";
                        }
                        else
                        {
                            InsuredType = "Retail";
                        }

                        if (Premium_Amt == "" || Premium_Amt == " " || Premium_Amt == null)
                        {
                            Premium_Amt = 0;
                        }
                        if (Revenue_Amt == "" || Revenue_Amt == " " || Revenue_Amt == null)
                        {
                            Revenue_Amt = "0";
                        }
                        if (Revenue_Pcnt == "" || Revenue_Pcnt == " " || Revenue_Pcnt == null)
                        {
                            Revenue_Pcnt = "0";
                        }
                        if (Terrorism == "" || Terrorism == " " || Terrorism == null)
                        {
                            Terrorism = "0";
                        }
                        sql.ExecuteSQLNonQuery("SP_NewIndiaPDFTransaction",
                                   new SqlParameter { ParameterName = "@Imode", Value = 1 },
                                   new SqlParameter { ParameterName = "@RDate", Value = RDate },
                                   new SqlParameter { ParameterName = "@Rmonth", Value = Rmonth },
                                   new SqlParameter { ParameterName = "@ClientName", Value = InsuredName },
                                   new SqlParameter { ParameterName = "@Itype", Value = InsuredType },
                                   new SqlParameter { ParameterName = "@Policy_No", Value = PolicyNo },
                                   new SqlParameter { ParameterName = "@Revenue_Pct", Value = Revenue_Pcnt },
                                   new SqlParameter { ParameterName = "@Policy_Type", Value = Policy_Type },
                                   new SqlParameter { ParameterName = "@Premium_Amt", Value = Premium_Amt },
                                   new SqlParameter { ParameterName = "@Endo_Effective_Date", Value = Endo_Effective_Date },
                                   new SqlParameter { ParameterName = "@Effective_Date", Value = Effective_Date },
                                   new SqlParameter { ParameterName = "@END_Date", Value = END_Date },
                                   new SqlParameter { ParameterName = "@TranID", Value = TranID },
                                   new SqlParameter { ParameterName = "@Revenue_Amt", Value = Revenue_Amt },
                                   new SqlParameter { ParameterName = "@Terrorism", Value = Terrorism },
                                   new SqlParameter { ParameterName = "@Insurance", Value = Insurance },
                                   new SqlParameter { ParameterName = "@Salesby", Value = Salesby },
                                   new SqlParameter { ParameterName = "@Serviceby", Value = Serviceby },
                                   new SqlParameter { ParameterName = "@location", Value = location },
                                   new SqlParameter { ParameterName = "@Support", Value = Support },
                                   new SqlParameter { ParameterName = "@Policy_Endorsement", Value = Policy_Endorsement },
                                   new SqlParameter { ParameterName = "@RFormat", Value = "F1" },
                                   new SqlParameter { ParameterName = "@InvNo", Value = "NIP1" },
                                   new SqlParameter { ParameterName = "@ReportId", Value = "NIP1" },
                                   new SqlParameter { ParameterName = "@DocName", Value = "New India Assurance Company Ltd." }
                                   );
                    }
                }
            }
        }

        public static void InsertTransaction2(Microsoft.Office.Interop.Excel.Workbook WB, string TranID, string RDate, string Insurance, string Salesby, string Serviceby, string location, string Support, string Rmonth)
        {
            //New India Assurance Company Limited.
            SQLProcs sql = new SQLProcs();
            Microsoft.Office.Interop.Excel.Worksheet wks = (Microsoft.Office.Interop.Excel.Worksheet)WB.Worksheets[2];
            Microsoft.Office.Interop.Excel.Range lastCell = wks.Cells.SpecialCells(Microsoft.Office.Interop.Excel.XlCellType.xlCellTypeLastCell, Type.Missing);
            int lastrow = lastCell.Row;
            string CheckFormat = ((Microsoft.Office.Interop.Excel.Range)wks.Cells[1, 3]).Value;
            FormatResult = "Not OK";
            if (CheckFormat != null)
            {
                CheckFormat = CheckFormat.Replace("\n", "").TrimStart();
                if (CheckFormat.Contains("Dept"))
                {
                    FormatResult = "OK";
                }
                else
                {
                    FormatResult = "Not OK";
                }
            }
            if (FormatResult == "OK")
            {
                for (int i = 2; i <= lastrow; i++)
                {
                    var Terrorism = ""; var Policy_Endorsement = ""; var InsuredType = ""; var Policy_Type = ""; var New_Renewal = ""; var Revenue_Amt = "";
                    string InsuredName = ((Microsoft.Office.Interop.Excel.Range)wks.Cells[i, 9]).Value;
                    if (InsuredName != null && InsuredName != "" && InsuredName != " ")
                    {
                        InsuredName = InsuredName.Replace("\n", " ").TrimStart();
                        InsuredType = ((Microsoft.Office.Interop.Excel.Range)wks.Cells[i, 12]).Value;
                        string PolicyNo = Convert.ToString(((Microsoft.Office.Interop.Excel.Range)wks.Cells[i, 4]).Value).Replace("\n", "").TrimStart();
                        //var Endorsementno = Convert.ToString(((Microsoft.Office.Interop.Excel.Range)wks.Cells[i, 15]).Value);

                        var Endo_Effective_Date = ((Microsoft.Office.Interop.Excel.Range)wks.Cells[i, 15]).Value;
                        var Effective_Date = ((Microsoft.Office.Interop.Excel.Range)wks.Cells[i, 15]).Value;
                        var END_Date = ((Microsoft.Office.Interop.Excel.Range)wks.Cells[i, 17]).Value;
                        int ENdolen = Convert.ToString(Endo_Effective_Date).Length;
                        if (ENdolen > 11)
                        {
                            Endo_Effective_Date = Endo_Effective_Date.ToString("dd/MM/yyyy");
                        }
                        int Efflen = Convert.ToString(Effective_Date).Length;
                        if (Efflen > 11)
                        {
                            Effective_Date = Effective_Date.ToString("dd/MM/yyyy");
                        }
                        int ENDlen = Convert.ToString(END_Date).Length;
                        if (ENDlen > 11)
                        {
                            END_Date = END_Date.ToString("dd/MM/yyyy");
                        }
                        var Premium_Amt = Convert.ToString(((Microsoft.Office.Interop.Excel.Range)wks.Cells[i, 20]).Value);
                        Premium_Amt = Premium_Amt.Replace(",", "").Replace(".", "").Replace("(", "").Replace(")", "").TrimStart();
                        //Terrorism = Convert.ToString(((Microsoft.Office.Interop.Excel.Range)wks.Cells[i, 23]).Value).Replace(",", "").Replace("(", "").Replace(")", "").TrimStart();
                        Revenue_Amt = Convert.ToString(((Microsoft.Office.Interop.Excel.Range)wks.Cells[i, 21]).Value);
                        Revenue_Amt = Revenue_Amt.Replace(",", "").Replace("(", "").Replace(")", "").Replace(".", "").TrimStart();
                        var Revenue_Pcnt = Convert.ToString(((Microsoft.Office.Interop.Excel.Range)wks.Cells[i, 18]).Text);
                        Revenue_Pcnt = Revenue_Pcnt.Replace("\n", "").Replace("%", "").Replace(",", "").TrimStart();
                        if (PolicyNo != null && PolicyNo != "" && PolicyNo != " ")
                        {
                            PolicyNo = Regex.Replace(PolicyNo, "[^a-zA-Z0-9_.]+", "", RegexOptions.Compiled);
                        }
                        Policy_Type = Convert.ToString(((Microsoft.Office.Interop.Excel.Range)wks.Cells[i, 3]).Value);
                        if (InsuredType == "Organizational")
                        {
                            InsuredType = "Corporate";
                        }
                        else
                        {
                            InsuredType = "Retail";
                        }
                        string endno = Convert.ToString(((Microsoft.Office.Interop.Excel.Range)wks.Cells[i, 7]).Value).Replace("\n", "").TrimStart();
                        if (endno != null && endno != "" && endno != " " && endno != ":")
                        {
                            Policy_Endorsement = "Endorsement";
                            PolicyNo = PolicyNo + " " + Regex.Replace(endno, "[^a-zA-Z0-9_.]+", "", RegexOptions.Compiled);
                        }
                        else
                        {
                            Policy_Endorsement = "Policy";
                        }
                        if (Premium_Amt == "" || Premium_Amt == " " || Premium_Amt == null)
                        {
                            Premium_Amt = 0;
                        }
                        if (Revenue_Amt == "" || Revenue_Amt == " " || Revenue_Amt == null)
                        {
                            Revenue_Amt = "0";
                        }
                        if (Revenue_Pcnt == "" || Revenue_Pcnt == " " || Revenue_Pcnt == null)
                        {
                            Revenue_Pcnt = "0";
                        }
                        if (Terrorism == "" || Terrorism == " " || Terrorism == null)
                        {
                            Terrorism = "0";
                        }
                        sql.ExecuteSQLNonQuery("SP_NewIndiaPDFTransaction",
                                   new SqlParameter { ParameterName = "@Imode", Value = 1 },
                                   new SqlParameter { ParameterName = "@RDate", Value = RDate },
                                   new SqlParameter { ParameterName = "@Rmonth", Value = Rmonth },
                                   new SqlParameter { ParameterName = "@ClientName", Value = InsuredName },
                                   new SqlParameter { ParameterName = "@Itype", Value = InsuredType },
                                   new SqlParameter { ParameterName = "@Policy_No", Value = PolicyNo },
                                   new SqlParameter { ParameterName = "@Revenue_Pct", Value = Revenue_Pcnt },
                                   new SqlParameter { ParameterName = "@Policy_Type", Value = Policy_Type },
                                   new SqlParameter { ParameterName = "@Premium_Amt", Value = Premium_Amt },
                                   new SqlParameter { ParameterName = "@Endo_Effective_Date", Value = Endo_Effective_Date },
                                   new SqlParameter { ParameterName = "@Effective_Date", Value = Effective_Date },
                                   new SqlParameter { ParameterName = "@END_Date", Value = END_Date },
                                   new SqlParameter { ParameterName = "@TranID", Value = TranID },
                                   new SqlParameter { ParameterName = "@Revenue_Amt", Value = Revenue_Amt },
                                   new SqlParameter { ParameterName = "@Terrorism", Value = Terrorism },
                                   new SqlParameter { ParameterName = "@Insurance", Value = Insurance },
                                   new SqlParameter { ParameterName = "@Salesby", Value = Salesby },
                                   new SqlParameter { ParameterName = "@Serviceby", Value = Serviceby },
                                   new SqlParameter { ParameterName = "@location", Value = location },
                                   new SqlParameter { ParameterName = "@Support", Value = Support },
                                   new SqlParameter { ParameterName = "@Policy_Endorsement", Value = Policy_Endorsement },
                                   new SqlParameter { ParameterName = "@RFormat", Value = "F2" },
                                   new SqlParameter { ParameterName = "@InvNo", Value = "NIP2" },
                                   new SqlParameter { ParameterName = "@ReportId", Value = "NIP2" },
                                   new SqlParameter { ParameterName = "@DocName", Value = "New India Assurance Company Ltd." }
                                   );
                    }
                }
            }
        }
        public static void InsertTransaction3(Microsoft.Office.Interop.Excel.Workbook WB, string TranID, string RDate, string Insurance, string Salesby, string Serviceby, string location, string Support, string Rmonth)
        {
            //New India Assurance Company Limited.
            SQLProcs sql = new SQLProcs();
            Microsoft.Office.Interop.Excel.Worksheet wks = (Microsoft.Office.Interop.Excel.Worksheet)WB.Worksheets[2];
            Microsoft.Office.Interop.Excel.Range lastCell = wks.Cells.SpecialCells(Microsoft.Office.Interop.Excel.XlCellType.xlCellTypeLastCell, Type.Missing);
            int lastrow = lastCell.Row; string Policy_Type = ""; var PType = "";

            string CheckFormat = ((Microsoft.Office.Interop.Excel.Range)wks.Cells[1, 1]).Value;
            FormatResult = "Not OK";
            if (CheckFormat != null)
            {
                CheckFormat = CheckFormat.Replace("\n", "").TrimStart();
                if (CheckFormat.Contains("Department"))
                {
                    FormatResult = "OK";
                }
                else
                {
                    FormatResult = "Not OK";
                }
            }
            if (FormatResult == "OK")
            {
                for (int i = 2; i <= lastrow; i++)
                {
                    var Terrorism = ""; var Policy_Endorsement = ""; var InsuredType = ""; var Revenue_Pcnt = ""; var Revenue_Amt = ""; var PolicyNo = "";
                    string InsuredName = ((Microsoft.Office.Interop.Excel.Range)wks.Cells[i, 6]).Value;
                    if (InsuredName != null && InsuredName != "" && InsuredName != " " && InsuredName != "Insured Name")
                    {
                        InsuredName = InsuredName.Replace("\n", " ").TrimStart();
                        InsuredType = ((Microsoft.Office.Interop.Excel.Range)wks.Cells[i, 7]).Value;
                        var PNo = Convert.ToString(((Microsoft.Office.Interop.Excel.Range)wks.Cells[i, 2]).Value);
                        if (PNo != null && PNo != "" && PNo != " ")
                        {
                            PType = PNo.Replace("\n", "").TrimStart();
                            PType = Regex.Replace(PType, "[^a-zA-Z0-9_.]+", "", RegexOptions.Compiled);
                        }
                        var endno = Convert.ToString(((Microsoft.Office.Interop.Excel.Range)wks.Cells[i, 5]).Value);
                        if (endno != null && endno != "" && endno != " " && endno != ":")
                        {
                            Policy_Endorsement = "Endorsement";
                            PolicyNo = PType + " " + Regex.Replace(endno, "[^a-zA-Z0-9_.]+", "", RegexOptions.Compiled);
                        }
                        else
                        {
                            Policy_Endorsement = "Policy";
                            PolicyNo = PType;
                        }
                        if (InsuredType == "Organizational")
                        {
                            InsuredType = "Corporate";
                        }
                        else
                        {
                            InsuredType = "Retail";
                        }
                        var Endo_Effective_Date = ((Microsoft.Office.Interop.Excel.Range)wks.Cells[i, 8]).Value;
                        var Effective_Date = ((Microsoft.Office.Interop.Excel.Range)wks.Cells[i, 8]).Value;
                        var END_Date = ((Microsoft.Office.Interop.Excel.Range)wks.Cells[i, 9]).Value;
                        int ENdolen = Convert.ToString(Endo_Effective_Date).Length;
                        if (ENdolen > 11)
                        {
                            Endo_Effective_Date = Endo_Effective_Date.ToString("dd/MM/yyyy");
                        }
                        int Efflen = Convert.ToString(Effective_Date).Length;
                        if (Efflen > 11)
                        {
                            Effective_Date = Effective_Date.ToString("dd/MM/yyyy");
                        }
                        int ENDlen = Convert.ToString(END_Date).Length;
                        if (ENDlen > 11)
                        {
                            END_Date = END_Date.ToString("dd/MM/yyyy");
                        }
                        var Premium_Amt = Convert.ToString(((Microsoft.Office.Interop.Excel.Range)wks.Cells[i, 10]).Value).Replace(",", "").Replace("(", "").Replace(")", "").Replace(".", "").TrimStart();
                        //Terrorism = Convert.ToString(((Microsoft.Office.Interop.Excel.Range)wks.Cells[i, 23]).Value).Replace(",", "").Replace("(", "").Replace(")", "").TrimStart();
                        Revenue_Amt = Convert.ToString(((Microsoft.Office.Interop.Excel.Range)wks.Cells[i, 11]).Value).Replace(",", "").Replace("(", "").Replace(")", "").Replace(".", "").TrimStart();
                        //var Revenue_Pcnt = Convert.ToString(((Microsoft.Office.Interop.Excel.Range)wks.Cells[i, 21]).Text);
                        //Revenue_Pcnt = Revenue_Pcnt.Replace("\n", "").Replace("%", "").Replace(",", "").TrimStart();

                        string dept = Convert.ToString(((Microsoft.Office.Interop.Excel.Range)wks.Cells[i, 1]).Value);
                        if (dept != null && dept != "" && dept != " ")
                        {
                            Policy_Type = Regex.Match(dept, @"\d+").Value;
                        }

                        if (Premium_Amt == "" || Premium_Amt == " " || Premium_Amt == null)
                        {
                            Premium_Amt = 0;
                        }
                        if (Revenue_Amt == "" || Revenue_Amt == " " || Revenue_Amt == null)
                        {
                            Revenue_Amt = "0";
                        }
                        if (Revenue_Pcnt == "" || Revenue_Pcnt == " " || Revenue_Pcnt == null)
                        {
                            Revenue_Pcnt = "0";
                        }
                        if (Terrorism == "" || Terrorism == " " || Terrorism == null)
                        {
                            Terrorism = "0";
                        }
                        sql.ExecuteSQLNonQuery("SP_NewIndiaPDFTransaction",
                                   new SqlParameter { ParameterName = "@Imode", Value = 1 },
                                   new SqlParameter { ParameterName = "@RDate", Value = RDate },
                                   new SqlParameter { ParameterName = "@Rmonth", Value = Rmonth },
                                   new SqlParameter { ParameterName = "@ClientName", Value = InsuredName },
                                   new SqlParameter { ParameterName = "@Itype", Value = InsuredType },
                                   new SqlParameter { ParameterName = "@Policy_No", Value = PolicyNo },
                                   new SqlParameter { ParameterName = "@Revenue_Pct", Value = Revenue_Pcnt },
                                   new SqlParameter { ParameterName = "@Policy_Type", Value = Policy_Type },
                                   new SqlParameter { ParameterName = "@Premium_Amt", Value = Premium_Amt },
                                   new SqlParameter { ParameterName = "@Endo_Effective_Date", Value = Endo_Effective_Date },
                                   new SqlParameter { ParameterName = "@Effective_Date", Value = Effective_Date },
                                   new SqlParameter { ParameterName = "@END_Date", Value = END_Date },
                                   new SqlParameter { ParameterName = "@TranID", Value = TranID },
                                   new SqlParameter { ParameterName = "@Revenue_Amt", Value = Revenue_Amt },
                                   new SqlParameter { ParameterName = "@Terrorism", Value = Terrorism },
                                   new SqlParameter { ParameterName = "@Insurance", Value = Insurance },
                                   new SqlParameter { ParameterName = "@Salesby", Value = Salesby },
                                   new SqlParameter { ParameterName = "@Serviceby", Value = Serviceby },
                                   new SqlParameter { ParameterName = "@location", Value = location },
                                   new SqlParameter { ParameterName = "@Support", Value = Support },
                                   new SqlParameter { ParameterName = "@Policy_Endorsement", Value = Policy_Endorsement },
                                   new SqlParameter { ParameterName = "@RFormat", Value = "F3" },
                                   new SqlParameter { ParameterName = "@InvNo", Value = "NIP3" },
                                   new SqlParameter { ParameterName = "@ReportId", Value = "NIP3" },
                                   new SqlParameter { ParameterName = "@DocName", Value = "New India Assurance Company Ltd." }
                                   );
                    }
                }
            }
        }
    }
}
