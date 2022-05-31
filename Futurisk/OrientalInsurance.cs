﻿using System;
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
using System.Windows.Forms;
//using Spire.Xls;
using System.Data.SqlClient;
using System.Configuration;
using System.Text.RegularExpressions;

namespace Futurisk
{
    public partial class OrientalInsurance : Form
    {
        static string Dir, Filepath, RDate, Result;
        static string Filename, Filenamewithext, Filewithext, name, TranID, BatchID, NoRecord;

        private void kryptonButton2_Click(object sender, EventArgs e)
        {
            Home obj = new Home();
            obj.Show();
            this.Close();
        }

        private void kryptonButton6_Click(object sender, EventArgs e)
        {
            string promptValue = ShowDialog("Batch");
            if (promptValue != "")
            {
                Fileinfo.Insurer = "OICL,Oriental Insurance Co.Ltd.";
                Fileinfo.InsurerCode = "OICL";
                Fileinfo.ReportId = "OIC1";
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
            string com = "select distinct(Inv_No) as No,Inv_No+','+[Filename] as Name from BDSMaster where InsurerCode = 'OICL' and ReportCode = 'OIC1'";
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
        private void kryptonButton3_Click(object sender, EventArgs e)
        {
            OrientalTemplate obj = new OrientalTemplate();
            obj.Show();
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
                    Fileinfo.TName = "OrientalTransaction";

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

        private void btnConvert_Click(object sender, EventArgs e)
        {
            try
            {
                var confirmResult = MessageBox.Show("Have you got the right document?", "Confirm",
                                     MessageBoxButtons.YesNo);
                if (confirmResult == DialogResult.Yes)
                {
                    DateTime a = DateTime.Now;
                    lblmsg.Text = "Please wait.......";
                    lblmsg.Refresh();
                    btnBrowse.Enabled = false;
                    btnConvert.Enabled = false;
                    btnCancel.Enabled = false;
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

                    //Aspose.Pdf.License license = new Aspose.Pdf.License();
                    //license.SetLicense(ConfigurationManager.AppSettings["aposePDFLicense"]);
                    Aspose.Pdf.Document pdfDocument = new Aspose.Pdf.Document(Filepath);
                    ExcelSaveOptions options = new ExcelSaveOptions();
                    // Set output format
                    options.Format = ExcelSaveOptions.ExcelFormat.XLSX;
                    // Minimize number of Worksheets
                    options.MinimizeTheNumberOfWorksheets = true;

                    SQLProcs sql = new SQLProcs();
                    DataSet ds1 = new DataSet();
                    ds1 = sql.SQLExecuteDataset("SP_Oriental_Transactions",
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

                        Filenamewithext = Filename + DateTime.Now.AddDays(0).ToString("ddMMyyyyhhmmss") + ".xlsx";
                        name = Dir + "\\" + Filenamewithext;
                        pdfDocument.Save(name, options);

                        IWorkbook workbook = OpenWorkBook(name);

                        ISheet sheet1 = workbook.GetSheet("Sheet1");

                        DeleteRows(sheet1);

                        SaveWorkBook(workbook, name);

                        string filewithourext = Filename + DateTime.Now.AddDays(0).ToString("ddMMyyyyhhmmss");
                        filewithourext = Dir + "\\" + filewithourext;

                        //SQLProcs sql = new SQLProcs();
                        DataSet ds = new DataSet();
                        ds = sql.SQLExecuteDataset("SP_Oriental_Transactions",
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

                        Microsoft.Office.Interop.Excel.Worksheet sheet = WB.ActiveSheet;
                        InsertTransaction(WB, TranID, RDate, Insurance, Salesby, Serviceby, location, Support, Rmonth);

                        DateTime b = DateTime.Now;
                        TimeSpan diff = b - a;
                        var Sec = String.Format("{0}", diff.Seconds);
                        lblmsg.Text = "";
                        lblSuccMsg.Text = "                 Smart Read completed in " + Sec + " Seconds.\n" +
                                          "                Batch ID: " + BatchID + " ,Number of records: " + NoRecord;
                        //linkLabel2.Text = "Click here to export the extracted data.";
                        linkLabel2.Enabled = true;
                        linkLabel2.Text = "Click here to edit data.";
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
            }
            catch (Exception ex)
            {
                lblmsg.Text = "";
                lblmsg1.Text = "              Smart Read data extraction failed.";
                lblmsg1.ForeColor = System.Drawing.Color.Red;
                btnCancel.Enabled = true;
                DDInsurance.SelectedValue = 0;
                DDLocation.SelectedValue = 0;
                DDsales.SelectedValue = 0;
                DDService.SelectedValue = 0;
                DDSupport.SelectedValue = 0;
                DDMonth.SelectedText = "";
                DDInsurance.Enabled = true;
                DDLocation.Enabled = true;
                DDsales.Enabled = true;
                DDService.Enabled = true;
                DDSupport.Enabled = true;
                DDMonth.Enabled = true;
            }
        }
        public void ExcelExport()
        {
            try
            {

                string pathUser = Environment.GetFolderPath(Environment.SpecialFolder.UserProfile);
                string pathDownload = Path.Combine(pathUser, "Downloads\\");

                SQLProcs sql = new SQLProcs();
                DataSet ResultsTable = new DataSet();

                ResultsTable = sql.SQLExecuteDataset("SP_Oriental_Transactions",
               new SqlParameter { ParameterName = "@Imode", Value = 3 },
               new SqlParameter { ParameterName = "@TranID", Value = TranID },
               new SqlParameter { ParameterName = "@DocName", Value = "Oriental Insurance Co.Ltd." }
               );

                string date = DateTime.Now.ToString();
                date = date.Replace("/", "_").Replace(":", "").Replace(" ", "").Replace("AM", "").Replace("PM", "");

                string FileName = "OIC1_" + date + ".xlsx";

                using (ClosedXML.Excel.XLWorkbook wb = new ClosedXML.Excel.XLWorkbook())
                {
                    for (int i = 0; i < ResultsTable.Tables.Count; i++)
                    {
                        wb.Worksheets.Add(ResultsTable.Tables[i], ResultsTable.Tables[i].TableName);
                    }
                    wb.Style.Alignment.Horizontal = ClosedXML.Excel.XLAlignmentHorizontalValues.Center;
                    wb.Style.Font.Bold = true;
                    string path = pathDownload + "\\" + FileName;
                    wb.SaveAs(path);

                    sql.SQLExecuteDataset("SP_Oriental_Transactions",
                    new SqlParameter { ParameterName = "@Imode", Value = 6 },
                    new SqlParameter { ParameterName = "@BatchID", Value = BatchID },
                    new SqlParameter { ParameterName = "@version", Value = LoginInfo.version },
                    new SqlParameter { ParameterName = "@ReportId", Value = "OIC1" },
                    new SqlParameter { ParameterName = "@UserId", Value = LoginInfo.UserID }
                    );

                    lblmsg1.ForeColor = System.Drawing.Color.Green;
                    lblmsg1.Text = "              Data downloaded successfully.\n     (File Name:" + FileName + ")";

                }
            }
            catch (Exception ex)
            {
                //return ex.Message;
                lblSuccMsg.Text = "";
                lblmsg1.Text = "Data export failed.";
                lblmsg1.ForeColor = System.Drawing.Color.Red;
                linkLabel2.Text = "";
                linkLabel2.Enabled = false;
            }
        }
        public void BindDDInsurance()
        {
            DataRow dr;
            //string com = "select Code,InsurerCode + ','+ UPPER(LEFT(Description, 1)) + LOWER(RIGHT(Description, LEN(Description) - 1)) as Description from tblBRInsurancelkup where GroupBy = 'OR' and Code != '' order by Description asc";
            //string com = "select Code,InsurerCode + ','+ UPPER(LEFT(Description, 1)) + LOWER(RIGHT(Description, LEN(Description) - 1)) as Description from tblBRInsurancelkup where GroupBy = 'OR' and Code != '' order by Description asc";
            //string com = "select Code,InsurerCode + ',' + Code +' '+ UPPER(LEFT(Description, 1)) + LOWER(RIGHT(Description, LEN(Description) - 1)) as Description from tblBRInsurancelkup where GroupBy = 'OR' and Code != '' order by Description asc";
            string com = "select Code,Code +' '+ InsurerCode + ',' + UPPER(LEFT(Description, 1)) + LOWER(RIGHT(Description, LEN(Description) - 1)) as Description from tblBRInsurancelkup where GroupBy = 'OR' and Code != '' order by Description asc";

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
        }

        private void linkLabel2_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            Fileinfo.Insurer = "OICL,Oriental Insurance Co.Ltd.";
            Fileinfo.InsurerCode = "OICL";
            Fileinfo.ReportId = "OIC1";
            Fileinfo.BatchId = BatchID;
            EditForm obj = new EditForm();
            obj.Show();
            obj.WindowState = FormWindowState.Normal;
        }

        private void kryptonButton1_Click(object sender, EventArgs e)
        {
            Login obj = new Login();
            obj.Show();
            this.Close();
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
        public OrientalInsurance()
        {
            InitializeComponent();
            BindDDInsurance();
            BindDDSales();
            BindDDService();
            BindDDLocation();
            BindDDSupport();
        }

        private static IWorkbook OpenWorkBook(string workBookName)
        {
            using (FileStream file = new FileStream(workBookName, FileMode.Open, FileAccess.Read))
            {
                return new XSSFWorkbook(file);
            }
        }
        private static void SaveWorkBook(IWorkbook workbook, string workBookName)
        {
            string newFileName = System.IO.Path.ChangeExtension(workBookName, "new.xlsx");
            using (FileStream file = new FileStream(newFileName, FileMode.Create, FileAccess.Write))
            {
                workbook.Write(file);
            }

            string backupFileName = System.IO.Path.ChangeExtension(workBookName, "bak.xlsx");
            File.Replace(newFileName, workBookName, backupFileName);
        }
        public static void DeleteRows(ISheet sheet)
        {
            var num = sheet.LastRowNum;
            int termsList1 = -1, termsList2 = -1, termsList3 = -1, termsList4 = -1;
            for (int rowIndex = sheet.LastRowNum; rowIndex >= 0; rowIndex--)
            {
                IRow row = sheet.GetRow(rowIndex);
                if (row == null) continue;
                ICell cell = row.GetCell(0);
                if (cell != null && cell.StringCellValue.Contains("Evaluation Only. Created with Aspose.PDF."))
                {
                    if (rowIndex != num)
                    {
                        sheet.ShiftRows(row.RowNum + 1, sheet.LastRowNum, -1);
                    }
                }
                ICell cell1 = row.GetCell(1);
                if (cell1 != null && cell1.StringCellValue.Contains("The Oriental Insurance Company Ltd."))
                {
                    sheet.ShiftRows(row.RowNum + 1, sheet.LastRowNum, -1);
                }
                ICell cell2 = row.GetCell(0);
                if (cell2 != null && cell2.StringCellValue.Contains("All premium amount shown in respect to MNAIS & WBCIS which is inclusive of Subsidy premium."))
                {
                    if (rowIndex != num - 1)
                    {
                        termsList1 = rowIndex;
                    }
                }
                ICell cell3 = row.GetCell(0);
                if (cell3 != null && cell3.StringCellValue.Contains("Description"))
                {
                    if (rowIndex != 11)
                    {
                        termsList2 = rowIndex + 1;
                    }
                }
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
            for (int rowIndex = sheet.LastRowNum; rowIndex >= 0; rowIndex--)
            {
                IRow row = sheet.GetRow(rowIndex);
                if (row == null) continue;
                ICell cell4 = row.GetCell(0);
                if (cell4 != null && cell4.StringCellValue.Contains("Department Total"))
                {
                    termsList3 = rowIndex;
                }
                ICell cell5 = row.GetCell(7);
                ICell cell6 = row.GetCell(8);
                if ((cell5 != null && cell5.StringCellValue.Contains("Net Payable")) 
                    || (cell6 != null && cell6.StringCellValue.Contains("Net Payable")))
                {
                    termsList4 = rowIndex;
                }
                if (termsList3 != -1 && termsList4 != -1)
                {

                    for (int list = termsList4; list >= termsList3; list--)
                    {
                        IRow row1 = sheet.GetRow(list);
                        sheet.ShiftRows(row1.RowNum + 1, sheet.LastRowNum, -1);
                    }
                    termsList4 = -1;
                    termsList3 = -1;
                }

                ICell cell = row.GetCell(0);
                if (cell != null && cell.StringCellValue.Contains("Office"))
                {
                    if (rowIndex != 0)
                    {
                        sheet.ShiftRows(row.RowNum + 1, sheet.LastRowNum, -1);
                    }
                }
            }
        }
        //public static void InsertTransaction(Microsoft.Office.Interop.Excel.Workbook WB, string TranID, string RDate, string Insurance, string Salesby, string Serviceby, string location, string Support, string Rmonth)
        //{
        //    SQLProcs sql = new SQLProcs();
        //    Microsoft.Office.Interop.Excel.Worksheet wks = (Microsoft.Office.Interop.Excel.Worksheet)WB.Worksheets[1];
        //    Microsoft.Office.Interop.Excel.Range lastCell = wks.Cells.SpecialCells(Microsoft.Office.Interop.Excel.XlCellType.xlCellTypeLastCell, Type.Missing);
        //    int lastrow = lastCell.Row; var Terrorism = ""; var Policy_Endorsement = ""; var InsuredType = ""; string Policy_Type = "";

        //    for (int i = 14; i < lastrow; i++)
        //    {
        //        var InsuredName = ((Microsoft.Office.Interop.Excel.Range)wks.Cells[i, 5]).Value.Replace("\n", "").TrimStart();
        //        Policy_Endorsement = ((Microsoft.Office.Interop.Excel.Range)wks.Cells[i, 3]).Value.Replace("\n", "").TrimStart();
        //        string PolicyNo = ((Microsoft.Office.Interop.Excel.Range)wks.Cells[i, 4]).Value.Replace("\n", "").TrimStart();
        //        string BillNo = ((Microsoft.Office.Interop.Excel.Range)wks.Cells[4, 2]).Value.Replace(":", "").Replace("\n", "").TrimStart();
        //        string LicenseNo = ((Microsoft.Office.Interop.Excel.Range)wks.Cells[7, 6]).Value.Replace(":", "").Replace("\n", "").TrimStart();
        //        string BRCode = ((Microsoft.Office.Interop.Excel.Range)wks.Cells[5, 2]).Value.Replace(":", "").Replace("\n", "").TrimStart();
        //        var END_Date = ((Microsoft.Office.Interop.Excel.Range)wks.Cells[i, 11]).Value.Replace("\n", "").TrimStart();
        //        var CvrType = ((Microsoft.Office.Interop.Excel.Range)wks.Cells[i, 7]).Value.Replace("\n", "").TrimStart();
        //        string Dept = ((Microsoft.Office.Interop.Excel.Range)wks.Cells[i, 1]).Value;
        //        var Premium_Amt = ((Microsoft.Office.Interop.Excel.Range)wks.Cells[i, 9]).Value.Replace("\n", "").Replace(",", "").TrimStart();
        //        var Revenue_Amt = ((Microsoft.Office.Interop.Excel.Range)wks.Cells[i, 11]).Value.Replace("\n", "").Replace(",", "").TrimStart();
        //        var Revenue_Pcnt = ((Microsoft.Office.Interop.Excel.Range)wks.Cells[i, 10]).Value.Replace("\n", "").Replace("%", "").Replace(",", "").TrimStart();
        //        BillNo = BillNo.Substring(0, BillNo.IndexOf('/'));
        //        var InstlNo = Regex.Replace(InsuredName, @"\D", "");
        //        InsuredName = Regex.Replace(InsuredName, @"\d", "");

        //        if (Dept != null && Dept != "")
        //        {
        //            bool isNo = !String.IsNullOrEmpty(Dept) && Char.IsDigit(Dept[0]);
        //            if (isNo == true)
        //            {
        //                Policy_Type = Dept.Replace("-", "").Replace("\n", "").TrimStart();
        //                Policy_Type = Regex.Replace(Policy_Type, @"\d", "");
        //            }
        //            else
        //            {
        //                Policy_Endorsement = ((Microsoft.Office.Interop.Excel.Range)wks.Cells[i, 2]).Value.Replace("\n", "").TrimStart();
        //                PolicyNo = ((Microsoft.Office.Interop.Excel.Range)wks.Cells[i, 3]).Value.Replace("\n", "").TrimStart();
        //            }
        //        }
        //        if (CvrType == "TP")
        //        {
        //            Terrorism = Premium_Amt; 
        //            Premium_Amt = "0";
        //        }

        //        if (Policy_Endorsement == "ENDMT")
        //        {
        //            Policy_Endorsement = "Endorsement";
        //            if(i+1 != lastrow)
        //            {
        //                PolicyNo = PolicyNo + ((Microsoft.Office.Interop.Excel.Range)wks.Cells[i+1, 4]).Value.Replace("\n", "").TrimStart();
        //            }
        //        }
        //        else
        //        {
        //            Policy_Endorsement = "Policy";
        //        }
        //        if (Policy_Type == "Motor TP" || Policy_Type == "Motor")
        //        {
        //            InsuredType = "Retail";
        //        }
        //        if (Policy_Endorsement != "") 
        //        {
        //            if (Premium_Amt == "" || Premium_Amt == " ")
        //            {
        //                Premium_Amt = 0;
        //            }
        //            if (Revenue_Amt == "" || Revenue_Amt == " ")
        //            {
        //                Revenue_Amt = 0;
        //            }
        //            if (Terrorism == "" || Terrorism == " ")
        //            {
        //                Terrorism = "0";
        //            }

        //            sql.ExecuteSQLNonQuery("SP_Oriental_Transactions",
        //                        new SqlParameter { ParameterName = "@Imode", Value = 1 },
        //                        new SqlParameter { ParameterName = "@RDate", Value = RDate },
        //                        new SqlParameter { ParameterName = "@Rmonth", Value = Rmonth },
        //                        new SqlParameter { ParameterName = "@ClientName", Value = InsuredName },
        //                        new SqlParameter { ParameterName = "@Itype", Value = InsuredType },
        //                        new SqlParameter { ParameterName = "@Policy_No", Value = PolicyNo },
        //                        //new SqlParameter { ParameterName = "@Endo_Effective_Date", Value = Endo_Effective_Date },
        //                        //new SqlParameter { ParameterName = "@Effective_Date", Value = Effective_Date },
        //                        new SqlParameter { ParameterName = "@END_Date", Value = END_Date },
        //                        new SqlParameter { ParameterName = "@Policy_Type", Value = Policy_Type },
        //                        new SqlParameter { ParameterName = "@Premium_Amt", Value = Premium_Amt },
        //                        new SqlParameter { ParameterName = "@TranID", Value = TranID },
        //                        new SqlParameter { ParameterName = "@Revenue_Amt", Value = Revenue_Amt },
        //                        new SqlParameter { ParameterName = "@Terrorism", Value = Terrorism },
        //                        new SqlParameter { ParameterName = "@Revenue_Pcnt", Value = Revenue_Pcnt },
        //                        new SqlParameter { ParameterName = "@InstlNo", Value = InstlNo },
        //                        new SqlParameter { ParameterName = "@BillNo", Value = BillNo },
        //                        new SqlParameter { ParameterName = "@LicenseNo", Value = LicenseNo },
        //                        new SqlParameter { ParameterName = "@BRCode", Value = BRCode },
        //                        new SqlParameter { ParameterName = "@Insurance", Value = Insurance },
        //                        new SqlParameter { ParameterName = "@Salesby", Value = Salesby },
        //                        new SqlParameter { ParameterName = "@Serviceby", Value = Serviceby },
        //                        new SqlParameter { ParameterName = "@location", Value = location },
        //                        new SqlParameter { ParameterName = "@Support", Value = Support },
        //                        new SqlParameter { ParameterName = "@Policy_Endorsement", Value = Policy_Endorsement },
        //                        new SqlParameter { ParameterName = "@RFormat", Value = "F1" },
        //                        new SqlParameter { ParameterName = "@InvNo", Value = "OIC1" },
        //                        new SqlParameter { ParameterName = "@DocName", Value = "Oriental Insurance Co.Ltd." }
        //                        );
        //        }
        //    }


        //    DataSet ds = new DataSet();

        //    ds = sql.SQLExecuteDataset("SP_Oriental_Transactions",
        //         new SqlParameter { ParameterName = "@Imode", Value = 4 },
        //         new SqlParameter { ParameterName = "@TranID", Value = TranID }
        //         );

        //    if (ds != null && ds.Tables.Count > 0 && ds.Tables[0].Rows.Count > 0)
        //    {
        //        BatchID = ds.Tables[0].Rows[0]["BatchID"].ToString();
        //    }
        //    else
        //    {
        //        BatchID = "";
        //    }

        //    DataSet dsR = new DataSet();
        //    dsR = sql.SQLExecuteDataset("SP_Oriental_Transactions",
        //         new SqlParameter { ParameterName = "@Imode", Value = 5 },
        //         new SqlParameter { ParameterName = "@BatchID", Value = BatchID },
        //         new SqlParameter { ParameterName = "@UserId", Value = LoginInfo.UserID }
        //         );
        //    if (dsR != null && dsR.Tables.Count > 0 && dsR.Tables[0].Rows.Count > 0)
        //    {
        //        NoRecord = dsR.Tables[0].Rows[0]["NoRecord"].ToString();
        //    }
        //    else
        //    {
        //        NoRecord = "";
        //    }
        //}
        public static void InsertTransaction(Microsoft.Office.Interop.Excel.Workbook WB, string TranID, string RDate, string Insurance, string Salesby, string Serviceby, string location, string Support, string Rmonth)
        {
            SQLProcs sql = new SQLProcs();
            Microsoft.Office.Interop.Excel.Worksheet wks = (Microsoft.Office.Interop.Excel.Worksheet)WB.Worksheets[1];
            Microsoft.Office.Interop.Excel.Range lastCell = wks.Cells.SpecialCells(Microsoft.Office.Interop.Excel.XlCellType.xlCellTypeLastCell, Type.Missing);
            int lastrow = lastCell.Row; var Terrorism = ""; var Policy_Endorsement = ""; var InsuredType = ""; string Policy_Type = "";
            int J = 0; int count = 0; var InsuredName = ""; var InstlNo = ""; string PolicyNo =""; string Previous = ""; int depno;
            int[] PEarray = new int[lastrow - 14];
            for (int i = 14; i < lastrow; i++)
            {
                string PEIndex = ((Microsoft.Office.Interop.Excel.Range)wks.Cells[i, 3]).Value;
                string PEIndex1 = ((Microsoft.Office.Interop.Excel.Range)wks.Cells[i, 2]).Value;
                if (PEIndex != null && PEIndex != "" && (PEIndex == "ENDMT" || PEIndex == "NEW" || PEIndex == "RENEW"))
                {
                    PEarray[J] = i;
                    J++;
                }
                if(PEIndex1 != null && PEIndex1 != "" && (PEIndex1 == "ENDMT" || PEIndex1 == "NEW" || PEIndex1 == "RENEW"))
                {
                    PEarray[J] = i;
                    J++;
                }
            }
            for (int i = 14; i < lastrow; i++)
            {
                int PEarraypos = Array.IndexOf(PEarray, i);
                if (PEarraypos > -1)
                {
                    if (PEarray[PEarraypos + 1] != 0)
                    {
                        count = PEarray[PEarraypos + 1] - i;
                    }
                    else if(PEarray[PEarraypos] != 0 && PEarray[PEarraypos + 1] == 0)
                    {
                        count = (lastrow - 1) - i;
                    }
                    //string Dept = ((Microsoft.Office.Interop.Excel.Range)wks.Cells[i, 2]).Value;
                    Policy_Endorsement = ((Microsoft.Office.Interop.Excel.Range)wks.Cells[i, 3]).Value.Replace("\n", "").TrimStart();
                    if (Policy_Endorsement != "ENDMT" && Policy_Endorsement != "NEW" && Policy_Endorsement != "RENEW")
                    {
                        Policy_Endorsement = ((Microsoft.Office.Interop.Excel.Range)wks.Cells[i, 2]).Value.Replace("\n", "").TrimStart();
                        PolicyNo = ((Microsoft.Office.Interop.Excel.Range)wks.Cells[i, 3]).Value.Replace("\n", "").TrimStart();
                        Previous = "Yes";
                    }
                    else
                    {
                        PolicyNo = ((Microsoft.Office.Interop.Excel.Range)wks.Cells[i, 4]).Value.Replace("\n", "").TrimStart();
                    }
                    var Premium_Amt = ((Microsoft.Office.Interop.Excel.Range)wks.Cells[i, 8]).Value.Replace("\n", "").Replace(",", "").Replace(" ", "").TrimStart();
                    var Revenue_Amt = ((Microsoft.Office.Interop.Excel.Range)wks.Cells[i, 10]).Value.Replace("\n", "").Replace(",", "").Replace(" ", "").TrimStart();
                    var Revenue_Pcnt = ((Microsoft.Office.Interop.Excel.Range)wks.Cells[i, 9]).Value.Replace("\n", "").Replace("%", "").Replace(",", "").TrimStart();
                    string BillNo = ((Microsoft.Office.Interop.Excel.Range)wks.Cells[4, 2]).Value.Replace(":", "").Replace("\n", "").TrimStart();
                    string LicenseNo = ((Microsoft.Office.Interop.Excel.Range)wks.Cells[7, 6]).Value.Replace(":", "").Replace("\n", "").TrimStart();
                    string BRCode = ((Microsoft.Office.Interop.Excel.Range)wks.Cells[5, 2]).Value.Replace(":", "").Replace("\n", "").TrimStart();
                    string END_Date = ((Microsoft.Office.Interop.Excel.Range)wks.Cells[i, 12]).Value.Replace("\n", "").TrimStart();
                    var CvrType = ((Microsoft.Office.Interop.Excel.Range)wks.Cells[i, 7]).Value.Replace("\n", "").TrimStart();
                    bool result = Regex.IsMatch(Premium_Amt, @"^[a-zA-Z]+$");
                    if(result == true)
                    {
                        Premium_Amt = ((Microsoft.Office.Interop.Excel.Range)wks.Cells[i, 9]).Value.Replace("\n", "").Replace(",", "").Replace(" ", "").TrimStart();
                        if (Premium_Amt == "")
                        {
                            Premium_Amt = ((Microsoft.Office.Interop.Excel.Range)wks.Cells[i, 10]).Value.Replace("\n", "").Replace(",", "").Replace(" ", "").TrimStart();
                            if (Premium_Amt == "")
                            {
                                Premium_Amt = ((Microsoft.Office.Interop.Excel.Range)wks.Cells[i, 11]).Value.Replace("\n", "").Replace(",", "").Replace(" ", "").TrimStart();
                                Revenue_Amt = ((Microsoft.Office.Interop.Excel.Range)wks.Cells[i, 13]).Value.Replace("\n", "").Replace(",", "").Replace(" ", "").TrimStart();
                                Revenue_Pcnt = ((Microsoft.Office.Interop.Excel.Range)wks.Cells[i, 12]).Value.Replace("\n", "").Replace("%", "").Replace(",", "").TrimStart();
                                if (END_Date == "" || END_Date.Contains("."))
                                {
                                    END_Date = ((Microsoft.Office.Interop.Excel.Range)wks.Cells[i, 15]).Value.Replace("\n", "").TrimStart();
                                }
                            }
                            else
                            {
                                Revenue_Amt = ((Microsoft.Office.Interop.Excel.Range)wks.Cells[i, 12]).Value.Replace("\n", "").Replace(",", "").Replace(" ", "").TrimStart();
                                Revenue_Pcnt = ((Microsoft.Office.Interop.Excel.Range)wks.Cells[i, 11]).Value.Replace("\n", "").Replace("%", "").Replace(",", "").TrimStart();
                                if (END_Date == "" || END_Date.Contains("."))
                                {
                                    END_Date = ((Microsoft.Office.Interop.Excel.Range)wks.Cells[i, 14]).Value.Replace("\n", "").TrimStart();
                                }
                            }
                        }
                        else
                        {
                            Revenue_Amt = ((Microsoft.Office.Interop.Excel.Range)wks.Cells[i, 11]).Value.Replace("\n", "").Replace(",", "").Replace(" ", "").TrimStart();
                            Revenue_Pcnt = ((Microsoft.Office.Interop.Excel.Range)wks.Cells[i, 10]).Value.Replace("\n", "").Replace("%", "").Replace(",", "").TrimStart();
                            CvrType = ((Microsoft.Office.Interop.Excel.Range)wks.Cells[i, 8]).Value.Replace("\n", "").TrimStart();
                            if (END_Date == "" || END_Date.Contains("."))
                            {
                                END_Date = ((Microsoft.Office.Interop.Excel.Range)wks.Cells[i, 13]).Value.Replace("\n", "").TrimStart();
                            }
                        }
                    }

                    BillNo = BillNo.Substring(0, BillNo.IndexOf('/'));
                    if (Policy_Endorsement == "ENDMT")
                    {
                        Policy_Endorsement = "Endorsement";
                        if (i + 1 != lastrow)
                        {
                            if (Previous == "Yes")
                            {
                                PolicyNo = PolicyNo + ((Microsoft.Office.Interop.Excel.Range)wks.Cells[i + 1, 3]).Value.Replace("\n", "").TrimStart();
                            }
                            else
                            {
                                PolicyNo = PolicyNo + ((Microsoft.Office.Interop.Excel.Range)wks.Cells[i + 1, 4]).Value.Replace("\n", "").TrimStart();
                            }
                        }
                    }
                    else
                    {
                        Policy_Endorsement = "Policy";
                    }
                    if (Premium_Amt == "" || Premium_Amt == " ")
                    {
                        Premium_Amt = 0;
                    }
                    if (Revenue_Amt == "" || Revenue_Amt == " ")
                    {
                        Revenue_Amt = 0;
                    }
                    if (Terrorism == "" || Terrorism == " ")
                    {
                        Terrorism = "0";
                    }
                    if (Revenue_Pcnt == "" || Revenue_Pcnt == " ")
                    {
                        Revenue_Pcnt = "0";
                    }
                    if (result == false)
                    {
                        for (int K = 0; K < count; K++)
                        {
                            var IName = ((Microsoft.Office.Interop.Excel.Range)wks.Cells[i + K, 5]).Value;
                            if (IName == null || IName == "")
                            {
                                if (Previous == "Yes")
                                {
                                    string IN = ((Microsoft.Office.Interop.Excel.Range)wks.Cells[i + K, 4]).Value;
                                    IName = IN.Replace("\n", "").TrimStart();
                                }
                                else
                                {
                                    string IN = ((Microsoft.Office.Interop.Excel.Range)wks.Cells[i + K, 4]).Value;
                                    bool isNo = IN.Any(char.IsDigit);
                                    if (isNo == false)
                                    {
                                        IName = IN.Replace("\n", "").TrimStart();
                                    }
                                }
                            }
                            else
                            {
                                IName = IName.Replace("\n", "").TrimStart();
                            }
                            InsuredName = InsuredName + IName;
                            if (Previous == "Yes")
                            {
                                depno = 1;
                            }
                            else
                            {
                                depno = 2;
                            }
                            string Dept = ((Microsoft.Office.Interop.Excel.Range)wks.Cells[i + K, depno]).Value;
                            if (Dept != null && Dept != "")
                            {
                                Policy_Type = Policy_Type + Dept;
                            }
                        }
                        InstlNo = Regex.Replace(InsuredName, @"\D", "");
                        InsuredName = Regex.Replace(InsuredName, @"\d", "");
                    }
                    else
                    {
                        for (int K = 0; K < count; K++)
                        {
                            var IName = ((Microsoft.Office.Interop.Excel.Range)wks.Cells[i + K, 6]).Value;
                            if (IName == null || IName == "")
                            {
                                string IN = ((Microsoft.Office.Interop.Excel.Range)wks.Cells[i + K, 4]).Value;
                                bool isNo = IN.Any(char.IsDigit);
                                if (isNo == false)
                                {
                                    IName = IN.Replace("\n", "").TrimStart();
                                }
                                if (IN == "")
                                {
                                    string IN1 = ((Microsoft.Office.Interop.Excel.Range)wks.Cells[i + K, 3]).Value;
                                    IName = IN1.Replace("\n", "").TrimStart();
                                    if (IN1 == "")
                                    {
                                        IName = ((Microsoft.Office.Interop.Excel.Range)wks.Cells[i + K, 2]).Value.Replace("\n", "").TrimStart();
                                    }
                                }
                            }
                            else
                            {
                                IName = IName.Replace("\n", "").TrimStart();
                            }
                            InsuredName = InsuredName + IName;
                            if (Previous == "Yes")
                            {
                                depno = 1;
                            }
                            else
                            {
                                depno = 2;
                            }
                            string Dept = ((Microsoft.Office.Interop.Excel.Range)wks.Cells[i + K, depno]).Value.Trim();
                            if (Dept != null && Dept != "")
                            {
                                Policy_Type = Policy_Type + Dept;
                            }
                        }
                        InstlNo = ((Microsoft.Office.Interop.Excel.Range)wks.Cells[i , 5]).Value;
                    }
                    sql.ExecuteSQLNonQuery("SP_Oriental_Transactions",
                                new SqlParameter { ParameterName = "@Imode", Value = 1 },
                                new SqlParameter { ParameterName = "@RDate", Value = RDate },
                                new SqlParameter { ParameterName = "@Rmonth", Value = Rmonth },
                                new SqlParameter { ParameterName = "@ClientName", Value = InsuredName },
                                new SqlParameter { ParameterName = "@Itype", Value = InsuredType },
                                new SqlParameter { ParameterName = "@Policy_No", Value = PolicyNo },
                                new SqlParameter { ParameterName = "@END_Date", Value = END_Date },
                                new SqlParameter { ParameterName = "@Policy_Type", Value = Policy_Type },
                                new SqlParameter { ParameterName = "@Premium_Amt", Value = Premium_Amt },
                                new SqlParameter { ParameterName = "@TranID", Value = TranID },
                                new SqlParameter { ParameterName = "@Revenue_Amt", Value = Revenue_Amt },
                                new SqlParameter { ParameterName = "@Terrorism", Value = Terrorism },
                                new SqlParameter { ParameterName = "@Revenue_Pcnt", Value = Revenue_Pcnt },
                                new SqlParameter { ParameterName = "@InstlNo", Value = InstlNo },
                                new SqlParameter { ParameterName = "@BillNo", Value = BillNo },
                                new SqlParameter { ParameterName = "@LicenseNo", Value = LicenseNo },
                                new SqlParameter { ParameterName = "@BRCode", Value = BRCode },
                                new SqlParameter { ParameterName = "@Insurance", Value = Insurance },
                                new SqlParameter { ParameterName = "@Salesby", Value = Salesby },
                                new SqlParameter { ParameterName = "@Serviceby", Value = Serviceby },
                                new SqlParameter { ParameterName = "@location", Value = location },
                                new SqlParameter { ParameterName = "@Support", Value = Support },
                                new SqlParameter { ParameterName = "@Policy_Endorsement", Value = Policy_Endorsement },
                                new SqlParameter { ParameterName = "@RFormat", Value = "F1" },
                                new SqlParameter { ParameterName = "@InvNo", Value = "OIC1" },
                                new SqlParameter { ParameterName = "@ReportId", Value = "OIC1" },
                                new SqlParameter { ParameterName = "@DocName", Value = "Oriental Insurance Co.Ltd." }
                                );
                    InsuredName = "";
                    Policy_Type = "";
                }                
                            
            }


            DataSet ds = new DataSet();

            ds = sql.SQLExecuteDataset("SP_Oriental_Transactions",
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
            dsR = sql.SQLExecuteDataset("SP_Oriental_Transactions",
                 new SqlParameter { ParameterName = "@Imode", Value = 5 },
                 new SqlParameter { ParameterName = "@BatchID", Value = BatchID },
                 new SqlParameter { ParameterName = "@Filename", Value = Fileinfo.Filename },
                 new SqlParameter { ParameterName = "@version", Value = LoginInfo.version },
                 new SqlParameter { ParameterName = "@ReportId", Value = "OIC1" },
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
    }
}
