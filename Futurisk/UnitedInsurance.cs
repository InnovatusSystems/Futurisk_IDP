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
using System.Windows.Forms;
//using Spire.Xls;
using System.Data.SqlClient;
using System.Configuration;

namespace Futurisk
{
    public partial class UnitedInsurance : Form
    {
        static string Dir, Filepath, RDate;
        static string Filename, Filenamewithext, Filewithext, name, TranID, BatchID, NoRecord;
        private string strconn = ConfigurationManager.ConnectionStrings["IDP"].ToString();
        private void btnConvert_Click(object sender, EventArgs e)
        {
            try
            {
                var confirmResult = MessageBox.Show("Kindly ensure the selected data before extract the data to the Smart Reader database.\nDo you want to continue?", "Confirm",
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
                    //Aspose.Pdf.License license = new Aspose.Pdf.License();
                    //license.SetLicense(ConfigurationManager.AppSettings["aposePDFLicense"]);
                    Aspose.Pdf.Document pdfDocument = new Aspose.Pdf.Document(Filepath);
                    ExcelSaveOptions options = new ExcelSaveOptions();
                    // Set output format
                    options.Format = ExcelSaveOptions.ExcelFormat.XLSX;
                    // Minimize number of Worksheets
                    options.MinimizeTheNumberOfWorksheets = true;

                    Filenamewithext = Filename + DateTime.Now.AddDays(0).ToString("ddMMyyyyhhmmss") + ".xlsx";
                    name = Dir + "\\" + Filenamewithext;
                    pdfDocument.Save(name, options);

                    IWorkbook workbook = OpenWorkBook(name);

                    ISheet sheet1 = workbook.GetSheet("Sheet1");

                    DeleteRows(sheet1);

                    SaveWorkBook(workbook, name);

                    string filewithourext = Filename + DateTime.Now.AddDays(0).ToString("ddMMyyyyhhmmss");
                    filewithourext = Dir + "\\" + filewithourext;

                    SQLProcs sql = new SQLProcs();
                    DataSet ds = new DataSet();
                    ds = sql.SQLExecuteDataset("SP_Insert_Transactions",
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

                    RDate = DateTime.Now.AddDays(0).ToString("dd-MM-yyyy");
                    Microsoft.Office.Interop.Excel.Application oExcel = new Microsoft.Office.Interop.Excel.Application();
                    Microsoft.Office.Interop.Excel.Workbook WB = oExcel.Workbooks.Open(name);

                    Microsoft.Office.Interop.Excel.Worksheet sheet = WB.ActiveSheet;
                    InsertTransaction(WB, TranID, RDate, Insurance, Salesby, Serviceby, location, Support);
                    //SaveDB(sheet, WB, Filename);

                    //Workbook book = new Workbook();
                    //book.LoadFromFile(name);
                    ////book.Replace(",","");
                    //book.SaveToFile(filewithourext + ".xls", ExcelVersion.Version97to2003);

                    //xlsFilename = filewithourext + ".xls";

                    DateTime b = DateTime.Now;
                    TimeSpan diff = b - a;
                    var Sec = String.Format("{0}", diff.Seconds);
                    lblmsg.Text = "";
                    lblSuccMsg.Text = "Smart Read completed in " + Sec + " Seconds, Batch ID: " + BatchID + "\n" +
                                      "                     Number of records: " + NoRecord;
                    //linkLabel2.Text = "Click here to export the extracted data.";
                    oExcel.Workbooks.Close();
                    var confirmExportResult = MessageBox.Show("Do you wish to download the data copied from the Smart Reader database?", "Confirm",
                                    MessageBoxButtons.YesNo);
                    if (confirmExportResult == DialogResult.Yes)
                    {
                        ExcelExport();
                    }
                     
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
            }
            catch (Exception ex)
            {
                lblmsg.Text = "";
                lblmsg1.Text = "Smart Read data extraction failed.";
                lblmsg1.ForeColor = System.Drawing.Color.Red;
                btnCancel.Enabled = true;
                DDInsurance.SelectedValue = 0;
                DDLocation.SelectedValue = 0;
                DDsales.SelectedValue = 0;
                DDService.SelectedValue = 0;
                DDSupport.SelectedValue = 0;
                DDInsurance.Enabled = true;
                DDLocation.Enabled = true;
                DDsales.Enabled = true;
                DDService.Enabled = true;
                DDSupport.Enabled = true;
            }
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
            DDInsurance.Enabled = true;
            DDLocation.Enabled = true;
            DDsales.Enabled = true;
            DDService.Enabled = true;
            DDSupport.Enabled = true;
        }

        private void linkLabel2_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            ExcelExport();
        }

        private void kryptonButton3_Click(object sender, EventArgs e)
        {
            UnitedTemplate obj = new UnitedTemplate();
            obj.Show();
        }

        private void kryptonButton2_Click(object sender, EventArgs e)
        {
            Home obj = new Home();
            obj.Show();
            this.Close();
        }

        private void kryptonButton6_Click(object sender, EventArgs e)
        {
            Login obj = new Login();
            obj.Show();
            this.Close();
        }

        public UnitedInsurance()
        {
            InitializeComponent();
            BindDDInsurance();
            BindDDSales();
            BindDDService();
            BindDDLocation();
            BindDDSupport();
        }
        public void BindDDInsurance()
        {
            DataRow dr;
            string com = "select Code,InsurerCode + ','+ UPPER(LEFT(Description, 1)) + LOWER(RIGHT(Description, LEN(Description) - 1)) as Description from tblBRInsurancelkup where GroupBy = 'UN' and Code != '' order by Description asc";
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

        private void DDService_SelectedIndexChanged(object sender, EventArgs e)
        {
            if(DDService.SelectedValue.ToString() != "0")
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
        private void kryptonButton1_Click(object sender, EventArgs e)
        {
            lblSuccMsg.Text = "";
            if (DDInsurance.SelectedValue.ToString() != "0" && DDLocation.SelectedValue.ToString() != "0" && DDsales.SelectedValue.ToString() != "0" && DDService.SelectedValue.ToString() != "0")
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
                btnConvert.Enabled = false;
                btnCancel.Enabled = false;
                btnBrowse.Enabled = true;
            }
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
            int termsList1 = -1, termsList2 = -1;//, termsList3 = -1, termsList4 = -1;
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
                //ICell cell1 = row.GetCell(2);
                //if (cell1 != null && cell1.StringCellValue.Contains("OPENING BALANCE"))
                //{
                //    sheet.ShiftRows(row.RowNum + 1, sheet.LastRowNum, -1);
                //}
                //ICell cell5 = row.GetCell(2);
                //if (cell5 != null && cell5.StringCellValue.Contains("TRANSACTION TOTAL DR/CR"))
                //{
                //  termsList1 = num - 1;
                //}
                ICell cell6 = row.GetCell(0);
                if (cell6 != null && cell6.StringCellValue.Contains("Dept Code"))
                {
                    if (rowIndex != 7 && rowIndex != 8)
                    {
                        termsList1 = rowIndex;
                    }
                }
                termsList2 = num - 1;

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
        public void ExcelExport()
        {
            try
            {

                string pathUser = Environment.GetFolderPath(Environment.SpecialFolder.UserProfile);
                string pathDownload = Path.Combine(pathUser, "Downloads\\");

                SQLProcs sql = new SQLProcs();
                DataSet ResultsTable = new DataSet();

                ResultsTable = sql.SQLExecuteDataset("SP_Insert_Transactions",
               new SqlParameter { ParameterName = "@Imode", Value = 3 },
               new SqlParameter { ParameterName = "@TranID", Value = TranID },
               new SqlParameter { ParameterName = "@DocName", Value = "United India Insurance Co.Ltd." }
               );

                string date = DateTime.Now.ToString();
                date = date.Replace("/", "_").Replace(":", "").Replace(" ", "").Replace("AM", "").Replace("PM", "");

                //SaveFileDialog saveDlg = new SaveFileDialog();
                //saveDlg.InitialDirectory = @"C:\";
                //saveDlg.Filter = "Excel files (*.xlsx)|*.xlsx";
                //saveDlg.FilterIndex = 0;
                //saveDlg.RestoreDirectory = true;
                //saveDlg.Title = "Export Excel File To";
                //saveDlg.FileName = "United India Insurance_" + date;

                string FileName = pathDownload + "\\United India Insurance_" + date + ".xlsx";

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
                    string path = FileName;
                    wb.SaveAs(path);
                        //result = "OK";
                        //lblSuccMsg.Text = "";
                        lblmsg1.ForeColor = System.Drawing.Color.Green;
                        lblmsg1.Text = "Data downloaded successfully.";
                        //linkLabel2.Text = "";
                        //linkLabel2.Enabled = false;
                        //DDInsurance.SelectedValue = 0;
                        //DDLocation.SelectedValue = 0;
                        //DDsales.SelectedValue = 0;
                        //DDService.SelectedValue = 0;
                        //DDSupport.SelectedValue = 0;
                    //}
                    //else
                    //{
                    //    lblSuccMsg.Text = "";
                    //    lblmsg1.ForeColor = System.Drawing.Color.Red;
                    //    lblmsg1.Text = "Data export canceled.";
                    //    linkLabel2.Text = "";
                    //    linkLabel2.Enabled = false;
                    //    DDInsurance.SelectedValue = 0;
                    //    DDLocation.SelectedValue = 0;
                    //    DDsales.SelectedValue = 0;
                    //    DDService.SelectedValue = 0;
                    //    DDSupport.SelectedValue = 0;
                    //}
                }
                //return result;
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
        //public void ExcelExport()
        //{
        //    try
        //    {
        //        SQLProcs sql = new SQLProcs();
        //        DataSet ResultsTable = new DataSet();

        //        ResultsTable = sql.SQLExecuteDataset("SP_Insert_Transactions",
        //       new SqlParameter { ParameterName = "@Imode", Value = 3 },
        //       new SqlParameter { ParameterName = "@TranID", Value = TranID },
        //       new SqlParameter { ParameterName = "@DocName", Value = "United India Insurance Co.Ltd." }
        //       );

        //        string date = DateTime.Now.ToString();
        //        date = date.Replace("/", "_").Replace(":", "").Replace(" ", "").Replace("AM", "").Replace("PM", "");

        //        SaveFileDialog saveDlg = new SaveFileDialog();
        //        saveDlg.InitialDirectory = @"C:\";
        //        saveDlg.Filter = "Excel files (*.xlsx)|*.xlsx";
        //        saveDlg.FilterIndex = 0;
        //        saveDlg.RestoreDirectory = true;
        //        saveDlg.Title = "Export Excel File To";
        //        saveDlg.FileName = "United India Insurance_" + date;

        //        using (ClosedXML.Excel.XLWorkbook wb = new ClosedXML.Excel.XLWorkbook())
        //        {
        //            for (int i = 0; i < ResultsTable.Tables.Count; i++)
        //            {
        //                wb.Worksheets.Add(ResultsTable.Tables[i], ResultsTable.Tables[i].TableName);
        //            }
        //            wb.Style.Alignment.Horizontal = ClosedXML.Excel.XLAlignmentHorizontalValues.Center;
        //            wb.Style.Font.Bold = true;
        //            if (saveDlg.ShowDialog() == DialogResult.OK)
        //            {
        //                string path = saveDlg.FileName;
        //                wb.SaveAs(path);
        //                //result = "OK";
        //                lblSuccMsg.Text = "";
        //                lblmsg1.ForeColor = System.Drawing.Color.Green;
        //                lblmsg1.Text = "Data exported successfully.";
        //                linkLabel2.Text = "";
        //                linkLabel2.Enabled = false;
        //                DDInsurance.SelectedValue = 0;
        //                DDLocation.SelectedValue = 0;
        //                DDsales.SelectedValue = 0;
        //                DDService.SelectedValue = 0;
        //                DDSupport.SelectedValue = 0;
        //            }
        //            else
        //            {
        //                lblSuccMsg.Text = "";
        //                lblmsg1.ForeColor = System.Drawing.Color.Red;
        //                lblmsg1.Text = "Data export canceled.";
        //                linkLabel2.Text = "";
        //                linkLabel2.Enabled = false;
        //                DDInsurance.SelectedValue = 0;
        //                DDLocation.SelectedValue = 0;
        //                DDsales.SelectedValue = 0;
        //                DDService.SelectedValue = 0;
        //                DDSupport.SelectedValue = 0;
        //            }
        //        }
        //        //return result;
        //    }
        //    catch (Exception ex)
        //    {
        //        //return ex.Message;
        //        lblSuccMsg.Text = "";
        //        lblmsg1.Text = "Data export failed.";
        //        lblmsg1.ForeColor = System.Drawing.Color.Red;
        //        linkLabel2.Text = "";
        //        linkLabel2.Enabled = false;
        //    }
        //}
        public static void InsertTransaction(Microsoft.Office.Interop.Excel.Workbook WB, string TranID, string RDate, string Insurance,string Salesby,string Serviceby,string location,string Support)
        {
            SQLProcs sql = new SQLProcs();
            Microsoft.Office.Interop.Excel.Worksheet wks = (Microsoft.Office.Interop.Excel.Worksheet)WB.Worksheets[1];
            Microsoft.Office.Interop.Excel.Range lastCell = wks.Cells.SpecialCells(Microsoft.Office.Interop.Excel.XlCellType.xlCellTypeLastCell, Type.Missing);
            int lastrow = lastCell.Row; var Terrorism = "";var Policy_Endorsement = "";

            for (int i = 9; i < lastrow; i++)
            {
                var InsuredName = ((Microsoft.Office.Interop.Excel.Range)wks.Cells[i, 4]).Value.Replace("\n", "").TrimStart();
                var InsuredType = ((Microsoft.Office.Interop.Excel.Range)wks.Cells[i, 9]).Value.Replace("\n", "").TrimStart();
                string PolicyNo = ((Microsoft.Office.Interop.Excel.Range)wks.Cells[i, 3]).Value.Replace("\n", "").TrimStart();
                string BillNo = ((Microsoft.Office.Interop.Excel.Range)wks.Cells[4, 7]).Value.Replace("\n", "").TrimStart();
                string LicenseNo = ((Microsoft.Office.Interop.Excel.Range)wks.Cells[6, 7]).Value.Replace("\n", "").TrimStart();
                string BRCode = ((Microsoft.Office.Interop.Excel.Range)wks.Cells[6, 2]).Value;
                string Endo_Effective_Date = ((Microsoft.Office.Interop.Excel.Range)wks.Cells[5, 7]).Value.Replace("\n", "").TrimStart();
                var Effective_Date = ((Microsoft.Office.Interop.Excel.Range)wks.Cells[5, 7]).Value.Replace("\n", "").TrimStart();
                var END_Date = ((Microsoft.Office.Interop.Excel.Range)wks.Cells[i, 5]).Value.Replace("\n", "").TrimStart();
                string Policy_Type = ((Microsoft.Office.Interop.Excel.Range)wks.Cells[i, 2]).Value.Replace("\n", "").TrimStart();
                var Premium_Amt = ((Microsoft.Office.Interop.Excel.Range)wks.Cells[i, 6]).Value.Replace("\n", "").Replace(",", "").TrimStart();
                var Revenue_Amt = ((Microsoft.Office.Interop.Excel.Range)wks.Cells[i, 8]).Value.Replace("\n", "").Replace(",", "").TrimStart();
                var Ineligible_Amt = ((Microsoft.Office.Interop.Excel.Range)wks.Cells[i, 7]).Value.Replace("\n", "").Replace(",", "").TrimStart();
                var DepCode = ((Microsoft.Office.Interop.Excel.Range)wks.Cells[i, 1]).Value.Replace("\n", "").TrimStart();
                var Presult = PolicyNo.Substring(PolicyNo.LastIndexOf('/') + 1);
                if (Policy_Type.Contains("Motor TP"))
                {
                    Terrorism = ((Microsoft.Office.Interop.Excel.Range)wks.Cells[i, 6]).Value.Replace("\n", "").Replace(",", "").TrimStart();
                    Premium_Amt = "0";
                }
                if (Presult.Contains("0"))
                {
                    Policy_Endorsement = "Policy";
                }
                else
                {
                    Policy_Endorsement = "Endorsement";
                }
                if(InsuredType == "Individual")
                {
                    InsuredType = "Retail";
                }
                if (Policy_Type == "Motor TP" || Policy_Type == "Motor" )
                {
                    InsuredType = "Retail";
                }
                if (Endo_Effective_Date.Contains("Bill Date"))
                {
                    var BillDate = ((Microsoft.Office.Interop.Excel.Range)wks.Cells[5, 8]).Value;
                    if (BillDate != "" && BillDate != " ")
                    {
                        Endo_Effective_Date = ((Microsoft.Office.Interop.Excel.Range)wks.Cells[5, 8]).Value.Replace("\n", "").TrimStart();
                        Effective_Date = ((Microsoft.Office.Interop.Excel.Range)wks.Cells[5, 8]).Value.Replace("\n", "").TrimStart();
                    }
                }
                if (BillNo.Contains("Bill Number"))
                {
                    BillNo = ((Microsoft.Office.Interop.Excel.Range)wks.Cells[4, 8]).Value.Replace("\n", "").TrimStart();
                }
                if (LicenseNo.Contains("License Number"))
                {
                    LicenseNo = ((Microsoft.Office.Interop.Excel.Range)wks.Cells[6, 8]).Value.Replace("\n", "").TrimStart();
                }
                if (BRCode == "" || BRCode == " ")
                {
                    BRCode = ((Microsoft.Office.Interop.Excel.Range)wks.Cells[6, 3]).Value.Replace("\n", "").TrimStart();
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
                if (Ineligible_Amt == "" || Ineligible_Amt == " ")
                {
                    Ineligible_Amt = "0";
                }

                sql.ExecuteSQLNonQuery("SP_Insert_Transactions",
                            new SqlParameter { ParameterName = "@Imode", Value = 1 },
                            new SqlParameter { ParameterName = "@RDate", Value = RDate },
                            new SqlParameter { ParameterName = "@ClientName", Value = InsuredName },
                            new SqlParameter { ParameterName = "@Itype", Value = InsuredType },
                            new SqlParameter { ParameterName = "@Policy_No", Value = PolicyNo },
                            new SqlParameter { ParameterName = "@Endo_Effective_Date", Value = Endo_Effective_Date },
                            new SqlParameter { ParameterName = "@Effective_Date", Value = Effective_Date },
                            new SqlParameter { ParameterName = "@END_Date", Value = END_Date },
                            new SqlParameter { ParameterName = "@Policy_Type", Value = Policy_Type },
                            new SqlParameter { ParameterName = "@Premium_Amt", Value = Premium_Amt },
                            new SqlParameter { ParameterName = "@TranID", Value = TranID },
                            new SqlParameter { ParameterName = "@Revenue_Amt", Value = Revenue_Amt },
                            new SqlParameter { ParameterName = "@Terrorism", Value = Terrorism },
                            new SqlParameter { ParameterName = "@Ineligible_Amt", Value = Ineligible_Amt },
                            new SqlParameter { ParameterName = "@DeptCode", Value = DepCode },
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
                            new SqlParameter { ParameterName = "@InvNo", Value = "UIIP" },
                            new SqlParameter { ParameterName = "@DocName", Value = "United India Insurance Co.Ltd." }
                            );
            }


            DataSet ds = new DataSet();

            ds = sql.SQLExecuteDataset("SP_Insert_Transactions",
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
            dsR = sql.SQLExecuteDataset("SP_Insert_Transactions",
                 new SqlParameter { ParameterName = "@Imode", Value = 5 },
                 new SqlParameter { ParameterName = "@BatchID", Value = BatchID }
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
