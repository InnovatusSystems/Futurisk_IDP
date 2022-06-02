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
using Spire.Xls;
using System.Data.SqlClient;
using System.Configuration;

namespace Futurisk
{
    public partial class GodigitInsurence : Form
    {
        static string Dir, Filepath, RDate, Result,Fileextn,DelFile;
        static string Filename, Filenamewithext, Filewithext, name, TranID, BatchID, NoRecord;

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

                    ////Aspose.Pdf.License license = new Aspose.Pdf.License();
                    ////license.SetLicense(ConfigurationManager.AppSettings["aposePDFLicense"]);
                    //Aspose.Pdf.Document pdfDocument = new Aspose.Pdf.Document(Filepath);
                    //ExcelSaveOptions options = new ExcelSaveOptions();
                    //// Set output format
                    //options.Format = ExcelSaveOptions.ExcelFormat.XLSX;
                    //// Minimize number of Worksheets
                    //options.MinimizeTheNumberOfWorksheets = true;


                    SQLProcs sql = new SQLProcs();
                    //DataSet ds1 = new DataSet();
                    //ds1 = sql.SQLExecuteDataset("SP_GoDigitTransactions",
                    //              new SqlParameter { ParameterName = "@Imode", Value = 7 },
                    //              new SqlParameter { ParameterName = "@Filename", Value = Fileinfo.Filename }
                    //     );
                    //if (ds1 != null && ds1.Tables.Count > 0 && ds1.Tables[0].Rows.Count > 0)
                    //{
                    //    Result = ds1.Tables[0].Rows[0]["Result"].ToString();
                    //}
                    //else
                    //{
                    //    Result = "Not Exists";
                    //}
                    //if (Result == "Not Exists")
                    //{
                    //Filenamewithext = Filename + DateTime.Now.AddDays(0).ToString("ddMMyyyyhhmmss") + ".xlsx";
                    //name = Dir + "\\" + Filename + ".xlsx";
                    if (Fileextn == ".xls")
                    {
                        name = Dir + "\\" + Filename + DateTime.Now.AddDays(0).ToString("ddMMyyyyhhmmss") + ".xlsx";
                        DelFile = name;
                        Workbook workbook = new Workbook();
                        workbook.LoadFromFile(Filepath);
                        workbook.SaveToFile(name, ExcelVersion.Version2013);
                    }
                    else
                    {
                        name = Filepath;
                    }

                    //pdfDocument.Save(name, options);

                    //IWorkbook workbook = OpenWorkBook(name);

                    //ISheet sheet1 = workbook.GetSheet("Sheet1");

                    //DeleteRows(sheet1);

                    //SaveWorkBook(workbook, name);

                    string filewithourext = Filename + DateTime.Now.AddDays(0).ToString("ddMMyyyyhhmmss");
                        filewithourext = Dir + "\\" + filewithourext;

                        DataSet ds = new DataSet();
                        ds = sql.SQLExecuteDataset("SP_GoDigitTransactions",
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
                        lblSuccMsg.Text = "Smart Read completed in " + Sec + " Seconds, Batch ID: " + BatchID + "\n" +
                                          "                     Number of records: " + NoRecord;
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
                            //ExcelExport();
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
                        btnCancel.Enabled = true;
                    if (File.Exists(DelFile))
                    {
                        File.Delete(DelFile);
                    }
                    //}
                    //else
                    //{
                    //    lblmsg.Text = "";
                    //    lblmsg.Refresh();
                    //    MessageBox.Show("Given file's BDS is already exists in the database.", "Warning!");
                    //    btnCancel.Enabled = true; btnBrowse.Enabled = true;
                    //    DDInsurance.SelectionLength = 0;
                    //    DDLocation.SelectionLength = 0;
                    //    DDsales.SelectionLength = 0;
                    //    DDService.SelectionLength = 0;
                    //    DDSupport.SelectionLength = 0;
                    //    DDMonth.SelectionLength = 0;
                    //}
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
                lblmsg1.Text = "Smart Read data extraction failed.";
                lblmsg1.ForeColor = System.Drawing.Color.Red;
                btnCancel.Enabled = true;
                DDMonth.SelectedIndex = 0;
                DDInsurance.SelectedValue = 0;
                DDLocation.SelectedValue = 0;
                DDsales.SelectedValue = 0;
                DDService.SelectedValue = 0;
                DDSupport.SelectedValue = 0;
                //DDMonth.SelectedText = "";
                DDInsurance.Enabled = true;
                DDLocation.Enabled = true;
                DDsales.Enabled = true;
                DDService.Enabled = true;
                DDSupport.Enabled = true;
                DDMonth.Enabled = true;
                DDInsurance.SelectionLength = 0;
                DDLocation.SelectionLength = 0;
                DDsales.SelectionLength = 0;
                DDService.SelectionLength = 0;
                DDSupport.SelectionLength = 0;
                DDMonth.SelectionLength = 0;
            }
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
            string com = "select distinct(Inv_No) as No,Inv_No+','+[Filename] as Name from BDSMaster where InsurerCode = 'GGIC' and ReportCode = 'GGI1'";
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
        private string strconn = ConfigurationManager.ConnectionStrings["IDP"].ToString();
        public GodigitInsurence()
        {
            InitializeComponent(); 
            BindDDInsurance();
            BindDDSales();
            BindDDService();
            BindDDLocation();
            BindDDSupport();
            DDMonth.Select();
        }
        public void BindDDInsurance()
        {
            DataRow dr;
            //string com = "select Code,InsurerCode + ','+ UPPER(LEFT(Description, 1)) + LOWER(RIGHT(Description, LEN(Description) - 1)) as Description from tblBRInsurancelkup where GroupBy = 'UN' and Code != '' order by Description asc";
            //string com = "select Code,InsurerCode + ',' + Code +' '+ UPPER(LEFT(Description, 1)) + LOWER(RIGHT(Description, LEN(Description) - 1)) as Description from tblBRInsurancelkup where GroupBy = 'UN' and Code != '' order by Description asc";
            string com = "select ROW_NUMBER() OVER (ORDER BY (SELECT 1)) as Code,InsurerCode + ',' + UPPER(LEFT(Description, 1)) + LOWER(RIGHT(Description, LEN(Description) - 1)) as Description from tblBRInsurancelkup where GroupBy = 'GO' order by Description asc";
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
        private void btnBrowse_Click(object sender, EventArgs e)
        {
            lblSuccMsg.Text = "";

            //if (DDInsurance.SelectedValue.ToString() != "0" && DDLocation.SelectedValue.ToString() != "0" && DDsales.SelectedValue.ToString() != "0" && DDService.SelectedValue.ToString() != "0" && (DDMonth.SelectedIndex != -1 && DDMonth.SelectedIndex != 0))
            //{
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
                    Fileextn = System.IO.Path.GetExtension(openFileDialog1.FileName);

                    txtfile.Text = Filewithext;
                    txtfile.ForeColor = System.Drawing.Color.Black;
                    Fileinfo.Filename = Filewithext;
                    Fileinfo.TName = "UnitedTransaction";
                    btnBrowse.Enabled = false;
                }
                else
                {
                    btnBrowse.Enabled = true;
                    btnCancel.Enabled = true;
                    btnConvert.Enabled = false;
                }
            //}
            //else
            //{
            //    if (DDInsurance.SelectedValue.ToString() == "0")
            //    {
            //        MessageBox.Show("Please select Insurance");
            //    }
            //    else if (DDLocation.SelectedValue.ToString() == "0")
            //    {
            //        MessageBox.Show("Please select Office Location");
            //    }
            //    else if (DDsales.SelectedValue.ToString() == "0")
            //    {
            //        MessageBox.Show("Please select Sales Generated By");
            //    }
            //    else if (DDService.SelectedValue.ToString() == "0")
            //    {
            //        MessageBox.Show("Please select Serviced By");
            //    }
            //    else if (DDMonth.SelectedIndex == -1 || DDMonth.SelectedIndex == 0)
            //    {
            //        MessageBox.Show("Please select Report Month");
            //    }
            //    btnConvert.Enabled = false;
            //    btnCancel.Enabled = false;
            //    btnBrowse.Enabled = true;
            //}
        }
        public static void InsertTransaction(Microsoft.Office.Interop.Excel.Workbook WB, string TranID, string RDate, string Insurance, string Salesby, string Serviceby, string location, string Support, string Rmonth)
        {
            SQLProcs sql = new SQLProcs();
            Microsoft.Office.Interop.Excel.Worksheet wks = (Microsoft.Office.Interop.Excel.Worksheet)WB.Worksheets[1];
            Microsoft.Office.Interop.Excel.Range lastCell = wks.Cells.SpecialCells(Microsoft.Office.Interop.Excel.XlCellType.xlCellTypeLastCell, Type.Missing);
            int lastrow = lastCell.Row; var Terrorism = ""; var Policy_Endorsement = ""; var InsuredType = "";
            for (int i = 2; i <= lastrow; i++)
            {
                string InsuredName = ((Microsoft.Office.Interop.Excel.Range)wks.Cells[i, 30]).Value.Replace("\n", "").TrimStart();
                if (InsuredName == null || InsuredName == "")
                {
                    InsuredName = ((Microsoft.Office.Interop.Excel.Range)wks.Cells[i, 37]).Value.Replace("\n", "").TrimStart();
                }
                string PolicyNo = ((Microsoft.Office.Interop.Excel.Range)wks.Cells[i, 38]).Value.Replace("\n", "").TrimStart();
                if (PolicyNo == null || PolicyNo == "")
                {
                    PolicyNo = ((Microsoft.Office.Interop.Excel.Range)wks.Cells[i, 1]).Value.Replace("\n", "").TrimStart();
                }
                //var Endo_Effective_Date = Convert.ToString(((Microsoft.Office.Interop.Excel.Range)wks.Cells[i, 4]).Value);//.Replace("\n", "").TrimStart();
                //string END_Date = Convert.ToString(((Microsoft.Office.Interop.Excel.Range)wks.Cells[i, 6]).Value);//.Replace("\n", "").TrimStart();
                var Endo_Effective_Date = ((Microsoft.Office.Interop.Excel.Range)wks.Cells[i, 4]).Value;
                var END_Date = ((Microsoft.Office.Interop.Excel.Range)wks.Cells[i, 6]).Value;
                Endo_Effective_Date = Endo_Effective_Date.ToString("dd/MM/yyyy");
                END_Date = END_Date.ToString("dd/MM/yyyy");
                var Premium_Amt = Convert.ToString(((Microsoft.Office.Interop.Excel.Range)wks.Cells[i, 31]).Value);
                var Revenue_Amt = Convert.ToString(((Microsoft.Office.Interop.Excel.Range)wks.Cells[i, 47]).Value);
                var Revenue_Pct = Convert.ToString(((Microsoft.Office.Interop.Excel.Range)wks.Cells[i, 43]).Value);
                var TP_Amt = Convert.ToString(((Microsoft.Office.Interop.Excel.Range)wks.Cells[i, 32]).Value);
                Terrorism = Convert.ToString(((Microsoft.Office.Interop.Excel.Range)wks.Cells[i, 41]).Value);
                var RewardOD_Pct = Convert.ToString(((Microsoft.Office.Interop.Excel.Range)wks.Cells[i, 46]).Value);
                var RewardTP_Pct = Convert.ToString(((Microsoft.Office.Interop.Excel.Range)wks.Cells[i, 46]).Value);
                var IRDARewardAmt = Convert.ToString(((Microsoft.Office.Interop.Excel.Range)wks.Cells[i, 48]).Value);
                var Policy_Type = ((Microsoft.Office.Interop.Excel.Range)wks.Cells[i, 15]).Value.Replace("\n", "").Replace(",", "").TrimStart();
                if (InsuredName.Contains("LIMITED"))
                {
                    InsuredType = "corporate";
                }
                else
                {
                    InsuredType = "Retail";
                }
                if (Premium_Amt == "" || Premium_Amt == " " || Terrorism == null)
                {
                    Premium_Amt = 0;
                }
                if (Revenue_Amt == "" || Revenue_Amt == " " || Terrorism == null)
                {
                    Revenue_Amt = 0;
                }
                if (Terrorism == "" || Terrorism == " " || Terrorism == null)
                {
                    Terrorism = "0";
                }
                if (TP_Amt == "" || TP_Amt == " " || Terrorism == null)
                {
                    TP_Amt = "0";
                }
                if (Revenue_Pct == "" || Revenue_Pct == " " || Terrorism == null)
                {
                    Revenue_Pct = "0";
                }
                if (RewardOD_Pct == "" || RewardOD_Pct == " " || Terrorism == null)
                {
                    RewardOD_Pct = "0";
                }
                if (RewardTP_Pct == "" || RewardTP_Pct == " " || Terrorism == null)
                {
                    RewardTP_Pct = "0";
                }
                if (IRDARewardAmt == "" || IRDARewardAmt == " " || Terrorism == null)
                {
                    IRDARewardAmt = "0";
                }
                sql.ExecuteSQLNonQuery("SP_GoDigitTransactions",
                           new SqlParameter { ParameterName = "@Imode", Value = 1 },
                           new SqlParameter { ParameterName = "@RDate", Value = RDate },
                           new SqlParameter { ParameterName = "@Rmonth", Value = Rmonth },
                           new SqlParameter { ParameterName = "@ClientName", Value = InsuredName },
                           new SqlParameter { ParameterName = "@Itype", Value = InsuredType },
                           new SqlParameter { ParameterName = "@Policy_No", Value = PolicyNo },
                           new SqlParameter { ParameterName = "@Endo_Effective_Date", Value = Endo_Effective_Date },
                           new SqlParameter { ParameterName = "@END_Date", Value = END_Date },
                           new SqlParameter { ParameterName = "@Policy_Type", Value = Policy_Type },
                           new SqlParameter { ParameterName = "@Premium_Amt", Value = Premium_Amt },
                           new SqlParameter { ParameterName = "@TranID", Value = TranID },
                           new SqlParameter { ParameterName = "@Revenue_Amt", Value = Revenue_Amt },
                           new SqlParameter { ParameterName = "@Terrorism", Value = Terrorism },
                           new SqlParameter { ParameterName = "@TP_Amt", Value = TP_Amt },
                           new SqlParameter { ParameterName = "@Revenue_Pct", Value = Revenue_Pct },
                           new SqlParameter { ParameterName = "@RewardOD_Pct", Value = RewardOD_Pct },
                           new SqlParameter { ParameterName = "@RewardTP_Pct", Value = RewardTP_Pct },
                           new SqlParameter { ParameterName = "@IRDARewardAmt", Value = IRDARewardAmt },
                           new SqlParameter { ParameterName = "@Insurance", Value = Insurance },
                           new SqlParameter { ParameterName = "@Salesby", Value = Salesby },
                           new SqlParameter { ParameterName = "@Serviceby", Value = Serviceby },
                           new SqlParameter { ParameterName = "@location", Value = location },
                           new SqlParameter { ParameterName = "@Support", Value = Support },
                           new SqlParameter { ParameterName = "@Policy_Endorsement", Value = Policy_Endorsement },
                           new SqlParameter { ParameterName = "@RFormat", Value = "F1" },
                           new SqlParameter { ParameterName = "@InvNo", Value = "GGIC" },
                           new SqlParameter { ParameterName = "@ReportId", Value = "GGI1" },
                           new SqlParameter { ParameterName = "@DocName", Value = "Godigit General Insurance Co. Ltd." }
                           );
            }
        }
    }
}
