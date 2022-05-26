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
        static string Dir, Filepath, RDate, Result;
        static string Filename, Filenamewithext, Filewithext, name, TranID, BatchID, NoRecord;
        private string strconn = ConfigurationManager.ConnectionStrings["IDP"].ToString();
        DataTable AllNames = new DataTable();
        List<string> Insurancelist = new List<string>();
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
                    ds1 = sql.SQLExecuteDataset("SP_Insert_Transactions",
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
                        string Rmonth = DDMonth.Text.ToString();

                        RDate = DateTime.Now.AddDays(0).ToString("dd-MM-yyyy");
                        Microsoft.Office.Interop.Excel.Application oExcel = new Microsoft.Office.Interop.Excel.Application();
                        Microsoft.Office.Interop.Excel.Workbook WB = oExcel.Workbooks.Open(name);

                        Microsoft.Office.Interop.Excel.Worksheet sheet = WB.ActiveSheet;
                        InsertTransaction(WB, TranID, RDate, Insurance, Salesby, Serviceby, location, Support, Rmonth);
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
                        MessageBox.Show("Given file's BDS is already exists in the database.","Warning!");
                        btnCancel.Enabled = true; btnBrowse.Enabled = true;
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

        private void linkLabel2_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            Fileinfo.Insurer = "UIIC,United India Insurance Co.Ltd.";
            Fileinfo.InsurerCode = "UIIC";
            Fileinfo.ReportId = "UIIP";
            Fileinfo.BatchId = BatchID;
            EditForm obj = new EditForm();
            obj.Show();
            obj.WindowState = FormWindowState.Normal;
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
            string promptValue = ShowDialog("Batch");
            if(promptValue != "")
            {
                Fileinfo.Insurer = "UIIC,United India Insurance Co.Ltd.";
                Fileinfo.InsurerCode = "UIIC";
                Fileinfo.ReportId = "UIIP";
                Fileinfo.BatchId = promptValue.Substring(0, promptValue.IndexOf(","));
                Fileinfo.Filename = promptValue.Substring(promptValue.IndexOf(",") + 1);
                //SQLProcs sql = new SQLProcs();
                //DataSet ds1 = new DataSet();
                //ds1 = sql.SQLExecuteDataset("SP_Insert_Transactions",
                //              new SqlParameter { ParameterName = "@Imode", Value = 8 },
                //              new SqlParameter { ParameterName = "@BatchID", Value = promptValue }
                //     );
                //if (ds1 != null && ds1.Tables.Count > 0 && ds1.Tables[0].Rows.Count > 0)
                //{
                //    Fileinfo.Filename = ds1.Tables[0].Rows[0]["Filename"].ToString();
                //}
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
            Label textLabel = new Label() { Left = 50, Top = 30, Text = "Batch"};
            ComboBox CB = new ComboBox() { Left = 100, Top = 30, Width = 450 };
            Label lblErmsg = new Label() { Left = 50, Top = 50, Width = 300 };
            Button confirmation = new Button() { Text = "Ok", Left = 220, Width = 80, Top = 100, DialogResult = DialogResult.OK,Enabled = false };
            Button confirmation1 = new Button() { Text = "Cancel", Left = 320, Width = 80, Top = 100, DialogResult = DialogResult.Cancel };
            textLabel.Font = new Font("Verdana",11);
            CB.Font = new Font("Verdana", 9);
            lblErmsg.Font = new Font("Verdana", 9);
            lblErmsg.ForeColor = System.Drawing.Color.Red;
            confirmation.Font = new Font("Verdana", 9);
            confirmation1.Font = new Font("Verdana", 9);
            DataRow dr;
            string com = "select distinct(Inv_No) as No,Inv_No+','+[Filename] as Name from BDSMaster where InsurerCode = 'UIIC' and ReportCode = 'UIIP'";
            SqlDataAdapter adpt = new SqlDataAdapter(com, strconn);
            DataTable dt = new DataTable();
            adpt.Fill(dt);
            dr = dt.NewRow();
            dr.ItemArray = new object[] { 0, "" };
            dt.Rows.InsertAt(dr, 0);

            CB.ValueMember = "No";
            CB.DisplayMember = "Name";
            CB.DataSource = dt;

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
            if (prompt.ShowDialog() == DialogResult.OK && CB.SelectedValue.ToString()!= "0")
            {
                // promptValue = CB.SelectedValue.ToString();
                promptValue = CB.Text;
            }
            else if (prompt.ShowDialog() == DialogResult.Cancel)
            {
                promptValue = "";
                prompt.Close();
            }
            //else if (prompt.ShowDialog() == DialogResult.OK && (CB.SelectedValue.ToString() == "0" || CB.SelectedIndex == -1))
            //{
            //    lblErmsg.Text = "Select Batch";
            //}
            return promptValue;
                //return prompt.ShowDialog() == DialogResult.OK ? CB.SelectedValue.ToString() : "0";
        }
        public UnitedInsurance()
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
            AllNames = dt;

           
            //foreach (DataRow row in dt.Rows)
            //{
            //    Insurancelist.Add(row.Field<string>("Description"));
            //}
            //this.DDInsurance.Items.AddRange(Insurancelist.ToArray<string>());
            //DDInsurance.AutoCompleteMode = AutoCompleteMode.SuggestAppend;
            //DDInsurance.AutoCompleteSource = AutoCompleteSource.ListItems;

            //DDInsurance.AutoCompleteMode = AutoCompleteMode.Suggest;
            //DDInsurance.AutoCompleteSource = AutoCompleteSource.CustomSource;
            //AutoCompleteStringCollection combData = new AutoCompleteStringCollection();
            //foreach (DataRow row in dt.Rows)
            //{
            //    combData.Add(row[1].ToString());
            //}
            //DDInsurance.AutoCompleteCustomSource = combData;

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

        private void btnLogout_Click(object sender, EventArgs e)
        {
            Login obj = new Login();
            obj.Show();
            this.Close();
        }

        private void DDInsurance_TextChanged(object sender, EventArgs e)
        {
            //string name = DDInsurance.Text;
            //DataRow[] rows = AllNames.Select(string.Format("Description LIKE '%{0}%'", name));
            //DataTable filteredTable = AllNames.Clone();
            //foreach (DataRow r in rows)
            //    filteredTable.ImportRow(r);
            //DDInsurance.DataSource = null;
            //DDInsurance.DataSource = filteredTable.DefaultView;
            //DDInsurance.DisplayMember = "Description";
            //DDInsurance.Text = name;
            //DDInsurance.DroppedDown = true;

            //string filter_param = DDInsurance.Text;

            ////List<string> filteredItems = Insurancelist.FindAll(x => x.Contains(filter_param));
            //List<string> filteredItems = Insurancelist.FindAll(x => x.ToLower().Contains(filter_param.ToLower()));
            //// another variant for filtering using StartsWith:
            //List<string> filtered = new List<string>();
            //filtered.Add("");
            //filtered.AddRange(filteredItems.ToArray<string>());
            //DDInsurance.DataSource = filteredItems;

            //// if all values removed, bind the original full list again
            //if (String.IsNullOrWhiteSpace(DDInsurance.Text))
            //{
            //    DDInsurance.DataSource = Insurancelist;
            //}
            //DDInsurance.DroppedDown = true;
            //DDInsurance.IntegralHeight = true;
            //DDInsurance.SelectedIndex = -1;
            //DDInsurance.Text = filter_param;
            ////DDInsurance.SelectionStart = filter_param.Length;
            //DDInsurance.SelectionLength = 0;

            //var insurance = DDInsurance.Text;
            //DataRow dr;
            //string com = "select Code,InsurerCode + ','+ UPPER(LEFT(Description, 1)) + LOWER(RIGHT(Description, LEN(Description) - 1)) as Description from tblBRInsurancelkup where Description like '%" + DDInsurance.Text + "%' and GroupBy = 'UN' and Code != '' order by Description asc";
            //SqlDataAdapter adpt = new SqlDataAdapter(com, strconn);
            //DataTable dt = new DataTable();
            //adpt.Fill(dt);
            //dr = dt.NewRow();
            //dr.ItemArray = new object[] { 0, "" };
            //dt.Rows.InsertAt(dr, 0);

            //DDInsurance.ValueMember = "Code";
            //DDInsurance.DisplayMember = "Description";
            //DDInsurance.DataSource = null;
            //DDInsurance.Text = insurance;
            //DDInsurance.DataSource = dt;

            //DDInsurance.DroppedDown = true;
            // IList<string> Values = Insurancelist
            //.Where(x => x.ToString().ToLower().Contains(DDInsurance.Text.ToLower()))
            //.ToList();
            // DDInsurance.Items.Clear();

            // if (DDInsurance.Text != string.Empty)
            //     DDInsurance.Items.AddRange(Values.ToArray<string>());
            // else
            //     DDInsurance.Items.AddRange(Insurancelist.ToArray<string>());
            // DDInsurance.AutoCompleteMode = AutoCompleteMode.Suggest;
            // DDInsurance.AutoCompleteSource = AutoCompleteSource.ListItems;
            // DDInsurance.SelectionStart = DDInsurance.Text.Length;
            // DDInsurance.DroppedDown = true;
        }

        private void DDInsurance_KeyPress(object sender, KeyPressEventArgs e)
        {
            //DataTable table = new DataTable();
            string name = string.Format("{0}{1}", DDInsurance.Text, e.KeyChar.ToString()); //join previous text and new pressed char
            DataRow[] rows = AllNames.Select(string.Format("Description LIKE '%{0}%'", name));
            DataTable filteredTable = AllNames.Clone();
            foreach (DataRow r in rows)
                filteredTable.ImportRow(r);
            DDInsurance.DataSource = null;
            DDInsurance.DataSource = filteredTable.DefaultView;
            DDInsurance.DisplayMember = "Description";
            DDInsurance.Text = name;
            DDInsurance.DroppedDown = true;
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
                    Fileinfo.TName = "UnitedTransaction";
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

                string FileName = "UIIP_" + date + ".xlsx";

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

                    sql.SQLExecuteDataset("SP_Insert_Transactions",
                    new SqlParameter { ParameterName = "@Imode", Value = 6 },
                    new SqlParameter { ParameterName = "@BatchID", Value = BatchID },
                    new SqlParameter { ParameterName = "@version", Value = LoginInfo.version },
                    new SqlParameter { ParameterName = "@UserId", Value = LoginInfo.UserID }
                    );

                    //result = "OK";
                    //lblSuccMsg.Text = "";
                    lblmsg1.ForeColor = System.Drawing.Color.Green;
                    lblmsg1.Text = "              Data downloaded successfully.\n     (File Name:" + FileName + ")";
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
        public static void InsertTransaction(Microsoft.Office.Interop.Excel.Workbook WB, string TranID, string RDate, string Insurance,string Salesby,string Serviceby,string location,string Support,string Rmonth)
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
                            new SqlParameter { ParameterName = "@Rmonth", Value = Rmonth },
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
    }
}
