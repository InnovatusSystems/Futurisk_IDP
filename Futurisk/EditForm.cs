using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Configuration;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Futurisk
{
    public partial class EditForm : Form
    {
        private string strconn = ConfigurationManager.ConnectionStrings["IDP"].ToString();
        public EditForm()
        {
            InitializeComponent();
            lblUser.Text = LoginInfo.UserID;
            lblFilename.Text = Fileinfo.Filename;
            lblInsurer.Text = Fileinfo.Insurer;
            lblReportid.Text = Fileinfo.ReportId;
            lblBatchno.Text = Fileinfo.BatchId;
            TimeUpdater();
            BindDDInsurance();
            BindDDSales();
            BindDDService();
            BindDDLocation();
            BindDDSupport();
            BindDDPolicyType();
            gridload();
            BindFromRSN();
           // BindToRSN();
        }
        public void BindDDInsurance()
        {
            DataRow dr;
            //string com = "select Code,InsurerCode + ','+ UPPER(LEFT(Description, 1)) + LOWER(RIGHT(Description, LEN(Description) - 1)) as Description from tblBRInsurancelkup where InsurerCode = '"+Fileinfo.InsurerCode +"' and Code != '' order by Description asc";
            string com = "select Code,Code +' '+ InsurerCode + ',' + UPPER(LEFT(Description, 1)) + LOWER(RIGHT(Description, LEN(Description) - 1)) as Description from tblBRInsurancelkup where InsurerCode = '" + Fileinfo.InsurerCode + "' order by Description asc";

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
        public void BindDDPolicyType()
        {
            DataRow dr;
            string com = "select Code,Description from tblPolicyTypelkup order by Description asc";
            SqlDataAdapter adpt = new SqlDataAdapter(com, strconn);
            DataTable dt = new DataTable();
            adpt.Fill(dt);
            dr = dt.NewRow();
            dr.ItemArray = new object[] { 0, "" };
            dt.Rows.InsertAt(dr, 0);

            DDPolicyType.ValueMember = "Code";
            DDPolicyType.DisplayMember = "Description";
            DDPolicyType.DataSource = dt;
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
        public void gridload()
        {
            SQLProcs sql = new SQLProcs();
            DataSet ds = new DataSet();
            ds = sql.SQLExecuteDataset("SP_Login",
                         new SqlParameter { ParameterName = "@Imode", Value = 6 },
                         new SqlParameter { ParameterName = "@BatchID", Value = Fileinfo.BatchId }
                );
            if (ds.Tables[0].Rows.Count == 0)
            {
                //dataGridView1.EmptyDataText = "No Records Found";
                var dataTable = new DataTable();
                dataTable.Columns.Add("Message", typeof(string));
                dataTable.Rows.Add("No records found");

                dataGridView1.DataSource = new BindingSource { DataSource = dataTable };
                dataGridView1.Columns[0].AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;
            }
            else
            {

                var bindingSource = new System.Windows.Forms.BindingSource();
                bindingSource.DataSource = ds.Tables[0];

                dataGridView1.DataSource = bindingSource;
                //dataGridView1.DefaultCellStyle.Font = new Font("Bookman Old Style", 8);
                //dataGridView1.Columns[0].Visible = false;
                //dataGridView1.Columns[1].Visible = false;
                dataGridView1.Columns[0].DefaultCellStyle.ForeColor = Color.Blue;
                dataGridView1.Columns[0].Width = 50;
                dataGridView1.Columns[1].Width = 80;
            }
        }

        public void BindFromRSN()
        {
            DataRow dr;
            string com = "SELECT RSN as Num,[RSN] FROM [BDSMaster] where Inv_NO = '" + Fileinfo.BatchId + "' order by RSN asc";
            //string com = "SELECT RSN as Num,[RSN] FROM [BDSMaster] where Inv_NO = 'UIIP1' order by RSN asc";
            SqlDataAdapter adpt = new SqlDataAdapter(com, strconn);
            DataTable dt = new DataTable();
            adpt.Fill(dt);
            dr = dt.NewRow();
            dr.ItemArray = new object[] { 0, null };
            dt.Rows.InsertAt(dr, 0);

            DDFrom.ValueMember = "Num";
            DDFrom.DisplayMember = "RSN";
            DDFrom.DataSource = dt;
        }
        public void BindToRSN()
        {
            DataRow dr;
            string com = "SELECT RSN as Num,[RSN] FROM [BDSMaster] where Inv_NO = '" + Fileinfo.BatchId + "' and RSN >= "+ DDFrom.SelectedValue + " order by RSN asc";
            //string com = "SELECT RSN as Num,[RSN] FROM [BDSMaster] where Inv_NO = 'UIIP1' order by RSN asc";
            SqlDataAdapter adpt = new SqlDataAdapter(com, strconn);
            DataTable dt = new DataTable();
            adpt.Fill(dt);
            dr = dt.NewRow();
            dr.ItemArray = new object[] { 0, null };
            dt.Rows.InsertAt(dr, 0);

            DDTo.ValueMember = "Num";
            DDTo.DisplayMember = "RSN";
            DDTo.DataSource = dt;
        }
        public void gridReload()
        {
            SQLProcs sql = new SQLProcs();
            DataSet ds = new DataSet();
            ds = sql.SQLExecuteDataset("SP_Login",
                         new SqlParameter { ParameterName = "@Imode", Value = 11 },
                         new SqlParameter { ParameterName = "@BatchID", Value = Fileinfo.BatchId },
                         new SqlParameter { ParameterName = "@RSNFrom", Value = DDFrom.SelectedValue.ToString() },
                         new SqlParameter { ParameterName = "@RSNTo", Value = DDTo.SelectedValue.ToString() }
                );
            if (ds.Tables[0].Rows.Count == 0)
            {
                //dataGridView1.EmptyDataText = "No Records Found";
                var dataTable = new DataTable();
                dataTable.Columns.Add("Message", typeof(string));
                dataTable.Rows.Add("No records found");

                dataGridView1.DataSource = new BindingSource { DataSource = dataTable };
                dataGridView1.Columns[0].AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;
            }
            else
            {

                var bindingSource = new System.Windows.Forms.BindingSource();
                bindingSource.DataSource = ds.Tables[0];

                dataGridView1.DataSource = bindingSource;
                //dataGridView1.DefaultCellStyle.Font = new Font("Bookman Old Style", 8);
                //dataGridView1.Columns[0].Visible = false;
                //dataGridView1.Columns[1].Visible = false;
                dataGridView1.Columns[0].DefaultCellStyle.ForeColor = Color.Blue;
                dataGridView1.Columns[0].Width = 50;
                dataGridView1.Columns[1].Width = 80;
            }
        }

        private void DDTo_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (DDTo.SelectedIndex.ToString() != "-1" && DDTo.SelectedIndex.ToString() != "0")
            {
                var RSNFrom = DDFrom.SelectedValue.ToString();
                var RSNTo = DDTo.SelectedValue.ToString();
                if (Convert.ToInt32(RSNFrom) > Convert.ToInt32(RSNTo))
                {
                    MessageBox.Show("ToRSN should be greater than FromRSN");
                    DDTo.SelectedIndex = 0;
                }
            }
            if (DDFrom.SelectedIndex.ToString() != "-1" && DDFrom.SelectedIndex.ToString() != "0" && DDTo.SelectedIndex.ToString() != "-1" && DDTo.SelectedIndex.ToString() != "0")
            {
                btnupdate.Enabled = true;
            }
            else
            {
                btnupdate.Enabled = false;
            }
            if (DDFrom.SelectedIndex.ToString() != "-1" && DDFrom.SelectedIndex.ToString() != "0" && DDTo.SelectedIndex.ToString() != "-1" && DDTo.SelectedIndex.ToString() != "0" && DDFrom.SelectedValue.ToString() != DDTo.SelectedValue.ToString())
            {
                gridReload();
            }
            else
            {
                gridload();
            }
        }

        private void DDFrom_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (DDFrom.SelectedIndex.ToString() != "-1" && DDFrom.SelectedIndex.ToString() != "0")
            {
                DDTo.Enabled = true;
                //if (DDTo.SelectedIndex.ToString() != "-1" && DDTo.SelectedIndex.ToString() != "0")
                //{
                //    var RSNFrom = DDFrom.SelectedValue.ToString();
                //    var RSNTo = DDTo.SelectedValue.ToString();
                //    if (Convert.ToInt32(RSNFrom) > Convert.ToInt32(RSNTo))
                //    {
                //        MessageBox.Show("FromRSN should be less than ToRSN");
                //        DDFrom.SelectedIndex = 0;
                //    }
                //}
                BindToRSN();
                DDTo.SelectedValue = DDFrom.SelectedValue.ToString();
            }
            else
            {
                DDTo.Enabled = false;
            }
            if (DDFrom.SelectedIndex.ToString() != "-1" && DDFrom.SelectedIndex.ToString() != "0" && DDTo.SelectedIndex.ToString() != "-1" && DDTo.SelectedIndex.ToString() != "0")
            {
                btnupdate.Enabled = true;
            }
            else
            {
                btnupdate.Enabled = false;
            }
        }

        private void btnCancel_Click(object sender, EventArgs e)
        {
            DDFrom.SelectedIndex = -1;
            DDTo.SelectedIndex = -1;
            DDTo.Enabled = false;
            DDInsurance.SelectedIndex = 0;
            DDLocation.SelectedIndex = 0;
            DDsales.SelectedIndex = 0;
            DDService.SelectedIndex = 0;
            DDPolicyType.SelectedIndex = 0;
            DDSupport.SelectedIndex = 0;
            RBClient1.Checked = false;
            RBClient2.Checked = false;
            RBType1.Checked = false;
            RBType2.Checked = false;
            gridload();
        }

        private void btnupdate_Click(object sender, EventArgs e)
        {
            var Type = "";var Client = "";
            if(RBType1.Checked == true)
            {
                Type = "Retail";
            }
            else if(RBType2.Checked == true)
            {
                Type = "Corporate";
            }
            if (RBClient1.Checked == true)
            {
                Client = "New Client";
            }
            else if (RBClient1.Checked == true)
            {
                Client = "Existing Client";
            }
            if (DDInsurance.Text.ToString() == "" && DDsales.Text.ToString() == "" && DDService.Text.ToString() == "" && DDLocation.Text.ToString() == "" &&
                DDSupport.Text.ToString() == "" && DDPolicyType.Text.ToString() == "" && Type == "" && Client == "") {
                MessageBox.Show("No data selected.");
            }
            else {
                SQLProcs sql = new SQLProcs();
                sql.ExecuteSQLNonQuery("SP_Login",
                                new SqlParameter { ParameterName = "@Imode", Value = 7 },
                                new SqlParameter { ParameterName = "@RSNFrom", Value = DDFrom.SelectedValue.ToString() },
                                new SqlParameter { ParameterName = "@RSNTo", Value = DDTo.SelectedValue.ToString() },
                                new SqlParameter { ParameterName = "@Insurance", Value = DDInsurance.Text.ToString() },
                                new SqlParameter { ParameterName = "@Salesby", Value = DDsales.Text.ToString() },
                                new SqlParameter { ParameterName = "@Serviceby", Value = DDService.Text.ToString() },
                                new SqlParameter { ParameterName = "@location", Value = DDLocation.Text.ToString() },
                                new SqlParameter { ParameterName = "@Support", Value = DDSupport.Text.ToString() },
                                new SqlParameter { ParameterName = "@PolicyType", Value = DDPolicyType.Text.ToString() },
                                new SqlParameter { ParameterName = "@Type", Value = Type },
                                new SqlParameter { ParameterName = "@NEClient", Value = Client },
                                new SqlParameter { ParameterName = "@BatchID", Value = Fileinfo.BatchId },
                                new SqlParameter { ParameterName = "@InsurerCode", Value = Fileinfo.InsurerCode },
                                new SqlParameter { ParameterName = "@ReportCode", Value = Fileinfo.ReportId },
                                new SqlParameter { ParameterName = "@UserId", Value = LoginInfo.UserID },
                                new SqlParameter { ParameterName = "@TName", Value = Fileinfo.TName }
                                );

                MessageBox.Show("Updated Sucessfully.");
                //DDFrom.SelectedIndex = -1;
                //DDTo.SelectedIndex = -1;
                //DDTo.Enabled = false;
                //DDInsurance.SelectedIndex = 0;
                //DDLocation.SelectedIndex = 0;
                //DDsales.SelectedIndex = 0;
                //DDService.SelectedIndex = 0;
                //DDSpeciality.SelectedIndex = 0;
                //DDSupport.SelectedIndex = 0;
                //RBType1.Checked = false;
                //RBType2.Checked = false;
                gridReload();
            }
        }

        private void kryptonButton1_Click(object sender, EventArgs e)
        {
            ExcelExport();
        }
        public void ExcelExport()
        {
            try
            {

                string pathUser = Environment.GetFolderPath(Environment.SpecialFolder.UserProfile);
                string pathDownload = Path.Combine(pathUser, "Downloads\\");

                SQLProcs sql = new SQLProcs();
                DataSet ResultsTable = new DataSet();

                ResultsTable = sql.SQLExecuteDataset("SP_Login",
               new SqlParameter { ParameterName = "@Imode", Value = 9 },
               new SqlParameter { ParameterName = "@BatchID", Value = Fileinfo.BatchId }
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
                    string path = pathDownload + "\\" + FileName;
                    wb.SaveAs(path);

                    MessageBox.Show("SmartRead data downloaded as XLSX file for your verification.\n     (File Name:" + FileName + ")");
                    //result = "OK";
                    //lblSuccMsg.Text = "";
                    //lblmsg1.ForeColor = System.Drawing.Color.Green;
                    //lblmsg1.Text = "              Data downloaded successfully.\n     (File Name:" + FileName + ")";
                    
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Data export failed.");
            }
        }

        private void kryptonButton2_Click(object sender, EventArgs e)
        {
            Editlogform obj = new Editlogform();
            obj.Show();
        }
        async void TimeUpdater()
        {
            while (true)
            {
                lblTimer.Text = DateTime.Now.ToString("dd-MM-yyyy hh:mm:ss tt");
                await Task.Delay(1000);
            }
        }
    }
}
