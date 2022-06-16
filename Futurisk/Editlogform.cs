using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Futurisk
{
    public partial class Editlogform : Form
    {
        public Editlogform()
        {
            InitializeComponent();
        }

        private void Editlogform_Load(object sender, EventArgs e)
        {
            gridReload();
        }
        public void gridReload()
        {
            SQLProcs sql = new SQLProcs();
            DataSet ds = new DataSet();
            ds = sql.SQLExecuteDataset("SP_Login",
                         new SqlParameter { ParameterName = "@Imode", Value = 12 },
                         new SqlParameter { ParameterName = "@BatchID", Value = Fileinfo.BatchId },
                         new SqlParameter { ParameterName = "@ReportCode", Value = Fileinfo.ReportId }
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
                //dataGridView1.Columns[0].DefaultCellStyle.ForeColor = Color.Blue;
                dataGridView1.Columns[0].Width = 180;
                dataGridView1.Columns[1].Width = 50;
                dataGridView1.Columns[2].Width = 40;
                dataGridView1.Columns[5].Width = 150;
                dataGridView1.Columns[6].Width = 120;
            }
        }
    }
}
