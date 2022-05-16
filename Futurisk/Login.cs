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
    public partial class Login : Form
    {
        private string strconn = ConfigurationManager.ConnectionStrings["IDP"].ToString();
        public Login()
        {
            InitializeComponent();
        }

        private void btnLogin_Click(object sender, EventArgs e)
        {
            string UserId = txtUsername.Text;
            string UserPIN = txtPassword.Text;
            if (UserId == "" && UserPIN == "")
            {
                lblErmsg1.Text = "Enter the Username";
                lblErmsg2.Text = "Enter the Password";
            }
            else if (UserId == "")
            {
                lblErmsg1.Text = "Enter the Username";
            }
            else if (UserPIN == "")
            {
                lblErmsg2.Text = "Enter the Password";
            }
            else
            {
                SQLProcs sql = new SQLProcs();
                DataSet ds = new DataSet();
                ds = sql.SQLExecuteDataset("SP_Login",
                             new SqlParameter { ParameterName = "@Imode", Value = 1 },
                             new SqlParameter { ParameterName = "@UserID", Value = UserId },
                             new SqlParameter { ParameterName = "@UserPIN", Value = UserPIN }
                    );
                if (ds != null && ds.Tables.Count > 0 && ds.Tables[0].Rows.Count > 0)
                {
                    string result = ds.Tables[0].Rows[0]["result"].ToString();
                    if (result == "Ok")
                    {
                        Home obj = new Home();
                        obj.Show();
                        this.Hide();
                    }
                    else
                    {
                        MessageBox.Show("Invalid Username and Password");
                    }
                }
            }
        }

        private void txtUsername_TextChanged(object sender, EventArgs e)
        {
            lblErmsg1.Text = "";
        }

        private void txtPassword_TextChanged(object sender, EventArgs e)
        {
            lblErmsg2.Text = "";
        }

        private void btnshowpassword_Click(object sender, EventArgs e)
        {
           // txtPassword_TextChanged.PasswordChar = '\0';
           if (txtPassword.UseSystemPasswordChar == true)
            {
                txtPassword.PasswordChar = '\0';
                txtPassword.UseSystemPasswordChar = false;
            }
            else if (txtPassword.UseSystemPasswordChar == false)
            {
                txtPassword.PasswordChar = '.';
                txtPassword.UseSystemPasswordChar = true;
            }
        }
        private void btnshowpassword_MouseDown(object sender, EventArgs e)
        {
            txtPassword.UseSystemPasswordChar = true;
        }

        private void btnshowpassword_MouseUp(object sender, EventArgs e)
        {
            txtPassword.UseSystemPasswordChar = false;
        }

        private void kryptonButton1_Click(object sender, EventArgs e)
        {
            lblErmsg1.Text = "";
            lblErmsg2.Text = "";
            lblErmsg1.Text = "";
            lblErmsg2.Text = "";
            txtPassword.PasswordChar = '.';
            txtPassword.UseSystemPasswordChar = true;
            txtPassword.Text = "";
            txtUsername.Text = "";
        }
    }
}
