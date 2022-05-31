﻿using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Configuration;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Futurisk
{
    public partial class Home : Form
    {
        private string strconn = ConfigurationManager.ConnectionStrings["IDP"].ToString();
        public Home()
        {
            InitializeComponent();
        }

        private void Home_Load(object sender, EventArgs e)
        {
            bindInsurerdropdown();
        }
        public void bindInsurerdropdown()
        {
            DataRow dr;
            //string com = "select InsurerCode,InsurerName from InsurerMaster where InsurerCode in ('UIIC','NACL') order by InsurerName asc";
            string com = "select InsurerCode,InsurerName from InsurerMaster where InsurerCode != 'NIAC' order by InsurerName asc";
            SqlDataAdapter adpt = new SqlDataAdapter(com, strconn);
            DataTable dt = new DataTable();
            adpt.Fill(dt);
            dr = dt.NewRow();
            dr.ItemArray = new object[] { 0, "Please select" };
            dt.Rows.InsertAt(dr, 0);

            DBInsurer.ValueMember = "InsurerCode";
            DBInsurer.DisplayMember = "InsurerName";
            DBInsurer.DataSource = dt;

        }
        public void bindTypedropdown(string InsurerCode)
        {
            DataRow dr;
            string com = "select ReportCode as ReportCode,ReportCode + ',' +ReportName + '_' + ReportType as Type from ReportsLookUp where InsurerCode = '" + InsurerCode + "' order by ReportName asc";
            SqlDataAdapter adpt = new SqlDataAdapter(com, strconn);
            DataTable dt = new DataTable();
            adpt.Fill(dt);
            dr = dt.NewRow();
            dr.ItemArray = new object[] { 0, "Please select" };
            dt.Rows.InsertAt(dr, 0);

            DBType.ValueMember = "ReportCode";
            DBType.DisplayMember = "Type";
            DBType.DataSource = dt;

            //strconn.Close();
        }

        private void DBInsurer_SelectedIndexChanged(object sender, EventArgs e)
        {
            string InsurerCode = DBInsurer.SelectedValue.ToString();
            bindTypedropdown(InsurerCode);
        }

        private void btnConvert_Click(object sender, EventArgs e)
        {
            string Insurer = DBInsurer.SelectedValue.ToString();
            string Type = DBType.SelectedValue.ToString();
            if (Insurer == "0" && Type == "0")
            {
                MessageBox.Show("Please select Insurer and Type");
            }
            else if (Type == "0")
            {
                MessageBox.Show("Please select Type");
            }
            else
            {
                if(Type == "UIIP") //United India Insurance
                {
                    UnitedTemplate obj = new UnitedTemplate();
                    obj.Show();
                }
                if (Type == "NACP") //National Insurance Co. Ltd.
                {
                    NationalTemplate obj = new NationalTemplate();
                    obj.Show();
                }
                if (Type == "NIAP") //New India Assurance Company
                {
                    NewIndiaTemplate obj = new NewIndiaTemplate();
                    obj.Show();
                }
                if (Type == "OIC1") //Oriental Insurance Company Ltd
                {
                    OrientalTemplate obj = new OrientalTemplate();
                    obj.Show();
                }
                if (Type == "OIC2") //Oriental Insurance Company Ltd
                {
                    OrientalTemplate2 obj = new OrientalTemplate2();
                    obj.Show();
                }
            }
        }

        private void DBType_SelectedIndexChanged(object sender, EventArgs e)
        {
            string Type = DBType.SelectedValue.ToString();
            if (Type != "0")
            {
                btnTemp.Enabled = true;
                btnTemp.ForeColor = System.Drawing.Color.White;
                btnContinue.Enabled = true;
                btnContinue.ForeColor = System.Drawing.Color.White;
            }
            else
            {
                btnTemp.Enabled = false;
                btnTemp.ForeColor = System.Drawing.Color.Black;
                btnContinue.Enabled = false;
                btnContinue.ForeColor = System.Drawing.Color.Black;
            }
        }

        private void btnContinue_Click(object sender, EventArgs e)
        {
            string Type = DBType.SelectedValue.ToString();
            if (Type == "UIIP") //United India Insurance
            {
                UnitedInsurance obj = new UnitedInsurance();
                obj.Show();
                this.Close();
            }
            if (Type == "NACP") //National Insurance Co. Ltd.
            {
                NationalInsurance obj = new NationalInsurance();
                obj.Show();
                this.Close();
            }
            if (Type == "NIAP") //New India Assurance Company
            {
                NewIndiaInsurance obj = new NewIndiaInsurance();
                obj.Show();
                this.Close();
            }
            if (Type == "OIC1") //Oriental Insurance Company Ltd
            {
                OrientalInsurance obj = new OrientalInsurance();
                obj.Show();
                this.Close();
            }
            if (Type == "OIC2") //Oriental Insurance Company Ltd
            {
                OrientalInsurence2 obj = new OrientalInsurence2();
                obj.Show();
                this.Close();
            }
        }

        private void linkLabel1_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            Login obj = new Login();
            obj.Show();
            this.Close();
        }
    }
}
