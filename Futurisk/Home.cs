using System;
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
            lblUser.Text = LoginInfo.UserID;
            bindInsurerdropdown();
        }
        public void bindInsurerdropdown()
        {
            DataRow dr;
            //string com = "select InsurerCode,InsurerName from InsurerMaster where InsurerCode in ('UIIC','NACL') order by InsurerName asc";
            string com = "select InsurerCode,InsurerCode + ',' + InsurerName as InsurerName from InsurerMaster order by InsurerName asc";
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
                MessageBox.Show("Please select Insurer and Document Type");
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
                if (Type == "GGI1") //Godigit General Insurance Co. Ltd
                {
                    Godigitsample1 obj = new Godigitsample1();
                    obj.Show();
                }
                if (Type == "TAI1") //TATA AIG General Insurance Co. Ltd.
                {
                    TATASample1 obj = new TATASample1();
                    obj.Show();
                }
                if (Type == "ILG1") //ICICI Lombard General Insurance Co. Ltd.
                {
                    ICICIsample1 obj = new ICICIsample1();
                    obj.Show();
                }
                if (Type == "STA1") //Star Health & Allied Insurance Co. Ltd.
                {
                    StarHealthSample1 obj = new StarHealthSample1();
                    obj.Show();
                }
                if (Type == "NACX") //National Insurance Co. Ltd.
                {
                    NationalExcelSample obj = new NationalExcelSample();
                    obj.Show();
                }
                if (Type == "BAGX") //Bajajational Insurance Co. Ltd.
                {
                    BajajSample1 obj = new BajajSample1();
                    obj.Show();
                }
                if (Type == "CHIX") //Care Health Insurance Co. Ltd.
                {
                    CareHealthSample obj = new CareHealthSample();
                    obj.Show();
                }
                if (Type == "CMGX") //Chola MS General Insurance Co. Ltd.
                {
                    CholaSample obj = new CholaSample();
                    obj.Show();
                }
                if (Type == "ABHX") //Aditya Birla Health Insurance Co. Ltd.
                {
                    AdityaTemplate1 obj = new AdityaTemplate1();
                    obj.Show();
                }
                if (Type == "EGIX") //Edelweiss General Insurance Co. Ltd.
                {
                    EdelweissSample obj = new EdelweissSample();
                    obj.Show();
                }
                if (Type == "HEGX") //HDFC Ergo General Insurance Co. Ltd.
                {
                    HDFCSample obj = new HDFCSample();
                    obj.Show();
                }
                if (Type == "ITGX") //Iffco Tokio General Insurance Co.Ltd.
                {
                    IFFCOSample obj = new IFFCOSample();
                    obj.Show();
                }
                if (Type == "LVGX") //Liberty Videocon General Insurance Co. Ltd.
                {
                    LibertySample obj = new LibertySample();
                    obj.Show();
                }
                if (Type == "RQGX") //Raheja QBE General Insurance Co. Ltd.
                {
                    RahejaQBESample obj = new RahejaQBESample();
                    obj.Show();
                }
                if (Type == "RGIX") //Reliance General Insurance Co. Ltd.
                {
                    RelianceSample obj = new RelianceSample();
                    obj.Show();
                }
                if (Type == "RSGX") //Royal Sundaram General Insurance Co Ltd.
                {
                    RoyalExcelSample obj = new RoyalExcelSample();
                    obj.Show();
                }
                if (Type == "SBIX") //SBI General Insurance Co. Ltd.
                {
                    SBISample obj = new SBISample();
                    obj.Show();
                }
            }
        }

        private void DBType_SelectedIndexChanged(object sender, EventArgs e)
        {
            string Type = DBType.SelectedValue.ToString();
            if (Type != "0")
            {
                //if (Type == "STA1")
                //{
                //    lbltype.Visible = true;
                //    RBType1.Visible = true;
                //    RBType2.Visible = true;
                //}
                //else if (Type == "ILG1")
                //{
                //    lblType2.Visible = true;
                //    RBType3.Visible = true;
                //    RBType4.Visible = true;
                //}
                //else
                //{
                //    lbltype.Visible = false;
                //    RBType1.Visible = false;
                //    RBType2.Visible = false;
                //    lblType2.Visible = false;
                //    RBType3.Visible = false;
                //    RBType4.Visible = false;
                //}
                btnTemp.Enabled = true;
                btnTemp.ForeColor = System.Drawing.Color.White;
                btnContinue.Enabled = true;
                btnContinue.ForeColor = System.Drawing.Color.White;
            }
            else
            {
                lbltype.Visible = false;
                RBType1.Visible = false;
                RBType2.Visible = false;
                lblType2.Visible = false;
                RBType3.Visible = false;
                RBType4.Visible = false;
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
            if (Type == "GGI1") //Godigit General Insurance Co. Ltd
            {
                Fileinfo.InsurerCode = "GGIC";
                Fileinfo.ReportId = Type;
                GodigitInsurence obj = new GodigitInsurence();
                obj.Show();
                this.Close();
            }
            if (Type == "TAI1") //TATA AIG General Insurance Co. Ltd.
            {
                Fileinfo.InsurerCode = "TAIG";
                Fileinfo.ReportId = Type;
                GodigitInsurence obj = new GodigitInsurence();
                obj.Show();
                this.Close();
            }
            if (Type == "ILG1") //ICICI Lombard General Insurance Co. Ltd.
            {
                Fileinfo.InsurerCode = "ILGI";
                Fileinfo.ReportId = Type;
                GodigitInsurence obj = new GodigitInsurence();
                obj.Show();
                this.Close();
            }
            if (Type == "STA1") //Star Health & Allied Insurance Co. Ltd.
            {
                Fileinfo.InsurerCode = "STAR";
                Fileinfo.ReportId = Type;
                GodigitInsurence obj = new GodigitInsurence();
                obj.Show();
                this.Close();
            }
            if (Type == "NACX") //National Insurance Co. Ltd.
            {
                Fileinfo.InsurerCode = "NACL";
                Fileinfo.ReportId = Type;
                GodigitInsurence obj = new GodigitInsurence();
                obj.Show();
                this.Close();
            }
            if (Type == "BAGX") //National Insurance Co. Ltd.
            {
                Fileinfo.InsurerCode = "BAGI";
                Fileinfo.ReportId = Type;
                GodigitInsurence obj = new GodigitInsurence();
                obj.Show();
                this.Close();
            }
            if (Type == "CHIX") //Care Health Insurance Co. Ltd.
            {
                Fileinfo.InsurerCode = "CHIC";
                Fileinfo.ReportId = Type;
                GodigitInsurence obj = new GodigitInsurence();
                obj.Show();
                this.Close();
            }
            if (Type == "CMGX") //Chola MS General Insurance Co. Ltd.
            {
                Fileinfo.InsurerCode = "CMGI";
                Fileinfo.ReportId = Type;
                GodigitInsurence obj = new GodigitInsurence();
                obj.Show();
                this.Close();
            }
            if (Type == "ABHX") //Aditya Birla Health Insurance Co. Ltd.
            {
                Fileinfo.InsurerCode = "ABHI";
                Fileinfo.ReportId = Type;
                GodigitInsurence obj = new GodigitInsurence();
                obj.Show();
                this.Close();
            }
            if (Type == "EGIX") //Edelweiss General Insurance Co. Ltd.
            {
                Fileinfo.InsurerCode = "EGIC";
                Fileinfo.ReportId = Type;
                GodigitInsurence obj = new GodigitInsurence();
                obj.Show();
                this.Close();
            }
            if (Type == "HEGX") //HDFC Ergo General Insurance Co. Ltd.
            {
                Fileinfo.InsurerCode = "HEGI";
                Fileinfo.ReportId = Type;
                GodigitInsurence obj = new GodigitInsurence();
                obj.Show();
                this.Close();
            }
            if (Type == "ITGX") //Iffco Tokio General Insurance Co.Ltd.
            {
                Fileinfo.InsurerCode = "ITGI";
                Fileinfo.ReportId = Type;
                GodigitInsurence obj = new GodigitInsurence();
                obj.Show();
                this.Close();
            }
            if (Type == "LVGX") //Liberty Videocon General Insurance Co. Ltd.
            {
                Fileinfo.InsurerCode = "LVGI";
                Fileinfo.ReportId = Type;
                GodigitInsurence obj = new GodigitInsurence();
                obj.Show();
                this.Close();
            }
            if (Type == "RQGX") //Raheja QBE General Insurance Co. Ltd.
            {
                Fileinfo.InsurerCode = "RQGI";
                Fileinfo.ReportId = Type;
                GodigitInsurence obj = new GodigitInsurence();
                obj.Show();
                this.Close();
            }
            if (Type == "RGIX") //Reliance General Insurance Co. Ltd.
            {
                Fileinfo.InsurerCode = "RGIC";
                Fileinfo.ReportId = Type;
                GodigitInsurence obj = new GodigitInsurence();
                obj.Show();
                this.Close();
            }
            if (Type == "SBIX") //SBI General Insurance Co. Ltd.
            {
                Fileinfo.InsurerCode = "SBII";
                Fileinfo.ReportId = Type;
                GodigitInsurence obj = new GodigitInsurence();
                obj.Show();
                this.Close();
            }
            if (Type == "RSGX") //Royal Sundaram General Insurance Co Ltd.
            {
                Fileinfo.InsurerCode = "RSGI";
                Fileinfo.ReportId = Type;
                GodigitInsurence obj = new GodigitInsurence();
                obj.Show();
                this.Close();
            }
            if (Type == "UIIX") // United India Insurance Co Ltd (Excel).
            {
                Fileinfo.InsurerCode = "UIIC";
                Fileinfo.ReportId = Type;
                AckoInsurence obj = new AckoInsurence();
                obj.Show();
                this.Close();
            }
            if (Type == "NIX1") // New India Assurance Company Ltd.(Excel).
            {
                Fileinfo.InsurerCode = "NIAC";
                Fileinfo.ReportId = Type;
                AckoInsurence obj = new AckoInsurence();
                obj.Show();
                this.Close();
            }
            if (Type == "NIX2") // New India Assurance Company Ltd.(Excel).
            {
                Fileinfo.InsurerCode = "NIAC";
                Fileinfo.ReportId = Type;
                AckoInsurence obj = new AckoInsurence();
                obj.Show();
                this.Close();
            }
            if (Type == "NIX3") // New India Assurance Company Ltd.(Excel).
            {
                Fileinfo.InsurerCode = "NIAC";
                Fileinfo.ReportId = Type;
                AckoInsurence obj = new AckoInsurence();
                obj.Show();
                this.Close();
            }
            if (Type == "FGIX") //Future General India Insurance Co. Ltd..
            {
                Fileinfo.InsurerCode = "FGII";
                Fileinfo.ReportId = Type;
                AckoInsurence obj = new AckoInsurence();
                obj.Show();
                this.Close();
            }
            if (Type == "KMGX") //Kotak Mahindra General Insurance Co. Ltd.
            {
                Fileinfo.InsurerCode = "KMGI";
                Fileinfo.ReportId = Type;
                AckoInsurence obj = new AckoInsurence();
                obj.Show();
                this.Close();
            }
            if (Type == "ACKX") //Acko General Insurance  Co.Ltd.
            {
                Fileinfo.InsurerCode = "ACKO";
                Fileinfo.ReportId = Type;
                AckoInsurence obj = new AckoInsurence();
                obj.Show();
                this.Close();
            }
            if (Type == "NBHX") //Acko General Insurance  Co.Ltd.
            {
                Fileinfo.InsurerCode = "NBHI";
                Fileinfo.ReportId = Type;
                AckoInsurence obj = new AckoInsurence();
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
