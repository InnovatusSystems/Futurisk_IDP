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
using Smartreader_DLL;
using System.Windows.Forms;
using Spire.Xls;
using System.Data.SqlClient;
using System.Configuration;
using Syncfusion.XlsIO;
using System.Globalization;
using System.Text.RegularExpressions;

namespace Futurisk
{
    public partial class GodigitInsurence : Form
    {
        static string Dir, Filepath, RDate, Result, Fileextn, DelFile, ProcName;
        static string Filename, Filenamewithext, Filewithext, name, TranID, BatchID, NoRecord;

        private void btnConvert_Click(object sender, EventArgs e)
        {
            try
            {
                btnCancel.Enabled = false;
                var confirmResult = MessageBox.Show("Excel file to be uploaded must meet the guidelines below:\n" +
                    "\nExcel file should be in XLS or XLSX format.Any other file extension (XLT,XLM) will be REJECTED.\n" +
                    "The excel file may be rejected incase of Column name/address mismatch with this ViewSample format.\n" +
                    "Excel file should not contain any empty rows / columns.\n" +
                    "Extracting data should be on the first sheet of the Excel.\n" +
                    "\nHave you got the right document?", "Confirm",
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
                    RBType1.Enabled = false;
                    RBType2.Enabled = false;
                    RBType3.Enabled = false;
                    RBType4.Enabled = false;

                    DDInsurance.SelectionLength = 0;
                    DDLocation.SelectionLength = 0;
                    DDsales.SelectionLength = 0;
                    DDService.SelectionLength = 0;
                    DDSupport.SelectionLength = 0;
                    DDMonth.SelectionLength = 0;

                    if (Fileinfo.InsurerCode == "GGIC")
                    {
                        ProcName = "SP_GoDigitTransactions";
                    }
                    else if (Fileinfo.InsurerCode == "TAIG")
                    {
                        ProcName = "SP_TATATransactions";
                    }
                    else if (Fileinfo.InsurerCode == "STAR")
                    {
                        ProcName = "SP_StarHealthTransactions";
                    }
                    else if (Fileinfo.InsurerCode == "ILGI")
                    {
                        ProcName = "SP_ICICITransactions";
                    }
                    else if (Fileinfo.InsurerCode == "NACL")
                    {
                        ProcName = "SP_NationalExcel_Transactions";
                    }
                    else if (Fileinfo.InsurerCode == "BAGI")
                    {
                        ProcName = "SP_BajajTransactions";
                    }
                    else if (Fileinfo.InsurerCode == "CHIC")
                    {
                        ProcName = "SP_CareHealthTransactions";
                    }
                    else if (Fileinfo.InsurerCode == "CMGI")
                    {
                        ProcName = "SP_CholaTransactions";
                    }
                    else if (Fileinfo.InsurerCode == "ABHI")
                    {
                        ProcName = "SP_AdityaBirlaTransactions";
                    }
                    else if (Fileinfo.InsurerCode == "EGIC")
                    {
                        ProcName = "SP_EdelweissTransactions";
                    }
                    else if (Fileinfo.InsurerCode == "HEGI")
                    {
                        ProcName = "SP_HDFCTransactions";
                    }
                    else if (Fileinfo.InsurerCode == "ITGI")
                    {
                        ProcName = "SP_IffcoTransactions";
                    }
                    else if (Fileinfo.InsurerCode == "LVGI")
                    {
                        ProcName = "SP_LibertyTransactions";
                    }
                    else if (Fileinfo.InsurerCode == "RQGI")
                    {
                        ProcName = "SP_RahejaTransactions";
                    }
                    else if (Fileinfo.InsurerCode == "RGIC")
                    {
                        ProcName = "SP_RelianceTransactions";
                    }
                    else if (Fileinfo.InsurerCode == "SBII")
                    {
                        ProcName = "SP_SBITransactions";
                    }
                    else if (Fileinfo.InsurerCode == "RSGI")
                    {
                        ProcName = "SP_RoyalTransactions";
                    }
                    SQLProcs sql = new SQLProcs();
                    DataSet ds1 = new DataSet();
                    ds1 = sql.SQLExecuteDataset(ProcName,
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
                        //if (Fileextn == ".xls")
                        //{
                        //    name = Dir + "\\" + Filename + DateTime.Now.AddDays(0).ToString("ddMMyyyyhhmmss") + ".xlsx";
                        //    DelFile = name;
                        //    Workbook workbook = new Workbook();
                        //    workbook.LoadFromFile(Filepath);
                        //    workbook.SaveToFile(name, Spire.Xls.ExcelVersion.Version2013);
                        //}
                        //else
                        //{
                            name = Filepath;
                        //}


                        string filewithourext = Filename + DateTime.Now.AddDays(0).ToString("ddMMyyyyhhmmss");
                        filewithourext = Dir + "\\" + filewithourext;

                        DataSet ds = new DataSet();
                        ds = sql.SQLExecuteDataset(ProcName,
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


                        string Insurance = DDInsurance.Text.ToString();
                        string Salesby = DDsales.Text.ToString();
                        string Serviceby = DDService.Text.ToString();
                        string location = DDLocation.Text.ToString();
                        string Support = DDSupport.Text.ToString();
                        string Rmonth = DDMonth.Text.ToString();
                        if (Fileinfo.ReportId == "GGI1")
                        {
                            //InsertTransaction(WB, TranID, RDate, Insurance, Salesby, Serviceby, location, Support, Rmonth);
                            //ExceltoDatatable(TranID, RDate, Insurance, Salesby, Serviceby, location, Support, Rmonth, Filepath);
                            GodigitInsurence1.ExceltoDatatable(TranID, RDate, Insurance, Salesby, Serviceby, location, Support, Rmonth, Filepath, strconn);
                        }
                        else if (Fileinfo.ReportId == "ILG1" && Fileinfo.Type == "General")
                        {
                                //InsertTransaction5(WB, TranID, RDate, Insurance, Salesby, Serviceby, location, Support, Rmonth);
                                //ExceltoDatatable1(TranID, RDate, Insurance, Salesby, Serviceby, location, Support, Rmonth);
                            ICICITransaction1.ExceltoDatatable1(TranID, RDate, Insurance, Salesby, Serviceby, location, Support, Rmonth, Filepath, strconn);
                        }
                        else
                        {
                            RDate = DateTime.Now.AddDays(0).ToString("dd-MM-yyyy");
                            Microsoft.Office.Interop.Excel.Application oExcel = new Microsoft.Office.Interop.Excel.Application();
                            Microsoft.Office.Interop.Excel.Workbook WB = oExcel.Workbooks.Open(name);

                            Microsoft.Office.Interop.Excel.Worksheet sheet = WB.ActiveSheet;
                            if (Fileinfo.ReportId == "TAI1")
                            {
                                //InsertTransaction2(WB, TranID, RDate, Insurance, Salesby, Serviceby, location, Support, Rmonth);
                                TATAInsurence.InsertTransaction(WB, TranID, RDate, Insurance, Salesby, Serviceby, location, Support, Rmonth, strconn);

                            }
                            else if (Fileinfo.ReportId == "ILG1" && Fileinfo.Type == "Terrorism")
                            {
                                // InsertTransaction5(WB, TranID, RDate, Insurance, Salesby, Serviceby, location, Support, Rmonth);
                                ICICITransaction1.InsertTransaction(WB, LoginInfo.UserID, strconn);
                            }
                            else if (Fileinfo.ReportId == "NACX")
                            {
                                //InsertTransaction(WB, TranID, RDate, Insurance, Salesby, Serviceby, location, Support, Rmonth);
                                NationalInsurence.InsertExcelTransaction(WB, TranID, RDate, Insurance, Salesby, Serviceby, location, Support, Rmonth, strconn);
                            }
                            else if (Fileinfo.ReportId == "BAGX")
                            {
                                //InsertTransaction(WB, TranID, RDate, Insurance, Salesby, Serviceby, location, Support, Rmonth);
                                BajajInsurance.InsertTransaction(WB, TranID, RDate, Insurance, Salesby, Serviceby, location, Support, Rmonth, strconn);
                            }
                            else if (Fileinfo.ReportId == "CHIX")
                            {
                                //InsertTransaction(WB, TranID, RDate, Insurance, Salesby, Serviceby, location, Support, Rmonth);
                                CareHealthInsurance.InsertTransaction(WB, TranID, RDate, Insurance, Salesby, Serviceby, location, Support, Rmonth, strconn);
                            }
                            else if (Fileinfo.ReportId == "CMGX")
                            {
                                //InsertTransaction(WB, TranID, RDate, Insurance, Salesby, Serviceby, location, Support, Rmonth);
                                CholaInsurance.InsertTransaction(WB, TranID, RDate, Insurance, Salesby, Serviceby, location, Support, Rmonth, strconn);
                            }
                            else if (Fileinfo.ReportId == "ABHX")
                            {
                                //InsertTransaction(WB, TranID, RDate, Insurance, Salesby, Serviceby, location, Support, Rmonth);
                                AdityaInsurance.InsertTransaction(WB, TranID, RDate, Insurance, Salesby, Serviceby, location, Support, Rmonth, strconn);
                            }
                            else if (Fileinfo.ReportId == "EGIX")
                            {
                                //InsertTransaction(WB, TranID, RDate, Insurance, Salesby, Serviceby, location, Support, Rmonth);
                                EdelweissInsurance.InsertTransaction(WB, TranID, RDate, Insurance, Salesby, Serviceby, location, Support, Rmonth, strconn);
                            }
                            else if (Fileinfo.ReportId == "HEGX")
                            {
                                //InsertTransaction(WB, TranID, RDate, Insurance, Salesby, Serviceby, location, Support, Rmonth);
                                HDFCInsurance.InsertTransaction(WB, TranID, RDate, Insurance, Salesby, Serviceby, location, Support, Rmonth, strconn);
                            }
                            else if (Fileinfo.ReportId == "ITGX")
                            {
                                //InsertTransaction(WB, TranID, RDate, Insurance, Salesby, Serviceby, location, Support, Rmonth);
                                IffcoTransactions.InsertTransaction(WB, TranID, RDate, Insurance, Salesby, Serviceby, location, Support, Rmonth, strconn);
                            }
                            else if (Fileinfo.ReportId == "LVGX")
                            {
                                //InsertTransaction(WB, TranID, RDate, Insurance, Salesby, Serviceby, location, Support, Rmonth);
                                LibertyInsurance.InsertTransaction(WB, TranID, RDate, Insurance, Salesby, Serviceby, location, Support, Rmonth, strconn);
                            }
                            else if (Fileinfo.ReportId == "RQGX")
                            {
                                //InsertTransaction(WB, TranID, RDate, Insurance, Salesby, Serviceby, location, Support, Rmonth);
                                RahejaTransactions.InsertTransaction(WB, TranID, RDate, Insurance, Salesby, Serviceby, location, Support, Rmonth, strconn);
                            }
                            else if (Fileinfo.ReportId == "RGIX")
                            {
                                //InsertTransaction(WB, TranID, RDate, Insurance, Salesby, Serviceby, location, Support, Rmonth);
                                RelianceInsurance.InsertTransaction(WB, TranID, RDate, Insurance, Salesby, Serviceby, location, Support, Rmonth, strconn);
                            }
                            else if (Fileinfo.ReportId == "RSGX")
                            {
                                //InsertTransaction(WB, TranID, RDate, Insurance, Salesby, Serviceby, location, Support, Rmonth);
                                RoyalInsurance.InsertTransaction(WB, TranID, RDate, Insurance, Salesby, Serviceby, location, Support, Rmonth, strconn);
                            }
                            else if (Fileinfo.ReportId == "SBIX")
                            {
                                //InsertTransaction(WB, TranID, RDate, Insurance, Salesby, Serviceby, location, Support, Rmonth);
                                SBIInsurance.InsertTransaction(WB, TranID, RDate, Insurance, Salesby, Serviceby, location, Support, Rmonth, strconn);
                            }
                            else if (Fileinfo.ReportId == "STA1")
                            {
                                if (Fileinfo.Type == "Corporate")
                                {
                                    //InsertTransaction3(WB, TranID, RDate, Insurance, Salesby, Serviceby, location, Support, Rmonth);
                                    StarHealthInsurence1.InsertTransaction(WB, TranID, RDate, Insurance, Salesby, Serviceby, location, Support, Rmonth, strconn);
                                }
                                else if (Fileinfo.Type == "Retail")
                                {
                                    //InsertTransaction4(WB, TranID, RDate, Insurance, Salesby, Serviceby, location, Support, Rmonth);
                                    StarHealthInsurence1.InsertRetailTransaction(WB, TranID, RDate, Insurance, Salesby, Serviceby, location, Support, Rmonth, strconn);
                                }
                            }

                            oExcel.Workbooks.Close();
                        }
                        GetBachid_NoRecord(ProcName, TranID);


                        DateTime b = DateTime.Now;
                        TimeSpan diff = b - a;
                        var Sec = String.Format("{0}", diff.Seconds);
                        lblmsg.Text = "";

                        DDInsurance.SelectionLength = 0;
                        DDLocation.SelectionLength = 0;
                        DDsales.SelectionLength = 0;
                        DDService.SelectionLength = 0;
                        DDSupport.SelectionLength = 0;
                        DDMonth.SelectionLength = 0;

                        if (Fileinfo.ReportId == "ILG1" && Fileinfo.Type == "Terrorism")
                        {
                            lblSuccMsg.Text = "          SmartRead Done in " + Sec + " Seconds.";
                        }
                        else
                        {
                            lblSuccMsg.Text = "                SmartRead Done in " + Sec + " Seconds.\n" +
                                             "           Batch ID: " + BatchID + " ,Number of records: " + NoRecord;

                            var confirmExportResult = MessageBox.Show("Data is now in database. Do you wish to get it in Excel format for your checking?", "Confirm",
                                            MessageBoxButtons.YesNo);
                            if (confirmExportResult == DialogResult.Yes)
                            {
                                ExcelExport(ProcName);
                            }
                            else
                            {
                                lblmsg1.ForeColor = System.Drawing.Color.DarkGreen;
                                lblmsg1.Text = "You can check the data through another Menu Option.";
                            }
                            linkLabel2.Enabled = true;
                            linkLabel2.Text = "Click here to Edit the records if needed.";
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
                lblmsg1.Text = "                        SmartRead data extraction failed." +
                               "\n Possible reasons:Wrong document,Column mismatch or empty rows" +
                               "\n                       Please check the source file.";
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
                RBType1.Enabled = true;
                RBType2.Enabled = true;
                RBType3.Enabled = true;
                RBType4.Enabled = true;
                DDInsurance.SelectionLength = 0;
                DDLocation.SelectionLength = 0;
                DDsales.SelectionLength = 0;
                DDService.SelectionLength = 0;
                DDSupport.SelectionLength = 0;
                DDMonth.SelectionLength = 0;
            }
        }
       public void GetBachid_NoRecord(string ProcName, string TranID)
        {
            try
            {
                SQLProcs sql = new SQLProcs();
                DataSet ds = new DataSet();

                ds = sql.SQLExecuteDataset(ProcName,
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
                dsR = sql.SQLExecuteDataset(ProcName,
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
        public void ExcelExport(string Pname)
        {
            try
            {
                string pathUser = Environment.GetFolderPath(Environment.SpecialFolder.UserProfile);
                string pathDownload = Path.Combine(pathUser, "Downloads\\");

                SQLProcs sql = new SQLProcs();
                DataSet ResultsTable = new DataSet();

                ResultsTable = sql.SQLExecuteDataset(Pname,
               new SqlParameter { ParameterName = "@Imode", Value = 3 },
               new SqlParameter { ParameterName = "@TranID", Value = TranID }
               );

                string date = DateTime.Now.ToString();
                date = date.Replace("/", "_").Replace(":", "").Replace(" ", "").Replace("AM", "").Replace("PM", "");

                string FileName = Fileinfo.ReportId + "_" + date + ".xlsx";

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

                    sql.SQLExecuteDataset(Pname,
                    new SqlParameter { ParameterName = "@Imode", Value = 6 },
                    new SqlParameter { ParameterName = "@BatchID", Value = BatchID },
                    new SqlParameter { ParameterName = "@version", Value = LoginInfo.version },
                    new SqlParameter { ParameterName = "@UserId", Value = LoginInfo.UserID }
                    );

                    lblmsg1.ForeColor = System.Drawing.Color.Green;
                    lblmsg1.Text = "SmartRead data downloaded as XLSX file for your verification." +
                                    "\n            (File Name:" + FileName + ")";

                }
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

        private void btnEdit_Click(object sender, EventArgs e)
        {
            string promptValue = ShowDialog("Batch");
            if (promptValue != "")
            {
                if (Fileinfo.InsurerCode == "GGIC")
                {
                    Fileinfo.Insurer = "GGIC,Godigit General Insurance Co. Ltd.";
                }
                else if (Fileinfo.InsurerCode == "TAIG")
                {
                    Fileinfo.Insurer = "TAIG,TATA AIG General Insurance Co. Ltd.";
                }
                else if (Fileinfo.InsurerCode == "STAR")
                {
                    Fileinfo.Insurer = "STAR,Star Health and Allied Insurance Co.Ltd.";
                }
                else if (Fileinfo.InsurerCode == "ILGI")
                {
                    Fileinfo.Insurer = "ILGI,ICICI Lombard General Insurance Co. Ltd.";
                }
                else if (Fileinfo.InsurerCode == "NACL")
                {
                    Fileinfo.Insurer = "NACL,National Insurance Co. Ltd.";
                }
                else if (Fileinfo.InsurerCode == "BAGI")
                {
                    Fileinfo.Insurer = "BAGI,Bajaj Allianz General Insurance Co. Ltd.";
                }
                else if (Fileinfo.InsurerCode == "CHIC")
                {
                    Fileinfo.Insurer = "CHIC,Care Health Insurance Co. Ltd.";
                }
                else if (Fileinfo.InsurerCode == "CMGI")
                {
                    Fileinfo.Insurer = "CMGI,Chola MS General Insurance Co. Ltd.";
                }
                else if (Fileinfo.InsurerCode == "ABHI")
                {
                    Fileinfo.Insurer = "ABHI,Aditya Birla Health Insurance Co.Ltd.";
                }
                else if (Fileinfo.InsurerCode == "EGIC")
                {
                    Fileinfo.Insurer = "EGIC,Edelweiss General Insurance Co.Ltd.";
                }
                else if (Fileinfo.InsurerCode == "HEGI")
                {
                    Fileinfo.Insurer = "HEGI,HDFC Ergo General Insurance Co. Ltd.";
                }
                else if (Fileinfo.InsurerCode == "ITGI")
                {
                    Fileinfo.Insurer = "ITGI,Iffco Tokio General Insurance Co.Ltd.";
                }
                else if (Fileinfo.InsurerCode == "LVGI")
                {
                    Fileinfo.Insurer = "LVGI,Liberty Videocon General Insurance Co.Ltd.";
                }
                else if (Fileinfo.InsurerCode == "RQGI")
                {
                    Fileinfo.Insurer = "RQGI,Raheja QBE General Insurance Co. Ltd.";
                }
                else if (Fileinfo.InsurerCode == "RGIC")
                {
                    Fileinfo.Insurer = "RGIC,Reliance General Insurance Co. Ltd.";
                }
                else if (Fileinfo.InsurerCode == "SBII")
                {
                    Fileinfo.Insurer = "SBII,SBI General Insurance Co. Ltd.";
                }
                else if (Fileinfo.InsurerCode == "RSGI")
                {
                    Fileinfo.Insurer = "RSGI,Royal Sundaram General Insurance Co Ltd.";
                }

                Fileinfo.BatchId = promptValue.Substring(0, promptValue.IndexOf(","));
                Fileinfo.Filename = promptValue.Substring(promptValue.IndexOf(",") + 1);
                EditForm obj = new EditForm();
                obj.Show();
                obj.WindowState = FormWindowState.Normal;
            }
        }

        private void linkLabel2_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            if (Fileinfo.InsurerCode == "GGIC")
            {
                Fileinfo.Insurer = "GGIC,Godigit General Insurance Co. Ltd.";
            }
            else if (Fileinfo.InsurerCode == "TAIG")
            {
                Fileinfo.Insurer = "TAIG,TATA AIG General Insurance Co. Ltd.";
            }
            else if (Fileinfo.InsurerCode == "STAR")
            {
                Fileinfo.Insurer = "STAR,Star Health & Allied Insurance Co.Ltd.";
            }
            else if (Fileinfo.InsurerCode == "ILGI")
            {
                Fileinfo.Insurer = "ILGI,ICICI Lombard General Insurance Co. Ltd.";
            }
            else if (Fileinfo.InsurerCode == "NACL")
            {
                Fileinfo.Insurer = "NACL,National Insurance Co. Ltd.";
            }
            else if (Fileinfo.InsurerCode == "BAGI")
            {
                Fileinfo.Insurer = "BAGI,Bajaj Allianz General Insurance Co. Ltd.";
            }
            else if (Fileinfo.InsurerCode == "CHIC")
            {
                Fileinfo.Insurer = "CHIC,Care Health Insurance Co. Ltd.";
            }
            else if (Fileinfo.InsurerCode == "CMGI")
            {
                Fileinfo.Insurer = "CMGI,Chola MS General Insurance Co. Ltd.";
            }
            else if (Fileinfo.InsurerCode == "ABHI")
            {
                Fileinfo.Insurer = "ABHI,Aditya Birla Health Insurance Co.Ltd.";
            }
            else if (Fileinfo.InsurerCode == "EGIC")
            {
                Fileinfo.Insurer = "EGIC,Edelweiss General Insurance Co.Ltd.";
            }
            else if (Fileinfo.InsurerCode == "HEGI")
            {
                Fileinfo.Insurer = "HEGI,HDFC Ergo General Insurance Co. Ltd.";
            }
            else if (Fileinfo.InsurerCode == "ITGI")
            {
                Fileinfo.Insurer = "ITGI,Iffco Tokio General Insurance Co.Ltd.";
            }
            else if (Fileinfo.InsurerCode == "LVGI")
            {
                Fileinfo.Insurer = "LVGI,Liberty Videocon General Insurance Co.Ltd.";
            }
            else if (Fileinfo.InsurerCode == "RQGI")
            {
                Fileinfo.Insurer = "RQGI,Raheja QBE General Insurance Co. Ltd.";
            }
            else if (Fileinfo.InsurerCode == "RGIC")
            {
                Fileinfo.Insurer = "RGIC,Reliance General Insurance Co. Ltd.";
            }
            else if (Fileinfo.InsurerCode == "RSGI")
            {
                Fileinfo.Insurer = "RSGI,Royal Sundaram General Insurance Co Ltd.";
            }
            else if (Fileinfo.InsurerCode == "SBII")
            {
                Fileinfo.Insurer = "SBII,SBI General Insurance Co. Ltd.";
            }

            Fileinfo.BatchId = BatchID;
            EditForm obj = new EditForm();
            obj.Show();
            obj.WindowState = FormWindowState.Normal;
        }

        private void RBType3_CheckedChanged(object sender, EventArgs e)
        {
            if (RBType3.Checked == true)
            {
                btnBrowse.Enabled = true;
                Fileinfo.Type = "General";
                label15.Visible = true;
                label9.Visible = true;
                label11.Visible = true;
                label12.Visible = true;
                DDInsurance.Enabled = true;
                DDLocation.Enabled = true;
                DDsales.Enabled = true;
                DDService.Enabled = true;
                DDSupport.Enabled = true;
                DDMonth.Enabled = true;
            }
            else
            {
                btnBrowse.Enabled = false;
            }
        }

        private void RBType4_CheckedChanged(object sender, EventArgs e)
        {
            if (RBType4.Checked == true)
            {
                btnBrowse.Enabled = true;
                Fileinfo.Type = "Terrorism";
                label15.Visible = false;
                label9.Visible = false;
                label11.Visible = false;
                label12.Visible = false;
                DDInsurance.Enabled = false;
                DDLocation.Enabled = false;
                DDsales.Enabled = false;
                DDService.Enabled = false;
                DDSupport.Enabled = false;
                DDMonth.Enabled = false;
            }
            else
            {
                btnBrowse.Enabled = false;
            }
        }

        private void RBType2_CheckedChanged(object sender, EventArgs e)
        {
            if (RBType2.Checked == true)
            {
                Fileinfo.Type = "Corporate";
            }
        }

        private void RBType1_CheckedChanged(object sender, EventArgs e)
        {
            if (RBType1.Checked == true)
            {
                Fileinfo.Type = "Retail";
            }
        }

        private void kryptonButton3_Click(object sender, EventArgs e)
        {
            if (Fileinfo.ReportId == "GGI1") //Godigit General Insurance Co. Ltd
            {
                Godigitsample1 obj = new Godigitsample1();
                obj.Show();
            }
            if (Fileinfo.ReportId == "TAI1") //TATA AIG General Insurance Co. Ltd.
            {
                TATASample1 obj = new TATASample1();
                obj.Show();
            }
            if (Fileinfo.ReportId == "ILG1") //ICICI Lombard General Insurance Co. Ltd.
            {
                ICICIsample1 obj = new ICICIsample1();
                obj.Show();
            }
            if (Fileinfo.ReportId == "STA1") //Star Health & Allied Insurance Co. Ltd.
            {
                StarHealthSample1 obj = new StarHealthSample1();
                obj.Show();
            }
            if (Fileinfo.ReportId == "NACX") //National Insurance Co. Ltd.
            {
                NationalExcelSample obj = new NationalExcelSample();
                obj.Show();
            }
            if (Fileinfo.ReportId == "BAGX") //Bajajational Insurance Co. Ltd.
            {
                BajajSample1 obj = new BajajSample1();
                obj.Show();
            }
            if (Fileinfo.ReportId == "CHIX") //Care Health Insurance Co. Ltd.
            {
                CareHealthSample obj = new CareHealthSample();
                obj.Show();
            }
            if (Fileinfo.ReportId == "CMGX") //Chola MS General Insurance Co. Ltd.
            {
                CholaSample obj = new CholaSample();
                obj.Show();
            }
            if (Fileinfo.ReportId == "ABHX") //Aditya Birla Health Insurance Co. Ltd.
            {
                AdityaTemplate1 obj = new AdityaTemplate1();
                obj.Show();
            }
            if (Fileinfo.ReportId == "EGIX") //Edelweiss General Insurance Co. Ltd.
            {
                EdelweissSample obj = new EdelweissSample();
                obj.Show();
            }
            if (Fileinfo.ReportId == "HEGX") //HDFC Ergo General Insurance Co. Ltd.
            {
                HDFCSample obj = new HDFCSample();
                obj.Show();
            }
            if (Fileinfo.ReportId == "ITGX") //Iffco Tokio General Insurance Co.Ltd.
            {
                IFFCOSample obj = new IFFCOSample();
                obj.Show();
            }
            if (Fileinfo.ReportId == "LVGX") //Liberty Videocon General Insurance Co. Ltd.
            {
                LibertySample obj = new LibertySample();
                obj.Show();
            }
            if (Fileinfo.ReportId == "RQGX") //Raheja QBE General Insurance Co. Ltd.
            {
                RahejaQBESample obj = new RahejaQBESample();
                obj.Show();
            }
            if (Fileinfo.ReportId == "RGIX") //Reliance General Insurance Co. Ltd.
            {
                RelianceSample obj = new RelianceSample();
                obj.Show();
            }
            if (Fileinfo.ReportId == "RSGX") //Royal Sundaram General Insurance Co Ltd.
            {
                RoyalExcelSample obj = new RoyalExcelSample();
                obj.Show();
            }
            if (Fileinfo.ReportId == "SBIX") //SBI General Insurance Co. Ltd.
            {
                SBISample obj = new SBISample();
                obj.Show();
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
            if (Fileinfo.InsurerCode == "STAR" || Fileinfo.InsurerCode == "ILGI")
            {
                RBType1.Enabled = true;
                RBType2.Enabled = true;
                RBType3.Enabled = true;
                RBType4.Enabled = true;
            }
            if (Fileinfo.InsurerCode != "STAR")
            {
                lbltype.Visible = false;
                RBType1.Visible = false;
                RBType2.Visible = false;
                label3.Visible = false;
            }
            if (Fileinfo.InsurerCode != "ILGI")
            {
                lblType2.Visible = false;
                RBType3.Visible = false;
                RBType4.Visible = false;
                label3.Visible = false;
            }
            if (Fileinfo.InsurerCode == "ILGI")
            {
                if (Fileinfo.Type != "Terrorism")
                {
                    label15.Visible = true;
                    label9.Visible = true;
                    label11.Visible = true;
                    label12.Visible = true;
                }
                else if (Fileinfo.Type == "Terrorism")
                {
                    DDInsurance.Enabled = false;
                    DDLocation.Enabled = false;
                    DDsales.Enabled = false;
                    DDService.Enabled = false;
                    DDSupport.Enabled = false;
                    DDMonth.Enabled = false;
                }
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
            string com = "select distinct(Inv_No) as No,Inv_No+','+[Filename] as Name from BDSMaster where InsurerCode = '" + Fileinfo.InsurerCode + "' and ReportCode = '" + Fileinfo.ReportId + "'";
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

        public GodigitInsurence()
        {
            InitializeComponent();
            lblUser.Text = LoginInfo.UserID;
            if (Fileinfo.ReportId == "GGI1")
            {
                lblHeader.Text = "GGIC (Godigit General Insurance Co. Ltd.) - Report Id: GGI1";
            }
            else if (Fileinfo.ReportId == "TAI1")
            {
                lblHeader.Text = "TAIG (TATA General Insurance Co. Ltd.) - Report Id: TAI1";
            }
            else if (Fileinfo.ReportId == "NACX")
            {
                lblHeader.Text = "NACL (National Insurance Co. Ltd.) - Report Id: NACX";
            }
            else if (Fileinfo.ReportId == "BAGX")
            {
                lblHeader.Text = "BAGI (Bajaj Allianz General Insurance Co.Ltd.) - Report Id: BAGX";
            }
            else if (Fileinfo.ReportId == "CHIX")
            {
                lblHeader.Text = "CHIC (Care Health Insurance Co. Ltd.) - Report Id: CHIX";
            }
            else if (Fileinfo.ReportId == "CMGX")
            {
                lblHeader.Text = "CMGI (Chola MS General Insurance Co. Ltd.) - Report Id: CMGX";
            }
            else if (Fileinfo.ReportId == "ABHX")
            {
                lblHeader.Text = "ABHI (Aditya Birla Health Insurance Co.Ltd.) - Report Id: ABHX";
            }
            else if (Fileinfo.ReportId == "EGIX")
            {
                lblHeader.Text = "EGIC (Edelweiss General Insurance Co.Ltd.) - Report Id: EGIX";
            }
            else if (Fileinfo.ReportId == "HEGX")
            {
                lblHeader.Text = "HEGI (HDFC Ergo General Insurance Co. Ltd.) - Report Id: HEGX";
            }
            else if (Fileinfo.ReportId == "ITGX")
            {
                lblHeader.Text = "ITGI (Iffco Tokio General Insurance Co.Ltd.) - Report Id: ITGX";
            }
            else if (Fileinfo.ReportId == "LVGX")
            {
                lblHeader.Text = "LVGI(Liberty Videocon General Insurance Co.Ltd.)-Report Id: LVGX";
            }
            else if (Fileinfo.ReportId == "RQGX")
            {
                lblHeader.Text = "RQGI (Raheja QBE General Insurance Co.Ltd.) - Report Id: RQGX";
            }
            else if (Fileinfo.ReportId == "RGIX")
            {
                lblHeader.Text = "RGIC (Reliance General Insurance Co.Ltd.) - Report Id: RGIX";
            }
            else if (Fileinfo.ReportId == "SBIX")
            {
                lblHeader.Text = "SBII (SBI General Insurance Co. Ltd.) - Report Id: SBIX";
            }
            else if (Fileinfo.ReportId == "RSGX")
            {
                lblHeader.Text = "RSGI(Royal Sundaram General Insurance Co.Ltd.)-Report Id: RSGX";
            }
            else if (Fileinfo.ReportId == "STA1")
            {
                lblHeader.Text = "STAR(Star Health and Allied Insurance Co.Ltd.) - Report Id: STA1";
                lbltype.Visible = true;
                RBType1.Visible = true;
                RBType2.Visible = true;
                label3.Visible = true;
            }
            else if (Fileinfo.ReportId == "ILG1")
            {
                lblHeader.Text = "ILGI(ICICI Lombard General Insurance Co. Ltd.) - Report Id: ILG1";
                lblType2.Visible = true;
                RBType3.Visible = true;
                RBType4.Visible = true;
                label3.Visible = true;
            }
            BindDDInsurance();
            BindDDSales();
            BindDDService();
            BindDDLocation();
            BindDDSupport();
            DDMonth.Select();
            TimeUpdater();
            //if (Fileinfo.InsurerCode == "ILGI" && Fileinfo.Type == "Terrorism")
            //{
            //    btnBrowse.Enabled = true;
            //}
            //else
            //{
            //    btnBrowse.Enabled = false;
            //}
        }
        async void TimeUpdater()
        {
            while (true)
            {
                lblTimer.Text = DateTime.Now.ToString("dd-MM-yyyy hh:mm:ss tt");
                await Task.Delay(1000);
            }
        }
        public void BindDDInsurance()
        {
            DataRow dr; var groupby = ""; string com = "";
            if (Fileinfo.InsurerCode == "GGIC")
            {
                groupby = "GO";
            }
            else if (Fileinfo.InsurerCode == "TAIG")
            {
                groupby = "TA";
            }
            else if (Fileinfo.InsurerCode == "STAR")
            {
                groupby = "ST";
            }
            else if (Fileinfo.InsurerCode == "ILGI")
            {
                groupby = "IC";
            }
            else if (Fileinfo.InsurerCode == "NACL")
            {
                groupby = "NA";
            }
            else if (Fileinfo.InsurerCode == "BAGI")
            {
                groupby = "BA";
            }
            else if (Fileinfo.InsurerCode == "CHIC")
            {
                groupby = "CA";
            }
            else if (Fileinfo.InsurerCode == "CMGI")
            {
                groupby = "CH";
            }
            else if (Fileinfo.InsurerCode == "ABHI")
            {
                groupby = "AD";
            }
            else if (Fileinfo.InsurerCode == "EGIC")
            {
                groupby = "ED";
            }
            else if (Fileinfo.InsurerCode == "HEGI")
            {
                groupby = "HD";
            }
            else if (Fileinfo.InsurerCode == "ITGI")
            {
                groupby = "IF";
            }
            else if (Fileinfo.InsurerCode == "LVGI")
            {
                groupby = "LI";
            }
            else if (Fileinfo.InsurerCode == "RQGI")
            {
                groupby = "RA";
            }
            else if (Fileinfo.InsurerCode == "RGIC")
            {
                groupby = "RE";
            }
            else if (Fileinfo.InsurerCode == "RSGI")
            {
                groupby = "RO";
            }
            else if (Fileinfo.InsurerCode == "SBII")
            {
                groupby = "SB";
            }
            //string com = "select Code,InsurerCode + ','+ UPPER(LEFT(Description, 1)) + LOWER(RIGHT(Description, LEN(Description) - 1)) as Description from tblBRInsurancelkup where GroupBy = 'UN' and Code != '' order by Description asc";
            //string com = "select Code,InsurerCode + ',' + Code +' '+ UPPER(LEFT(Description, 1)) + LOWER(RIGHT(Description, LEN(Description) - 1)) as Description from tblBRInsurancelkup where GroupBy = 'UN' and Code != '' order by Description asc";
            if (Fileinfo.InsurerCode == "NACL")
            {
                com = "select Code,Code +' '+ InsurerCode + ',' + UPPER(LEFT(Description, 1)) + LOWER(RIGHT(Description, LEN(Description) - 1)) as Description from tblBRInsurancelkup where GroupBy = '" + groupby + "' order by Description asc";
            }
            else
            {
                com = "select ROW_NUMBER() OVER (ORDER BY (SELECT 1)) as Code,InsurerCode + ',' + UPPER(LEFT(Description, 1)) + LOWER(RIGHT(Description, LEN(Description) - 1)) as Description from tblBRInsurancelkup where GroupBy = '" + groupby + "' order by Description asc";
            }
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
            if (Fileinfo.InsurerCode == "ILGI" && Fileinfo.Type == "Terrorism")
            {
                if (RBType3.Checked != true && RBType4.Checked != true)
                {
                    MessageBox.Show("Please select Type");
                }
                else
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
                        Fileextn = System.IO.Path.GetExtension(openFileDialog1.FileName);

                        txtfile.Text = Filewithext;
                        txtfile.ForeColor = System.Drawing.Color.Black;
                        Fileinfo.Filename = Filewithext;
                        Fileinfo.TName = "ICICITransaction";

                        btnBrowse.Enabled = false;
                    }
                    else
                    {
                        btnBrowse.Enabled = true;
                        btnCancel.Enabled = true;
                        btnConvert.Enabled = false;
                    }
                }
            }
            else
            {
                if (DDInsurance.SelectedValue.ToString() != "0" && DDsales.SelectedValue.ToString() != "0" && DDService.SelectedValue.ToString() != "0" && (DDMonth.SelectedIndex != -1 && DDMonth.SelectedIndex != 0))
                {
                    if (Fileinfo.InsurerCode == "STAR" && RBType1.Checked != true && RBType2.Checked != true)
                    {
                        MessageBox.Show("Please select Type");
                    }
                    else
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
                            Fileextn = System.IO.Path.GetExtension(openFileDialog1.FileName);

                            txtfile.Text = Filewithext;
                            txtfile.ForeColor = System.Drawing.Color.Black;
                            Fileinfo.Filename = Filewithext;
                            if (Fileinfo.InsurerCode == "GGIC")
                            {
                                Fileinfo.TName = "GodigitTransaction";
                            }
                            else if (Fileinfo.InsurerCode == "TAIG")
                            {
                                Fileinfo.TName = "TATATransaction";
                            }
                            else if (Fileinfo.InsurerCode == "STAR")
                            {
                                Fileinfo.TName = "StarHelthTransaction";
                            }
                            else if (Fileinfo.InsurerCode == "ILGI")
                            {
                                Fileinfo.TName = "ICICITransaction";
                            }
                            else if (Fileinfo.InsurerCode == "NACL")
                            {
                                Fileinfo.TName = "NationalExcelTransaction";
                            }
                            else if (Fileinfo.InsurerCode == "BAGI")
                            {
                                Fileinfo.TName = "BAJAJTransaction";
                            }
                            else if (Fileinfo.InsurerCode == "CHIC")
                            {
                                Fileinfo.TName = "CareHealthTransaction";
                            }
                            else if (Fileinfo.InsurerCode == "CMGI")
                            {
                                Fileinfo.TName = "CholaTransaction";
                            }
                            else if (Fileinfo.InsurerCode == "ABHI")
                            {
                                Fileinfo.TName = "AdityaBirlaTransaction";
                            }
                            else if (Fileinfo.InsurerCode == "EGIC")
                            {
                                Fileinfo.TName = "EdelweissGeneralTransaction";
                            }
                            else if (Fileinfo.InsurerCode == "HEGI")
                            {
                                Fileinfo.TName = "HDFCTransaction";
                            }
                            else if (Fileinfo.InsurerCode == "ITGI")
                            {
                                Fileinfo.TName = "IFFCOTokioTransaction";
                            }
                            else if (Fileinfo.InsurerCode == "LVGI")
                            {
                                Fileinfo.TName = "LibertyVideoconTransaction";
                            }
                            else if (Fileinfo.InsurerCode == "RQGI")
                            {
                                Fileinfo.TName = "RahejaTransaction";
                            }
                            else if (Fileinfo.InsurerCode == "RGIC")
                            {
                                Fileinfo.TName = "RelianceTransaction";
                            }
                            else if (Fileinfo.InsurerCode == "SBII")
                            {
                                Fileinfo.TName = "SBITransaction";
                            }
                            else if (Fileinfo.InsurerCode == "RSGI")
                            {
                                Fileinfo.TName = "RoyalTransaction";
                            }
                            btnBrowse.Enabled = false;
                        }
                        else
                        {
                            btnBrowse.Enabled = true;
                            btnCancel.Enabled = true;
                            btnConvert.Enabled = false;
                        }
                    }
                }
                else
                {
                    if (DDInsurance.SelectedValue.ToString() == "0")
                    {
                        MessageBox.Show("Please select Insurance");
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
        }
        //public static void InsertTransaction(Microsoft.Office.Interop.Excel.Workbook WB, string TranID, string RDate, string Insurance, string Salesby, string Serviceby, string location, string Support, string Rmonth)
        //{
        //    SQLProcs sql = new SQLProcs();
        //    Microsoft.Office.Interop.Excel.Worksheet wks = (Microsoft.Office.Interop.Excel.Worksheet)WB.Worksheets[1];
        //    Microsoft.Office.Interop.Excel.Range lastCell = wks.Cells.SpecialCells(Microsoft.Office.Interop.Excel.XlCellType.xlCellTypeLastCell, Type.Missing);
        //    int lastrow = lastCell.Row; var Terrorism = ""; var Policy_Endorsement = ""; var InsuredType = "";
        //    for (int i = 2; i <= lastrow; i++)
        //    {
        //        string InsuredName = ((Microsoft.Office.Interop.Excel.Range)wks.Cells[i, 30]).Value.Replace("\n", "").TrimStart();
        //        if (InsuredName == null || InsuredName == "")
        //        {
        //            InsuredName = ((Microsoft.Office.Interop.Excel.Range)wks.Cells[i, 37]).Value.Replace("\n", "").TrimStart();
        //        }
        //        string PolicyNo = ((Microsoft.Office.Interop.Excel.Range)wks.Cells[i, 38]).Value.Replace("\n", "").TrimStart();
        //        if (PolicyNo == null || PolicyNo == "")
        //        {
        //            PolicyNo = ((Microsoft.Office.Interop.Excel.Range)wks.Cells[i, 1]).Value.Replace("\n", "").TrimStart();
        //        }
        //        //var Endo_Effective_Date = Convert.ToString(((Microsoft.Office.Interop.Excel.Range)wks.Cells[i, 4]).Value);//.Replace("\n", "").TrimStart();
        //        //string END_Date = Convert.ToString(((Microsoft.Office.Interop.Excel.Range)wks.Cells[i, 6]).Value);//.Replace("\n", "").TrimStart();
        //        var Endo_Effective_Date = ((Microsoft.Office.Interop.Excel.Range)wks.Cells[i, 4]).Value;
        //        var END_Date = ((Microsoft.Office.Interop.Excel.Range)wks.Cells[i, 6]).Value;
        //        Endo_Effective_Date = Endo_Effective_Date.ToString("dd/MM/yyyy");
        //        END_Date = END_Date.ToString("dd/MM/yyyy");
        //        var Premium_Amt = Convert.ToString(((Microsoft.Office.Interop.Excel.Range)wks.Cells[i, 31]).Value);
        //        var Revenue_Amt = Convert.ToString(((Microsoft.Office.Interop.Excel.Range)wks.Cells[i, 47]).Value);
        //        var Revenue_Pct = Convert.ToString(((Microsoft.Office.Interop.Excel.Range)wks.Cells[i, 43]).Value);
        //        var TP_Amt = Convert.ToString(((Microsoft.Office.Interop.Excel.Range)wks.Cells[i, 32]).Value);
        //        Terrorism = Convert.ToString(((Microsoft.Office.Interop.Excel.Range)wks.Cells[i, 41]).Value);
        //        var RewardOD_Pct = Convert.ToString(((Microsoft.Office.Interop.Excel.Range)wks.Cells[i, 44]).Value);
        //        var RewardTP_Pct = Convert.ToString(((Microsoft.Office.Interop.Excel.Range)wks.Cells[i, 46]).Value);
        //        var IRDARewardAmt = Convert.ToString(((Microsoft.Office.Interop.Excel.Range)wks.Cells[i, 48]).Value);
        //        var Policy_Type = ((Microsoft.Office.Interop.Excel.Range)wks.Cells[i, 15]).Value.Replace("\n", "").Replace(",", "").TrimStart();
        //        if (InsuredName.Contains("LIMITED"))
        //        {
        //            InsuredType = "corporate";
        //        }
        //        else
        //        {
        //            InsuredType = "Retail";
        //        }
        //        if (Premium_Amt == "" || Premium_Amt == " " || Terrorism == null)
        //        {
        //            Premium_Amt = 0;
        //        }
        //        if (Revenue_Amt == "" || Revenue_Amt == " " || Terrorism == null)
        //        {
        //            Revenue_Amt = 0;
        //        }
        //        if (Terrorism == "" || Terrorism == " " || Terrorism == null)
        //        {
        //            Terrorism = "0";
        //        }
        //        if (TP_Amt == "" || TP_Amt == " " || Terrorism == null)
        //        {
        //            TP_Amt = "0";
        //        }
        //        if (Revenue_Pct == "" || Revenue_Pct == " " || Terrorism == null)
        //        {
        //            Revenue_Pct = "0";
        //        }
        //        if (RewardOD_Pct == "" || RewardOD_Pct == " " || Terrorism == null)
        //        {
        //            RewardOD_Pct = "0";
        //        }
        //        if (RewardTP_Pct == "" || RewardTP_Pct == " " || Terrorism == null)
        //        {
        //            RewardTP_Pct = "0";
        //        }
        //        if (IRDARewardAmt == "" || IRDARewardAmt == " " || Terrorism == null)
        //        {
        //            IRDARewardAmt = "0";
        //        }
        //        sql.ExecuteSQLNonQuery("SP_GoDigitTransactions",
        //                   new SqlParameter { ParameterName = "@Imode", Value = 1 },
        //                   new SqlParameter { ParameterName = "@RDate", Value = RDate },
        //                   new SqlParameter { ParameterName = "@Rmonth", Value = Rmonth },
        //                   new SqlParameter { ParameterName = "@ClientName", Value = InsuredName },
        //                   new SqlParameter { ParameterName = "@Itype", Value = InsuredType },
        //                   new SqlParameter { ParameterName = "@Policy_No", Value = PolicyNo },
        //                   new SqlParameter { ParameterName = "@Endo_Effective_Date", Value = Endo_Effective_Date },
        //                   new SqlParameter { ParameterName = "@END_Date", Value = END_Date },
        //                   new SqlParameter { ParameterName = "@Policy_Type", Value = Policy_Type },
        //                   new SqlParameter { ParameterName = "@Premium_Amt", Value = Premium_Amt },
        //                   new SqlParameter { ParameterName = "@TranID", Value = TranID },
        //                   new SqlParameter { ParameterName = "@Revenue_Amt", Value = Revenue_Amt },
        //                   new SqlParameter { ParameterName = "@Terrorism", Value = Terrorism },
        //                   new SqlParameter { ParameterName = "@TP_Amt", Value = TP_Amt },
        //                   new SqlParameter { ParameterName = "@Revenue_Pct", Value = Revenue_Pct },
        //                   new SqlParameter { ParameterName = "@RewardOD_Pct", Value = RewardOD_Pct },
        //                   new SqlParameter { ParameterName = "@RewardTP_Pct", Value = RewardTP_Pct },
        //                   new SqlParameter { ParameterName = "@IRDARewardAmt", Value = IRDARewardAmt },
        //                   new SqlParameter { ParameterName = "@Insurance", Value = Insurance },
        //                   new SqlParameter { ParameterName = "@Salesby", Value = Salesby },
        //                   new SqlParameter { ParameterName = "@Serviceby", Value = Serviceby },
        //                   new SqlParameter { ParameterName = "@location", Value = location },
        //                   new SqlParameter { ParameterName = "@Support", Value = Support },
        //                   new SqlParameter { ParameterName = "@Policy_Endorsement", Value = Policy_Endorsement },
        //                   new SqlParameter { ParameterName = "@RFormat", Value = "F1" },
        //                   new SqlParameter { ParameterName = "@InvNo", Value = "GGIC" },
        //                   new SqlParameter { ParameterName = "@ReportId", Value = "GGI1" },
        //                   new SqlParameter { ParameterName = "@DocName", Value = "Godigit General Insurance Co. Ltd." }
        //                   );
        //    }
        //}

        //public static void InsertTransaction2(Microsoft.Office.Interop.Excel.Workbook WB, string TranID, string RDate, string Insurance, string Salesby, string Serviceby, string location, string Support, string Rmonth)
        //{

        //    SQLProcs sql = new SQLProcs();
        //    Microsoft.Office.Interop.Excel.Worksheet wks = (Microsoft.Office.Interop.Excel.Worksheet)WB.Worksheets[1];
        //    Microsoft.Office.Interop.Excel.Range lastCell = wks.Cells.SpecialCells(Microsoft.Office.Interop.Excel.XlCellType.xlCellTypeLastCell, Type.Missing);
        //    int lastrow = lastCell.Row; var Terrorism = ""; var Policy_Endorsement = ""; var InsuredType = "";
        //    for (int i = 2; i <= lastrow; i++)
        //    {
        //        string InsuredName = ((Microsoft.Office.Interop.Excel.Range)wks.Cells[i, 8]).Value;
        //        if (InsuredName != null && InsuredName != "" && InsuredName != " ")
        //        {
        //            InsuredName = InsuredName.Replace("\n", "").TrimStart();
        //            string PolicyNo = ((Microsoft.Office.Interop.Excel.Range)wks.Cells[i, 2]).Value.Replace("\n", "").TrimStart();
        //            var Endo_Effective_Date = ((Microsoft.Office.Interop.Excel.Range)wks.Cells[i, 7]).Value;
        //            var Effective_Date = ((Microsoft.Office.Interop.Excel.Range)wks.Cells[i, 5]).Value;
        //            var END_Date = ((Microsoft.Office.Interop.Excel.Range)wks.Cells[i, 6]).Value;
        //            //Endo_Effective_Date = Endo_Effective_Date.ToString("dd/MM/yyyy");
        //            //Effective_Date = Effective_Date.ToString("dd/MM/yyyy");
        //            //END_Date = END_Date.ToString("dd/MM/yyyy");
        //            Policy_Endorsement = ((Microsoft.Office.Interop.Excel.Range)wks.Cells[i, 11]).Value.Replace("\n", "").TrimStart();
        //            if (Policy_Endorsement == "New")
        //            {
        //                Policy_Endorsement = "Policy";
        //            }
        //            var Premium_Amt = Convert.ToString(((Microsoft.Office.Interop.Excel.Range)wks.Cells[i, 12]).Value);
        //            var Revenue_Amt = Convert.ToString(((Microsoft.Office.Interop.Excel.Range)wks.Cells[i, 18]).Value);
        //            var Total_Premium = Convert.ToString(((Microsoft.Office.Interop.Excel.Range)wks.Cells[i, 14]).Value);
        //            Terrorism = Convert.ToString(((Microsoft.Office.Interop.Excel.Range)wks.Cells[i, 13]).Value);

        //            var Policy_Type = ((Microsoft.Office.Interop.Excel.Range)wks.Cells[i, 9]).Value.Replace("\n", "").Replace(",", "").TrimStart();

        //            if (InsuredName.Contains("LIMITED") || InsuredName.Contains("LTD") || InsuredName.Contains(".COM"))
        //            {
        //                InsuredType = "Corporate";
        //            }
        //            else
        //            {
        //                InsuredType = "Retail";
        //            }
        //            if (Premium_Amt == "" || Premium_Amt == " " || Terrorism == null)
        //            {
        //                Premium_Amt = 0;
        //            }
        //            if (Revenue_Amt == "" || Revenue_Amt == " " || Terrorism == null)
        //            {
        //                Revenue_Amt = 0;
        //            }
        //            if (Terrorism == "" || Terrorism == " " || Terrorism == null)
        //            {
        //                Terrorism = "0";
        //            }
        //            if (Total_Premium == "" || Total_Premium == " " || Terrorism == null)
        //            {
        //                Total_Premium = "0";
        //            }
        //            sql.ExecuteSQLNonQuery("SP_TATATransactions",
        //                       new SqlParameter { ParameterName = "@Imode", Value = 1 },
        //                       new SqlParameter { ParameterName = "@RDate", Value = RDate },
        //                       new SqlParameter { ParameterName = "@Rmonth", Value = Rmonth },
        //                       new SqlParameter { ParameterName = "@ClientName", Value = InsuredName },
        //                       new SqlParameter { ParameterName = "@Itype", Value = InsuredType },
        //                       new SqlParameter { ParameterName = "@Policy_No", Value = PolicyNo },
        //                       new SqlParameter { ParameterName = "@Endo_Effective_Date", Value = Endo_Effective_Date },
        //                       new SqlParameter { ParameterName = "@Effective_Date", Value = Effective_Date },
        //                       new SqlParameter { ParameterName = "@END_Date", Value = END_Date },
        //                       new SqlParameter { ParameterName = "@Policy_Type", Value = Policy_Type },
        //                       new SqlParameter { ParameterName = "@Premium_Amt", Value = Premium_Amt },
        //                       new SqlParameter { ParameterName = "@TranID", Value = TranID },
        //                       new SqlParameter { ParameterName = "@Revenue_Amt", Value = Revenue_Amt },
        //                       new SqlParameter { ParameterName = "@Terrorism", Value = Terrorism },
        //                       new SqlParameter { ParameterName = "@Total_Premium", Value = Total_Premium },
        //                       new SqlParameter { ParameterName = "@Insurance", Value = Insurance },
        //                       new SqlParameter { ParameterName = "@Salesby", Value = Salesby },
        //                       new SqlParameter { ParameterName = "@Serviceby", Value = Serviceby },
        //                       new SqlParameter { ParameterName = "@location", Value = location },
        //                       new SqlParameter { ParameterName = "@Support", Value = Support },
        //                       new SqlParameter { ParameterName = "@Policy_Endorsement", Value = Policy_Endorsement },
        //                       new SqlParameter { ParameterName = "@RFormat", Value = "F1" },
        //                       new SqlParameter { ParameterName = "@InvNo", Value = "TAIG" },
        //                       new SqlParameter { ParameterName = "@ReportId", Value = "TAI1" },
        //                       new SqlParameter { ParameterName = "@DocName", Value = "TATA AIG General Insurance Co. Ltd." }
        //                       );
        //        }
        //    }
        //}

        //public static void InsertTransaction3(Microsoft.Office.Interop.Excel.Workbook WB, string TranID, string RDate, string Insurance, string Salesby, string Serviceby, string location, string Support, string Rmonth)
        //{
        //    //Star health  - F1
        //    SQLProcs sql = new SQLProcs();
        //    Microsoft.Office.Interop.Excel.Worksheet wks = (Microsoft.Office.Interop.Excel.Worksheet)WB.Worksheets[1];
        //    Microsoft.Office.Interop.Excel.Range lastCell = wks.Cells.SpecialCells(Microsoft.Office.Interop.Excel.XlCellType.xlCellTypeLastCell, Type.Missing);
        //    int lastrow = lastCell.Row; var Terrorism = ""; var Policy_Endorsement = ""; var InsuredType = "";
        //    for (int i = 2; i <= lastrow; i++)
        //    {
        //        string InsuredName = ((Microsoft.Office.Interop.Excel.Range)wks.Cells[i, 3]).Value;
        //        if (InsuredName != null && InsuredName != "" && InsuredName != " ")
        //        {
        //            InsuredName = InsuredName.Replace("\n", "").TrimStart();
        //            var Client_N_E = ((Microsoft.Office.Interop.Excel.Range)wks.Cells[i, 18]).Value.Replace("\n", "").TrimStart();
        //            if (Client_N_E == "FRESH")
        //            {
        //                Client_N_E = "New";
        //            }
        //            string PolicyNo = ((Microsoft.Office.Interop.Excel.Range)wks.Cells[i, 2]).Value.Replace("\n", "").TrimStart();
        //            var Endo_Effective_Date = ((Microsoft.Office.Interop.Excel.Range)wks.Cells[i, 13]).Value;
        //            var Effective_Date = ((Microsoft.Office.Interop.Excel.Range)wks.Cells[i, 11]).Value;
        //            var END_Date = ((Microsoft.Office.Interop.Excel.Range)wks.Cells[i, 12]).Value;
        //            int ENdolen = Convert.ToString(Endo_Effective_Date).Length;
        //            int Efflen = Convert.ToString(Effective_Date).Length;
        //            int ENdlen = Convert.ToString(END_Date).Length;
        //            if (ENdolen > 11)
        //            {
        //                Endo_Effective_Date = Endo_Effective_Date.ToString("dd/MM/yyyy");
        //            }
        //            if (Efflen > 11)
        //            {
        //                Effective_Date = Effective_Date.ToString("dd/MM/yyyy");
        //            }
        //            if (ENdlen > 11)
        //            {
        //                END_Date = END_Date.ToString("dd/MM/yyyy");
        //            }

        //            var Premium_Amt = Convert.ToString(((Microsoft.Office.Interop.Excel.Range)wks.Cells[i, 4]).Value);
        //            var Revenue_Amt = Convert.ToString(((Microsoft.Office.Interop.Excel.Range)wks.Cells[i, 25]).Value);
        //            var Revenue_Pcnt = Convert.ToString(((Microsoft.Office.Interop.Excel.Range)wks.Cells[i, 14]).Value);
        //            //Terrorism = Convert.ToString(((Microsoft.Office.Interop.Excel.Range)wks.Cells[i, 13]).Value);

        //            var Policy_Type = ((Microsoft.Office.Interop.Excel.Range)wks.Cells[i, 19]).Value.Replace("\n", "").Replace(",", "").TrimStart();
        //            var Presult = PolicyNo.Substring(PolicyNo.LastIndexOf('/') + 1);
        //            if (Presult.Contains("0"))
        //            {
        //                Policy_Endorsement = "Policy";
        //            }
        //            else
        //            {
        //                Policy_Endorsement = "Endorsement";
        //            }

        //            //if (InsuredName.Contains("LIMITED") || InsuredName.Contains("LTD") || InsuredName.Contains(".COM"))
        //            //{
        //            InsuredType = "Corporate";
        //            //}
        //            //else
        //            //{
        //            //    InsuredType = "Retail";
        //            //}
        //            if (Premium_Amt == "" || Premium_Amt == " " || Terrorism == null)
        //            {
        //                Premium_Amt = 0;
        //            }
        //            if (Revenue_Amt == "" || Revenue_Amt == " " || Terrorism == null)
        //            {
        //                Revenue_Amt = 0;
        //            }
        //            if (Terrorism == "" || Terrorism == " " || Terrorism == null)
        //            {
        //                Terrorism = "0";
        //            }
        //            if (Revenue_Pcnt == "" || Revenue_Pcnt == " " || Terrorism == null)
        //            {
        //                Revenue_Pcnt = "0";
        //            }
        //            sql.ExecuteSQLNonQuery("SP_StarHealthTransactions",
        //                       new SqlParameter { ParameterName = "@Imode", Value = 1 },
        //                       new SqlParameter { ParameterName = "@RDate", Value = RDate },
        //                       new SqlParameter { ParameterName = "@Rmonth", Value = Rmonth },
        //                       new SqlParameter { ParameterName = "@ClientName", Value = InsuredName },
        //                       new SqlParameter { ParameterName = "@Itype", Value = InsuredType },
        //                       new SqlParameter { ParameterName = "@Client_N_E", Value = Client_N_E },
        //                       new SqlParameter { ParameterName = "@Policy_No", Value = PolicyNo },
        //                       new SqlParameter { ParameterName = "@Endo_Effective_Date", Value = Endo_Effective_Date },
        //                       new SqlParameter { ParameterName = "@Effective_Date", Value = Effective_Date },
        //                       new SqlParameter { ParameterName = "@END_Date", Value = END_Date },
        //                       new SqlParameter { ParameterName = "@Policy_Type", Value = Policy_Type },
        //                       new SqlParameter { ParameterName = "@Premium_Amt", Value = Premium_Amt },
        //                       new SqlParameter { ParameterName = "@TranID", Value = TranID },
        //                       new SqlParameter { ParameterName = "@Revenue_Amt", Value = Revenue_Amt },
        //                       new SqlParameter { ParameterName = "@Terrorism", Value = Terrorism },
        //                       new SqlParameter { ParameterName = "@Revenue_Pcnt", Value = Revenue_Pcnt },
        //                       new SqlParameter { ParameterName = "@Insurance", Value = Insurance },
        //                       new SqlParameter { ParameterName = "@Salesby", Value = Salesby },
        //                       new SqlParameter { ParameterName = "@Serviceby", Value = Serviceby },
        //                       new SqlParameter { ParameterName = "@location", Value = location },
        //                       new SqlParameter { ParameterName = "@Support", Value = Support },
        //                       new SqlParameter { ParameterName = "@Policy_Endorsement", Value = Policy_Endorsement },
        //                       new SqlParameter { ParameterName = "@RFormat", Value = "F1" },
        //                       new SqlParameter { ParameterName = "@InvNo", Value = "STAR" },
        //                       new SqlParameter { ParameterName = "@ReportId", Value = "STA1" },
        //                       new SqlParameter { ParameterName = "@DocName", Value = "Star Health & Allied Insurance Co. Ltd." }
        //                       );
        //        }
        //    }
        //}
        //public static void InsertTransaction4(Microsoft.Office.Interop.Excel.Workbook WB, string TranID, string RDate, string Insurance, string Salesby, string Serviceby, string location, string Support, string Rmonth)
        //{
        //    //Star health Retail - F2
        //    SQLProcs sql = new SQLProcs();
        //    Microsoft.Office.Interop.Excel.Worksheet wks = (Microsoft.Office.Interop.Excel.Worksheet)WB.Worksheets[1];
        //    Microsoft.Office.Interop.Excel.Range lastCell = wks.Cells.SpecialCells(Microsoft.Office.Interop.Excel.XlCellType.xlCellTypeLastCell, Type.Missing);
        //    int lastrow = lastCell.Row; var Terrorism = ""; var Policy_Endorsement = ""; var InsuredType = "";
        //    for (int i = 2; i <= lastrow; i++)
        //    {
        //        string InsuredName = ((Microsoft.Office.Interop.Excel.Range)wks.Cells[i, 15]).Value;
        //        if (InsuredName != null && InsuredName != "" && InsuredName != " ")
        //        {
        //            InsuredName = InsuredName.Replace("\n", "").TrimStart();
        //            string Client_N_E = ((Microsoft.Office.Interop.Excel.Range)wks.Cells[i, 11]).Value.Replace("\n", "").TrimStart();
        //            if (Client_N_E.ToLower() == "fresh")
        //            {
        //                Client_N_E = "New";
        //            }
        //            string PolicyNo = ((Microsoft.Office.Interop.Excel.Range)wks.Cells[i, 10]).Value.Replace("\n", "").TrimStart();
        //            var Endo_Effective_Date = ((Microsoft.Office.Interop.Excel.Range)wks.Cells[i, 3]).Value;
        //            //var Effective_Date = ((Microsoft.Office.Interop.Excel.Range)wks.Cells[i, 11]).Value;
        //            //var END_Date = ((Microsoft.Office.Interop.Excel.Range)wks.Cells[i, 12]).Value;
        //            int ENdolen = Convert.ToString(Endo_Effective_Date).Length;
        //            //int Efflen = Convert.ToString(Effective_Date).Length;
        //            //int ENdlen = Convert.ToString(END_Date).Length;
        //            if (ENdolen > 11)
        //            {
        //                Endo_Effective_Date = Endo_Effective_Date.ToString("dd/MM/yyyy");
        //            }
        //            //if (Efflen > 11)
        //            //{
        //            //    Effective_Date = Effective_Date.ToString("dd/MM/yyyy");
        //            //}
        //            //if (ENdlen > 11)
        //            //{
        //            //    END_Date = END_Date.ToString("dd/MM/yyyy");
        //            //}

        //            var Premium_Amt = Convert.ToString(((Microsoft.Office.Interop.Excel.Range)wks.Cells[i, 23]).Value);
        //            var Revenue_Amt = Convert.ToString(((Microsoft.Office.Interop.Excel.Range)wks.Cells[i, 25]).Value);
        //            var Revenue_Pcnt = Convert.ToString(((Microsoft.Office.Interop.Excel.Range)wks.Cells[i, 24]).Value).Replace("%", "").Replace("\n", "").Replace("0.", "").TrimStart();
        //            //Terrorism = Convert.ToString(((Microsoft.Office.Interop.Excel.Range)wks.Cells[i, 13]).Value);

        //            var Policy_Type = ((Microsoft.Office.Interop.Excel.Range)wks.Cells[i, 6]).Value.Replace("\n", "").Replace(",", "").TrimStart();
        //            //InsuredType = ((Microsoft.Office.Interop.Excel.Range)wks.Cells[i, 5]).Value.Replace("\n", "").Replace(",", "").TrimStart();
        //            InsuredType = "Retail";
        //            if (PolicyNo.StartsWith("p"))
        //            {
        //                var Presult = PolicyNo.Substring(PolicyNo.LastIndexOf('/') + 1);

        //                if (Convert.ToDouble(Presult) > 0)
        //                {
        //                    Policy_Endorsement = "Policy";
        //                }
        //                else
        //                {
        //                    Policy_Endorsement = "Endorsement";
        //                }
        //            }
        //            else
        //            {
        //                Policy_Endorsement = "Policy";
        //            }
        //            if (Premium_Amt == "" || Premium_Amt == " " || Terrorism == null)
        //            {
        //                Premium_Amt = 0;
        //            }
        //            if (Revenue_Amt == "" || Revenue_Amt == " " || Terrorism == null)
        //            {
        //                Revenue_Amt = 0;
        //            }
        //            if (Terrorism == "" || Terrorism == " " || Terrorism == null)
        //            {
        //                Terrorism = "0";
        //            }
        //            if (Revenue_Pcnt == "" || Revenue_Pcnt == " " || Terrorism == null)
        //            {
        //                Revenue_Pcnt = "0";
        //            }
        //            sql.ExecuteSQLNonQuery("SP_StarHealthTransactions",
        //                       new SqlParameter { ParameterName = "@Imode", Value = 1 },
        //                       new SqlParameter { ParameterName = "@RDate", Value = RDate },
        //                       new SqlParameter { ParameterName = "@Rmonth", Value = Rmonth },
        //                       new SqlParameter { ParameterName = "@ClientName", Value = InsuredName },
        //                       new SqlParameter { ParameterName = "@Itype", Value = InsuredType },
        //                       new SqlParameter { ParameterName = "@Client_N_E", Value = Client_N_E },
        //                       new SqlParameter { ParameterName = "@Policy_No", Value = PolicyNo },
        //                       new SqlParameter { ParameterName = "@Endo_Effective_Date", Value = Endo_Effective_Date },
        //                       //new SqlParameter { ParameterName = "@Effective_Date", Value = Effective_Date },
        //                       //new SqlParameter { ParameterName = "@END_Date", Value = END_Date },
        //                       new SqlParameter { ParameterName = "@Policy_Type", Value = Policy_Type },
        //                       new SqlParameter { ParameterName = "@Premium_Amt", Value = Premium_Amt },
        //                       new SqlParameter { ParameterName = "@TranID", Value = TranID },
        //                       new SqlParameter { ParameterName = "@Revenue_Amt", Value = Revenue_Amt },
        //                       new SqlParameter { ParameterName = "@Terrorism", Value = Terrorism },
        //                       new SqlParameter { ParameterName = "@Revenue_Pcnt", Value = Revenue_Pcnt },
        //                       new SqlParameter { ParameterName = "@Insurance", Value = Insurance },
        //                       new SqlParameter { ParameterName = "@Salesby", Value = Salesby },
        //                       new SqlParameter { ParameterName = "@Serviceby", Value = Serviceby },
        //                       new SqlParameter { ParameterName = "@location", Value = location },
        //                       new SqlParameter { ParameterName = "@Support", Value = Support },
        //                       new SqlParameter { ParameterName = "@Policy_Endorsement", Value = Policy_Endorsement },
        //                       new SqlParameter { ParameterName = "@RFormat", Value = "F1" },
        //                       new SqlParameter { ParameterName = "@InvNo", Value = "STAR" },
        //                       new SqlParameter { ParameterName = "@ReportId", Value = "STA1" },
        //                       new SqlParameter { ParameterName = "@DocName", Value = "Star Health & Allied Insurance Co. Ltd." }
        //                       );
        //        }
        //    }
        //}
        //public static void InsertTransaction5(Microsoft.Office.Interop.Excel.Workbook WB, string TranID, string RDate, string Insurance, string Salesby, string Serviceby, string location, string Support, string Rmonth)
        //{
        //    SQLProcs sql = new SQLProcs();
        //    Microsoft.Office.Interop.Excel.Worksheet wks = (Microsoft.Office.Interop.Excel.Worksheet)WB.Worksheets[1];
        //    Microsoft.Office.Interop.Excel.Range lastCell = wks.Cells.SpecialCells(Microsoft.Office.Interop.Excel.XlCellType.xlCellTypeLastCell, Type.Missing);
        //    int lastrow = lastCell.Row; var Terrorism = ""; 
        //    for (int i = 2; i <= lastrow; i++)
        //    {
        //        string PolicyNo = ((Microsoft.Office.Interop.Excel.Range)wks.Cells[i, 14]).Value.Replace("\n", "").TrimStart();
        //        if (PolicyNo == "" || PolicyNo == " " || PolicyNo == null)
        //        {
        //            var Revenue_Amt = Convert.ToString(((Microsoft.Office.Interop.Excel.Range)wks.Cells[i, 26]).Value);
        //            Terrorism = Convert.ToString(((Microsoft.Office.Interop.Excel.Range)wks.Cells[i, 42]).Value);
        //            if (Revenue_Amt == "" || Revenue_Amt == " " || Terrorism == null)
        //            {
        //                Revenue_Amt = 0;
        //            }
        //            if (Terrorism == "" || Terrorism == " " || Terrorism == null)
        //            {
        //                Terrorism = "0";
        //            }
        //            sql.ExecuteSQLNonQuery("SP_ICICITransactions",
        //                       new SqlParameter { ParameterName = "@Imode", Value = 10 },
        //                       new SqlParameter { ParameterName = "@Policy_No", Value = PolicyNo },
        //                       new SqlParameter { ParameterName = "@Revenue_Amt", Value = Revenue_Amt },
        //                       new SqlParameter { ParameterName = "@Terrorism", Value = Terrorism },
        //                       new SqlParameter { ParameterName = "@RFormat", Value = "F1" },
        //                       new SqlParameter { ParameterName = "@InvNo", Value = "ILGI" },
        //                       new SqlParameter { ParameterName = "@ReportId", Value = "ILG1" },
        //                       new SqlParameter { ParameterName = "@UserId", Value = LoginInfo.UserID }
        //                       );
        //        }
        //    }
        //}
        //public void ExceltoDatatable(string TranID, string RDate, string Insurance, string Salesby, string Serviceby, string location, string Support, string Rmonth, string Filepath)
        //{
        //    SQLProcs sql = new SQLProcs();
        //    using (ExcelEngine excelEngine = new ExcelEngine())
        //    {
        //        IApplication application = excelEngine.Excel;
        //        application.DefaultVersion = Syncfusion.XlsIO.ExcelVersion.Xlsx;
        //        FileStream inputStream = new FileStream(Filepath, FileMode.Open, FileAccess.Read);
        //        Syncfusion.XlsIO.IWorkbook workbook = application.Workbooks.Open(inputStream);
        //        IWorksheet worksheet = workbook.Worksheets[0];

        //        //Read data from the worksheet and export to the DataTable.
        //        DataTable DT = worksheet.ExportDataTable(worksheet.UsedRange, ExcelExportDataTableOptions.ColumnNames);
        //        ///var Terrorism = ""; var InsuredType = ""; var Policy_Endorsement = "";

        //        DataTable RDT = copyDatatable(DT);

        //        //DataSet ds = new DataSet();

        //        //ds = 
        //        sql.ExecuteSQLNonQuery("SP_GoDigitTransactions",
        //                           new SqlParameter { ParameterName = "@Imode", Value = 9 },
        //                           new SqlParameter { ParameterName = "@RDate", Value = RDate },
        //                           new SqlParameter { ParameterName = "@Rmonth", Value = Rmonth },
        //                           new SqlParameter { ParameterName = "@TranID", Value = TranID },
        //                           new SqlParameter { ParameterName = "@Insurance", Value = Insurance },
        //                           new SqlParameter { ParameterName = "@Salesby", Value = Salesby },
        //                           new SqlParameter { ParameterName = "@Serviceby", Value = Serviceby },
        //                           new SqlParameter { ParameterName = "@location", Value = location },
        //                           new SqlParameter { ParameterName = "@Support", Value = Support },
        //                           new SqlParameter { ParameterName = "@GodigitTansaction", Value = RDT },
        //                           new SqlParameter { ParameterName = "@RFormat", Value = "F1" },
        //                           new SqlParameter { ParameterName = "@InvNo", Value = "GGIC" },
        //                           new SqlParameter { ParameterName = "@ReportId", Value = "GGI1" },
        //                           new SqlParameter { ParameterName = "@DocName", Value = "Godigit General Insurance Co. Ltd." }
        //                           );
        //    }
        //}
        //public DataTable copyDatatable(DataTable dt)
        //{
        //    DataTable dtcopy = new DataTable();
        //    DataRow dr = null;

        //    foreach (DataColumn column in dt.Columns)
        //    {
        //        if (column.ColumnName.ToLower() == "policy holder" || column.ColumnName.ToLower() == "insured person" || column.ColumnName.ToLower() == "master policy number" ||
        //            column.ColumnName.ToLower() == "policy number" || column.ColumnName.ToLower() == "policy issue date" || column.ColumnName.ToLower() == "risk exp date" ||
        //            column.ColumnName.ToLower() == "od premium" || column.ColumnName.ToLower() == "irda od amt" || column.ColumnName.ToLower() == "irda od %" || column.ColumnName.ToLower() == "tp premium" ||
        //            column.ColumnName.ToLower() == "terrorism_premium" || column.ColumnName.ToLower() == "reward od %" || column.ColumnName.ToLower() == "reward tp%" ||
        //            column.ColumnName.ToLower() == "irda reward amt" || column.ColumnName.ToLower() == "product name")
        //        {
        //            dtcopy.Columns.Add(column.ColumnName);
        //        }
        //    }

        //    foreach (DataRow row in dt.Rows)
        //    {
        //        dr = dtcopy.NewRow();
        //        int j = 0; int k = 0;
        //        for (int i = 0; i < row.ItemArray.Length; i++)
        //        {
        //            string name = dt.Columns[i].ColumnName.ToLower().ToString();
        //            if (name == "policy holder" || name == "insured person" || name == "master policy number" ||
        //            name == "policy number" || name == "policy issue date" || name == "risk exp date" ||
        //            name == "od premium" || name == "irda od amt" || name == "irda od %" || name == "tp premium" ||
        //            name == "terrorism_premium" || name == "reward od %" || name == "reward tp%" ||
        //            name == "irda reward amt" || name == "product name")
        //            {
        //                if (k != 0)
        //                {
        //                    j++;
        //                }
        //                k = 1;
        //            }
        //            if (name == "policy holder")
        //            {
        //                if (row[i].ToString() != "" && row[i].ToString() != null)
        //                {
        //                    dr[j] = row[i].ToString();
        //                }
        //                else
        //                {
        //                    dr[j] = "";
        //                }
        //            }
        //            else if (name == "insured person")
        //            {
        //                dr[j] = row[i].ToString();
        //            }
        //            else if (name == "master policy number")
        //            {
        //                //string teamname = (row[i].ToString()).Substring(0, 3);
        //                dr[j] = row[i].ToString();
        //            }
        //            else if (name == "policy number")
        //            {
        //                dr[j] = row[i].ToString();
        //            }
        //            else if (name == "policy issue date")
        //            {
        //                if (row[i].ToString() != "" && row[i].ToString() != null)
        //                {
        //                    string strDate = row[i].ToString();
        //                    DateTime date = Convert.ToDateTime(strDate, CultureInfo.InvariantCulture);
        //                    dr[j] = date.ToString("yyyy-MM-dd");
        //                }
        //                else
        //                {
        //                    dr[j] = row[i].ToString();
        //                }
        //            }
        //            else if (name == "risk exp date")
        //            {
        //                if (row[i].ToString() != "" && row[i].ToString() != null)
        //                {
        //                    string strDate = row[i].ToString();
        //                    DateTime date = Convert.ToDateTime(strDate, CultureInfo.InvariantCulture);
        //                    dr[j] = date.ToString("yyyy-MM-dd");
        //                }
        //                else
        //                {
        //                    dr[j] = row[i].ToString();
        //                }
        //            }
        //            else if (name == "od premium")
        //            {
        //                if (row[i].ToString() == "" || row[i].ToString() == null)
        //                {
        //                    dr[j] = "0";
        //                }
        //                else
        //                {
        //                    dr[j] = row[i].ToString();
        //                }
        //            }
        //            else if (name == "irda od amt")
        //            {
        //                if (row[i].ToString() == "" || row[i].ToString() == null)
        //                {
        //                    dr[j] = "0";
        //                }
        //                else
        //                {
        //                    dr[j] = row[i].ToString();
        //                }
        //            }
        //            else if (name == "irda od %")
        //            {
        //                if (row[i].ToString() == "" || row[i].ToString() == null)
        //                {
        //                    dr[j] = "0";
        //                }
        //                else
        //                {
        //                    dr[j] = row[i].ToString();
        //                }
        //            }
        //            else if (name == "tp premium")
        //            {
        //                if (row[i].ToString() == "" || row[i].ToString() == null)
        //                {
        //                    dr[j] = "0";
        //                }
        //                else
        //                {
        //                    dr[j] = row[i].ToString();
        //                }
        //            }
        //            else if (name == "terrorism_premium")
        //            {
        //                if (row[i].ToString() == "" || row[i].ToString() == null)
        //                {
        //                    dr[j] = "0";
        //                }
        //                else
        //                {
        //                    dr[j] = row[i].ToString();
        //                }
        //            }
        //            else if (name == "reward od %")
        //            {
        //                if (row[i].ToString() == "" || row[i].ToString() == null)
        //                {
        //                    dr[j] = "0";
        //                }
        //                else
        //                {
        //                    dr[j] = row[i].ToString();
        //                }
        //            }
        //            else if (name == "reward tp%")
        //            {
        //                if (row[i].ToString() == "" || row[i].ToString() == null)
        //                {
        //                    dr[j] = "0";
        //                }
        //                else
        //                {
        //                    dr[j] = row[i].ToString();
        //                }
        //            }
        //            else if (name == "irda reward amt")
        //            {
        //                if (row[i].ToString() == "" || row[i].ToString() == null)
        //                {
        //                    dr[j] = "0";
        //                }
        //                else
        //                {
        //                    dr[j] = row[i].ToString();
        //                }
        //            }
        //            else if (name == "product name")
        //            {
        //                if (row[i].ToString() == "" || row[i].ToString() == null)
        //                {
        //                    dr[j] = "0";
        //                }
        //                else
        //                {
        //                    dr[j] = row[i].ToString();
        //                }
        //            }
        //            //else
        //            //{
        //            //    dr[i] = row[i].ToString();
        //            //}
        //        }
        //        dtcopy.Rows.Add(dr);
        //    }
        //    return dtcopy;
        //}
        //public void ExceltoDatatable1(string TranID, string RDate, string Insurance, string Salesby, string Serviceby, string location, string Support, string Rmonth)
        //{
        //    SQLProcs sql = new SQLProcs();
        //    using (ExcelEngine excelEngine = new ExcelEngine())
        //    {
        //        IApplication application = excelEngine.Excel;
        //        application.DefaultVersion = Syncfusion.XlsIO.ExcelVersion.Xlsx;
        //        FileStream inputStream = new FileStream(Filepath, FileMode.Open, FileAccess.Read);
        //        Syncfusion.XlsIO.IWorkbook workbook = application.Workbooks.Open(inputStream);
        //        IWorksheet worksheet = workbook.Worksheets[0];

        //        //Read data from the worksheet and export to the DataTable.
        //        DataTable DT = worksheet.ExportDataTable(worksheet.UsedRange, ExcelExportDataTableOptions.ColumnNames);

        //        DataTable RDT = copyDatatable1(DT);

        //        //DataSet ds = new DataSet();

        //        //ds = 
        //        sql.ExecuteSQLNonQuery("SP_ICICITransactions",
        //                           new SqlParameter { ParameterName = "@Imode", Value = 9 },
        //                           new SqlParameter { ParameterName = "@RDate", Value = RDate },
        //                           new SqlParameter { ParameterName = "@Rmonth", Value = Rmonth },
        //                           new SqlParameter { ParameterName = "@TranID", Value = TranID },
        //                           new SqlParameter { ParameterName = "@Insurance", Value = Insurance },
        //                           new SqlParameter { ParameterName = "@Salesby", Value = Salesby },
        //                           new SqlParameter { ParameterName = "@Serviceby", Value = Serviceby },
        //                           new SqlParameter { ParameterName = "@location", Value = location },
        //                           new SqlParameter { ParameterName = "@Support", Value = Support },
        //                           new SqlParameter { ParameterName = "@ICICITransaction", Value = RDT },
        //                           new SqlParameter { ParameterName = "@RFormat", Value = "F1" },
        //                           new SqlParameter { ParameterName = "@InvNo", Value = "ILGI" },
        //                           new SqlParameter { ParameterName = "@ReportId", Value = "ILG1" },
        //                           new SqlParameter { ParameterName = "@DocName", Value = "ICICI General Insurance Co. Ltd." }
        //                           );

        //    }
        //}
        //public DataTable copyDatatable1(DataTable dt)
        //{
        //    DataTable dtcopy = new DataTable();
        //    DataRow dr = null;

        //    foreach (DataColumn column in dt.Columns)
        //    {
        //        if (column.ColumnName.ToUpper() == "CUSTOMER_NAME" || column.ColumnName.ToUpper() == "POL_NUM_TXT" || column.ColumnName.ToUpper() == "POLICY_START_DATE" ||
        //            column.ColumnName.ToUpper() == "POLICY_END_DATE" || column.ColumnName.ToUpper() == "PRODUCT_NAME" ||
        //            column.ColumnName.ToUpper() == "PREMIUM_FOR_PAYOUTS" || column.ColumnName.ToUpper() == "TERRORISM_PREMIUM_AMOUNT" ||
        //            column.ColumnName.ToUpper() == "COMMISSION_PAYOUTS_PERCENTAGE" || column.ColumnName.ToUpper() == "ACTUAL_COMMISSION_AMOUNT")
        //        {
        //            dtcopy.Columns.Add(column.ColumnName);
        //        }
        //    }

        //    foreach (DataRow row in dt.Rows)
        //    {
        //        dr = dtcopy.NewRow();
        //        int j = 0; int k = 0;
        //        for (int i = 0; i < row.ItemArray.Length; i++)
        //        {
        //            string name = dt.Columns[i].ColumnName.ToUpper().ToString();
        //            if (name == "CUSTOMER_NAME" || name == "POL_NUM_TXT" || name == "POLICY_START_DATE" ||
        //            name == "POLICY_END_DATE" || name == "PRODUCT_NAME" || name == "PREMIUM_FOR_PAYOUTS" || name == "TERRORISM_PREMIUM_AMOUNT" ||
        //            name == "COMMISSION_PAYOUTS_PERCENTAGE" || name == "ACTUAL_COMMISSION_AMOUNT")
        //            {
        //                if (k != 0)
        //                {
        //                    j++;
        //                }
        //                k = 1;
        //            }
        //            if (name == "CUSTOMER_NAME")
        //            {
        //                if (row[i].ToString() != "" && row[i].ToString() != null)
        //                {
        //                    dr[j] = row[i].ToString();
        //                }
        //                else
        //                {
        //                    dr[j] = "";
        //                }
        //            }
        //            else if (name == "POL_NUM_TXT")
        //            {
        //                dr[j] = row[i].ToString();
        //            }
        //            else if (name == "POLICY_START_DATE")
        //            {
        //                if (row[i].ToString() != "" && row[i].ToString() != null)
        //                {
        //                    string strDate = row[i].ToString();
        //                    DateTime date = Convert.ToDateTime(strDate, CultureInfo.InvariantCulture);
        //                    dr[j] = date.ToString("yyyy-MM-dd");
        //                }
        //                else
        //                {
        //                    dr[j] = row[i].ToString();
        //                }
        //            }
        //            else if (name == "POLICY_END_DATE")
        //            {
        //                if (row[i].ToString() != "" && row[i].ToString() != null)
        //                {
        //                    string strDate = row[i].ToString();
        //                    DateTime date = Convert.ToDateTime(strDate, CultureInfo.InvariantCulture);
        //                    dr[j] = date.ToString("yyyy-MM-dd");
        //                }
        //                else
        //                {
        //                    dr[j] = row[i].ToString();
        //                }
        //            }
        //            else if (name == "PRODUCT_NAME")
        //            {
        //                dr[j] = row[i].ToString();
        //            }
        //            else if (name == "PREMIUM_FOR_PAYOUTS")
        //            {
        //                if (row[i].ToString() == "" || row[i].ToString() == null)
        //                {
        //                    dr[j] = "0";
        //                }
        //                else
        //                {
        //                    dr[j] = row[i].ToString();
        //                }
        //            }
        //            else if (name == "TERRORISM_PREMIUM_AMOUNT")
        //            {
        //                if (row[i].ToString() == "" || row[i].ToString() == null)
        //                {
        //                    dr[j] = "0";
        //                }
        //                else
        //                {
        //                    dr[j] = row[i].ToString();
        //                }
        //            }
        //            else if (name == "COMMISSION_PAYOUTS_PERCENTAGE")
        //            {
        //                if (row[i].ToString() == "" || row[i].ToString() == null)
        //                {
        //                    dr[j] = "0";
        //                }
        //                else
        //                {
        //                    dr[j] = row[i].ToString();
        //                }
        //            }
        //            else if (name == "ACTUAL_COMMISSION_AMOUNT")
        //            {
        //                if (row[i].ToString() == "" || row[i].ToString() == null)
        //                {
        //                    dr[j] = "0";
        //                }
        //                else
        //                {
        //                    dr[j] = row[i].ToString();
        //                }
        //            }
        //            else if (name == "reward od %")
        //            {
        //                if (row[i].ToString() == "" || row[i].ToString() == null)
        //                {
        //                    dr[j] = "0";
        //                }
        //                else
        //                {
        //                    dr[j] = row[i].ToString();
        //                }
        //            }
        //        }
        //        dtcopy.Rows.Add(dr);
        //    }
        //    return dtcopy;
        //}

        //public static void InsertTransaction(Microsoft.Office.Interop.Excel.Workbook WB, string TranID, string RDate, string Insurance, string Salesby, string Serviceby, string location, string Support, string Rmonth)
        //{
        //    //National insurance
        //    SQLProcs sql = new SQLProcs();
        //    Microsoft.Office.Interop.Excel.Worksheet wks = (Microsoft.Office.Interop.Excel.Worksheet)WB.Worksheets[1];
        //    Microsoft.Office.Interop.Excel.Range lastCell = wks.Cells.SpecialCells(Microsoft.Office.Interop.Excel.XlCellType.xlCellTypeLastCell, Type.Missing);
        //    int lastrow = lastCell.Row; var Terrorism = ""; var Policy_Endorsement = ""; var InsuredType = ""; var Policy_Type = "";
        //    for (int i = 4; i < lastrow; i++)
        //    {
        //        string InsuredName = ((Microsoft.Office.Interop.Excel.Range)wks.Cells[i, 6]).Value;
        //        if (InsuredName != null && InsuredName != "" && InsuredName != " ")
        //        {
        //            InsuredName = InsuredName.Replace("\n", "").TrimStart();
        //            string PolicyNo = ((Microsoft.Office.Interop.Excel.Range)wks.Cells[i, 4]).Value.Replace("\n", "").TrimStart();
        //            //var Endo_Effective_Date = ((Microsoft.Office.Interop.Excel.Range)wks.Cells[i, 7]).Value;
        //            //var Effective_Date = ((Microsoft.Office.Interop.Excel.Range)wks.Cells[i, 5]).Value;
        //            var END_Date = ((Microsoft.Office.Interop.Excel.Range)wks.Cells[i, 5]).Value;
        //            //Endo_Effective_Date = Endo_Effective_Date.ToString("dd/MM/yyyy");
        //            //Effective_Date = Effective_Date.ToString("dd/MM/yyyy");
        //            END_Date = END_Date.ToString("dd/MM/yyyy");
        //            //Policy_Endorsement = ((Microsoft.Office.Interop.Excel.Range)wks.Cells[i, 11]).Value.Replace("\n", "").TrimStart();
        //            var Presult = PolicyNo.Substring(PolicyNo.LastIndexOf('/') + 1);
        //            if (Presult.Contains("0"))
        //            {
        //                Policy_Endorsement = "Policy";
        //            }
        //            else
        //            {
        //                Policy_Endorsement = "Endorsement";
        //            }
        //            var Premium_Amt = Convert.ToString(((Microsoft.Office.Interop.Excel.Range)wks.Cells[i, 13]).Value);
        //            var Revenue_Amt = Convert.ToString(((Microsoft.Office.Interop.Excel.Range)wks.Cells[i, 14]).Value);
        //            var Revenue_Pcnt = Convert.ToString(((Microsoft.Office.Interop.Excel.Range)wks.Cells[i, 8]).Text);
        //            Revenue_Pcnt = Revenue_Pcnt.Replace("\n", "").Replace("%", "").Replace(",", "").TrimStart();
        //            //var Total_Premium = Convert.ToString(((Microsoft.Office.Interop.Excel.Range)wks.Cells[i, 13]).Value);
        //            Terrorism = Convert.ToString(((Microsoft.Office.Interop.Excel.Range)wks.Cells[i, 10]).Value);
        //            var ODOtherPremium = Convert.ToString(((Microsoft.Office.Interop.Excel.Range)wks.Cells[i, 7]).Value).Replace("\n", "").Replace(",", "").TrimStart();
        //            var OD_OtherCommission = Convert.ToString(((Microsoft.Office.Interop.Excel.Range)wks.Cells[i, 9]).Value).Replace("\n", "").Replace(",", "").TrimStart();

        //            var Ptype = ((Microsoft.Office.Interop.Excel.Range)wks.Cells[i, 1]).Value;
        //            if (Ptype != "" && Ptype != null)
        //            {
        //                Policy_Type = Ptype.Replace("\n", "").Replace(",", "").TrimStart();
        //            }

        //            if (InsuredName.Contains("LIMITED") || InsuredName.Contains("LTD") || InsuredName.Contains(".COM"))
        //            {
        //                InsuredType = "Corporate";
        //            }
        //            else
        //            {
        //                InsuredType = "Retail";
        //            }
        //            if (Premium_Amt == "" || Premium_Amt == " " || Terrorism == null)
        //            {
        //                Premium_Amt = 0;
        //            }
        //            if (Revenue_Amt == "" || Revenue_Amt == " " || Terrorism == null)
        //            {
        //                Revenue_Amt = 0;
        //            }
        //            if (Revenue_Pcnt == "" || Revenue_Pcnt == " " || Revenue_Pcnt == null)
        //            {
        //                Revenue_Pcnt = "0";
        //            }
        //            if (Terrorism == "" || Terrorism == " " || Terrorism == null)
        //            {
        //                Terrorism = "0";
        //            }
        //            if (ODOtherPremium == "" || ODOtherPremium == " " || ODOtherPremium == null)
        //            {
        //                ODOtherPremium = "0";
        //            }
        //            if (OD_OtherCommission == "" || OD_OtherCommission == " " || OD_OtherCommission == null)
        //            {
        //                OD_OtherCommission = "0";
        //            }
        //            sql.ExecuteSQLNonQuery("SP_NationalExcel_Transactions",
        //                       new SqlParameter { ParameterName = "@Imode", Value = 1 },
        //                       new SqlParameter { ParameterName = "@RDate", Value = RDate },
        //                       new SqlParameter { ParameterName = "@Rmonth", Value = Rmonth },
        //                       new SqlParameter { ParameterName = "@ClientName", Value = InsuredName },
        //                       new SqlParameter { ParameterName = "@Itype", Value = InsuredType },
        //                       new SqlParameter { ParameterName = "@Policy_No", Value = PolicyNo },
        //                       new SqlParameter { ParameterName = "@Revenue_Pcnt", Value = Revenue_Pcnt },
        //                       //new SqlParameter { ParameterName = "@Endo_Effective_Date", Value = Endo_Effective_Date },
        //                       //new SqlParameter { ParameterName = "@Effective_Date", Value = Effective_Date },
        //                       new SqlParameter { ParameterName = "@END_Date", Value = END_Date },
        //                       new SqlParameter { ParameterName = "@Policy_Type", Value = Policy_Type },
        //                       new SqlParameter { ParameterName = "@Premium_Amt", Value = Premium_Amt },
        //                       new SqlParameter { ParameterName = "@TranID", Value = TranID },
        //                       new SqlParameter { ParameterName = "@Revenue_Amt", Value = Revenue_Amt },
        //                        new SqlParameter { ParameterName = "@ODOtherPremium", Value = ODOtherPremium },
        //                        new SqlParameter { ParameterName = "@OD_OtherCommission", Value = OD_OtherCommission },
        //                       new SqlParameter { ParameterName = "@Terrorism", Value = Terrorism },
        //                       // new SqlParameter { ParameterName = "@Total_Premium", Value = Total_Premium },
        //                       new SqlParameter { ParameterName = "@Insurance", Value = Insurance },
        //                       new SqlParameter { ParameterName = "@Salesby", Value = Salesby },
        //                       new SqlParameter { ParameterName = "@Serviceby", Value = Serviceby },
        //                       new SqlParameter { ParameterName = "@location", Value = location },
        //                       new SqlParameter { ParameterName = "@Support", Value = Support },
        //                       new SqlParameter { ParameterName = "@Policy_Endorsement", Value = Policy_Endorsement },
        //                       new SqlParameter { ParameterName = "@RFormat", Value = "F1" },
        //                       new SqlParameter { ParameterName = "@InvNo", Value = "NACX" },
        //                       new SqlParameter { ParameterName = "@ReportId", Value = "NACX" },
        //                       new SqlParameter { ParameterName = "@DocName", Value = "National Insurance Co. Ltd." }
        //                       );
        //        }
        //    }
        //}

        //public static void InsertTransaction(Microsoft.Office.Interop.Excel.Workbook WB, string TranID, string RDate, string Insurance, string Salesby, string Serviceby, string location, string Support, string Rmonth)
        //{
        //    //Bajaj insurance
        //    SQLProcs sql = new SQLProcs();
        //    Microsoft.Office.Interop.Excel.Worksheet wks = (Microsoft.Office.Interop.Excel.Worksheet)WB.Worksheets[1];
        //    Microsoft.Office.Interop.Excel.Range lastCell = wks.Cells.SpecialCells(Microsoft.Office.Interop.Excel.XlCellType.xlCellTypeLastCell, Type.Missing);
        //    int lastrow = lastCell.Row; var Terrorism = ""; var Policy_Endorsement = ""; var InsuredType = ""; var Policy_Type = "";
        //    for (int i = 2; i < lastrow; i++)
        //    {
        //        string InsuredName = ((Microsoft.Office.Interop.Excel.Range)wks.Cells[i, 26]).Value;
        //        if (InsuredName != null && InsuredName != "" && InsuredName != " ")
        //        {
        //            InsuredName = InsuredName.Replace("\n", "").TrimStart();
        //            string PolicyNo = ((Microsoft.Office.Interop.Excel.Range)wks.Cells[i, 2]).Value.Replace("\n", "").TrimStart();
        //            var Endo_Effective_Date = ((Microsoft.Office.Interop.Excel.Range)wks.Cells[i, 3]).Value;
        //            var Effective_Date = ((Microsoft.Office.Interop.Excel.Range)wks.Cells[i, 35]).Value;
        //            int ENdolen = Convert.ToString(Endo_Effective_Date).Length;
        //            if (ENdolen > 11)
        //            {
        //                Endo_Effective_Date = Endo_Effective_Date.ToString("dd/MM/yyyy");
        //            }
        //            if (Effective_Date != null)
        //            {
        //                int Efflen = Convert.ToString(Effective_Date).Length;
        //                if (Efflen > 11)
        //                {
        //                    Effective_Date = Effective_Date.ToString("dd/MM/yyyy");
        //                }
        //            }
        //            //var END_Date = ((Microsoft.Office.Interop.Excel.Range)wks.Cells[i, 5]).Value;
        //            //Endo_Effective_Date = Endo_Effective_Date.ToString("dd/MM/yyyy");
        //            //Effective_Date = Effective_Date.ToString("dd/MM/yyyy");
        //            //END_Date = END_Date.ToString("dd/MM/yyyy");
        //            //Policy_Endorsement = ((Microsoft.Office.Interop.Excel.Range)wks.Cells[i, 11]).Value.Replace("\n", "").TrimStart();

        //            var Presult = PolicyNo.Substring(PolicyNo.LastIndexOf('-') + 1);
        //            if (PolicyNo.Contains("-"))
        //            {
        //                var resultString = Regex.Match(Presult, @"\d+").Value;
        //                if (Convert.ToDouble(resultString) > 0)
        //                {
        //                    Policy_Endorsement = "Endorsement";
        //                }
        //                else
        //                {
        //                    Policy_Endorsement = "Policy";
        //                }
        //            }
        //            var Premium_Amt = Convert.ToString(((Microsoft.Office.Interop.Excel.Range)wks.Cells[i, 9]).Value);
        //            var Revenue_Amt = Convert.ToString(((Microsoft.Office.Interop.Excel.Range)wks.Cells[i, 15]).Value);
        //            var Revenue_Pcnt = Convert.ToString(((Microsoft.Office.Interop.Excel.Range)wks.Cells[i, 5]).Text);
        //            Revenue_Pcnt = Revenue_Pcnt.Replace("\n", "").Replace("%", "").Replace(",", "").TrimStart();
        //            var TP1 = Convert.ToString(((Microsoft.Office.Interop.Excel.Range)wks.Cells[i, 11]).Value);
        //            var TP = Convert.ToString(((Microsoft.Office.Interop.Excel.Range)wks.Cells[i, 10]).Value);

        //            Policy_Type = Convert.ToString(((Microsoft.Office.Interop.Excel.Range)wks.Cells[i, 19]).Value);

        //            if (InsuredName.Contains("LIMITED") || InsuredName.Contains("LTD") || InsuredName.Contains(".COM"))
        //            {
        //                InsuredType = "Corporate";
        //            }
        //            else
        //            {
        //                InsuredType = "Retail";
        //            }
        //            if (Premium_Amt == "" || Premium_Amt == " " || Premium_Amt == null)
        //            {
        //                Premium_Amt = 0;
        //            }
        //            if (Revenue_Amt == "" || Revenue_Amt == " " || Revenue_Amt == null)
        //            {
        //                Revenue_Amt = 0;
        //            }
        //            if (Revenue_Pcnt == "" || Revenue_Pcnt == " " || Revenue_Pcnt == null)
        //            {
        //                Revenue_Pcnt = "0";
        //            }
        //            if (TP1 == "" || TP1 == " " || TP1 == null)
        //            {
        //                TP1 = "0";
        //            }
        //            if (TP == "" || TP == " " || TP == null)
        //            {
        //                TP = "0";
        //            }
        //            if(TP == "0" && TP1 != "0")
        //            {
        //                Terrorism = TP1;
        //            }
        //            else if(TP != "0" && TP1 == "0")
        //            {
        //                Terrorism = TP;
        //            }
        //            if (Terrorism == "" || Terrorism == " " || Terrorism == null)
        //            {
        //                Terrorism = "0";
        //            }
        //            sql.ExecuteSQLNonQuery("SP_BajajTransactions",
        //                       new SqlParameter { ParameterName = "@Imode", Value = 1 },
        //                       new SqlParameter { ParameterName = "@RDate", Value = RDate },
        //                       new SqlParameter { ParameterName = "@Rmonth", Value = Rmonth },
        //                       new SqlParameter { ParameterName = "@ClientName", Value = InsuredName },
        //                       new SqlParameter { ParameterName = "@Itype", Value = InsuredType },
        //                       new SqlParameter { ParameterName = "@Policy_No", Value = PolicyNo },
        //                       new SqlParameter { ParameterName = "@Revenue_Pcnt", Value = Revenue_Pcnt },
        //                       new SqlParameter { ParameterName = "@Endo_Effective_Date", Value = Endo_Effective_Date },
        //                       new SqlParameter { ParameterName = "@Effective_Date", Value = Effective_Date },
        //                       //new SqlParameter { ParameterName = "@END_Date", Value = END_Date },
        //                       new SqlParameter { ParameterName = "@Policy_Type", Value = Policy_Type },
        //                       new SqlParameter { ParameterName = "@Premium_Amt", Value = Premium_Amt },
        //                       new SqlParameter { ParameterName = "@TranID", Value = TranID },
        //                       new SqlParameter { ParameterName = "@Revenue_Amt", Value = Revenue_Amt },
        //                       new SqlParameter { ParameterName = "@Terrorism", Value = Terrorism },
        //                       new SqlParameter { ParameterName = "@Insurance", Value = Insurance },
        //                       new SqlParameter { ParameterName = "@Salesby", Value = Salesby },
        //                       new SqlParameter { ParameterName = "@Serviceby", Value = Serviceby },
        //                       new SqlParameter { ParameterName = "@location", Value = location },
        //                       new SqlParameter { ParameterName = "@Support", Value = Support },
        //                       new SqlParameter { ParameterName = "@Policy_Endorsement", Value = Policy_Endorsement },
        //                       new SqlParameter { ParameterName = "@RFormat", Value = "F1" },
        //                       new SqlParameter { ParameterName = "@InvNo", Value = "BAGX" },
        //                       new SqlParameter { ParameterName = "@ReportId", Value = "BAGX" },
        //                       new SqlParameter { ParameterName = "@DocName", Value = "Bajaj Allianz General Insurance Co. Ltd." }
        //                       );
        //        }
        //    }
        //}

        //public static void InsertTransaction(Microsoft.Office.Interop.Excel.Workbook WB, string TranID, string RDate, string Insurance, string Salesby, string Serviceby, string location, string Support, string Rmonth)
        //{
        //    //Care Health insurance
        //    SQLProcs sql = new SQLProcs();
        //    Microsoft.Office.Interop.Excel.Worksheet wks = (Microsoft.Office.Interop.Excel.Worksheet)WB.Worksheets[1];
        //    Microsoft.Office.Interop.Excel.Range lastCell = wks.Cells.SpecialCells(Microsoft.Office.Interop.Excel.XlCellType.xlCellTypeLastCell, Type.Missing);
        //    int lastrow = lastCell.Row; var Terrorism = ""; var Policy_Endorsement = ""; var InsuredType = ""; var Policy_Type = "";var New_Renewal = "";
        //    for (int i = 2; i < lastrow; i++)
        //    {
        //        string InsuredName = ((Microsoft.Office.Interop.Excel.Range)wks.Cells[i, 2]).Value;
        //        if (InsuredName != null && InsuredName != "" && InsuredName != " ")
        //        {
        //            InsuredName = InsuredName.Replace("\n", "").TrimStart();
        //            string PolicyNo = ((Microsoft.Office.Interop.Excel.Range)wks.Cells[i, 1]).Value.Replace("\n", "").TrimStart();
        //            var EndosNo = ((Microsoft.Office.Interop.Excel.Range)wks.Cells[i, 11]).Value;
        //            if(EndosNo != null && EndosNo != "")
        //            {
        //                PolicyNo = PolicyNo + "/" + EndosNo.Replace("\n", "").TrimStart();
        //                Policy_Endorsement = "Endorsement";
        //            }
        //            else
        //            {
        //                PolicyNo = PolicyNo + "/0";
        //                Policy_Endorsement = "Policy";
        //            }
        //            var Client_N_E = ((Microsoft.Office.Interop.Excel.Range)wks.Cells[i, 5]).Value.Replace("\n", "").TrimStart();
        //            if(Client_N_E == "New Business")
        //            {
        //                Client_N_E = "New Client";
        //                New_Renewal = "New Policy";
        //            }
        //            else
        //            {
        //                Client_N_E = "Existing Client";
        //                New_Renewal = "Renewal Policy";
        //            }
        //            var Endo_Effective_Date = ((Microsoft.Office.Interop.Excel.Range)wks.Cells[i, 7]).Value;
        //            var Effective_Date = ((Microsoft.Office.Interop.Excel.Range)wks.Cells[i, 12]).Value;
        //            var END_Date = ((Microsoft.Office.Interop.Excel.Range)wks.Cells[i, 13]).Value;
        //            int ENdolen = Convert.ToString(Endo_Effective_Date).Length;
        //            if (ENdolen > 11)
        //            {
        //                Endo_Effective_Date = Endo_Effective_Date.ToString("dd/MM/yyyy");
        //            }
        //            int Efflen = Convert.ToString(Effective_Date).Length;
        //            if (Efflen > 11)
        //            {
        //                Effective_Date = Effective_Date.ToString("dd/MM/yyyy");
        //            }
        //            int ENDlen = Convert.ToString(END_Date).Length;
        //            if (ENDlen > 11)
        //            {
        //                END_Date = END_Date.ToString("dd/MM/yyyy");
        //            }
        //            //var END_Date = ((Microsoft.Office.Interop.Excel.Range)wks.Cells[i, 5]).Value;
        //            //Endo_Effective_Date = Endo_Effective_Date.ToString("dd/MM/yyyy");
        //            //Effective_Date = Effective_Date.ToString("dd/MM/yyyy");
        //            //END_Date = END_Date.ToString("dd/MM/yyyy");
        //            //Policy_Endorsement = ((Microsoft.Office.Interop.Excel.Range)wks.Cells[i, 11]).Value.Replace("\n", "").TrimStart();

        //            var Premium_Amt = Convert.ToString(((Microsoft.Office.Interop.Excel.Range)wks.Cells[i, 3]).Value);
        //            var Revenue_Amt = Convert.ToString(((Microsoft.Office.Interop.Excel.Range)wks.Cells[i, 9]).Value);
        //            var Revenue_Pcnt = Convert.ToString(((Microsoft.Office.Interop.Excel.Range)wks.Cells[i, 6]).Text);
        //            Revenue_Pcnt = Revenue_Pcnt.Replace("\n", "").Replace("%", "").Replace(",", "").TrimStart();

        //            Policy_Type = Convert.ToString(((Microsoft.Office.Interop.Excel.Range)wks.Cells[i, 4]).Value);

        //            if (InsuredName.Contains("LIMITED") || InsuredName.Contains("LTD") || InsuredName.Contains(".COM"))
        //            {
        //                InsuredType = "Corporate";
        //            }
        //            else
        //            {
        //                InsuredType = "Retail";
        //            }
        //            if (Premium_Amt == "" || Premium_Amt == " " || Premium_Amt == null)
        //            {
        //                Premium_Amt = 0;
        //            }
        //            if (Revenue_Amt == "" || Revenue_Amt == " " || Revenue_Amt == null)
        //            {
        //                Revenue_Amt = 0;
        //            }
        //            if (Revenue_Pcnt == "" || Revenue_Pcnt == " " || Revenue_Pcnt == null)
        //            {
        //                Revenue_Pcnt = "0";
        //            }
        //            if (Terrorism == "" || Terrorism == " " || Terrorism == null)
        //            {
        //                Terrorism = "0";
        //            }
        //            sql.ExecuteSQLNonQuery("SP_CareHealthTransactions",
        //                       new SqlParameter { ParameterName = "@Imode", Value = 1 },
        //                       new SqlParameter { ParameterName = "@RDate", Value = RDate },
        //                       new SqlParameter { ParameterName = "@Rmonth", Value = Rmonth },
        //                       new SqlParameter { ParameterName = "@ClientName", Value = InsuredName },
        //                       new SqlParameter { ParameterName = "@Itype", Value = InsuredType },
        //                       new SqlParameter { ParameterName = "@Client_N_E", Value = Client_N_E },
        //                       new SqlParameter { ParameterName = "@New_Renewal", Value = New_Renewal },
        //                       new SqlParameter { ParameterName = "@Policy_No", Value = PolicyNo },
        //                       new SqlParameter { ParameterName = "@Revenue_Pcnt", Value = Revenue_Pcnt },
        //                       new SqlParameter { ParameterName = "@Endo_Effective_Date", Value = Endo_Effective_Date },
        //                       new SqlParameter { ParameterName = "@Effective_Date", Value = Effective_Date },
        //                       new SqlParameter { ParameterName = "@END_Date", Value = END_Date },
        //                       new SqlParameter { ParameterName = "@Policy_Type", Value = Policy_Type },
        //                       new SqlParameter { ParameterName = "@Premium_Amt", Value = Premium_Amt },
        //                       new SqlParameter { ParameterName = "@TranID", Value = TranID },
        //                       new SqlParameter { ParameterName = "@Revenue_Amt", Value = Revenue_Amt },
        //                       new SqlParameter { ParameterName = "@Terrorism", Value = Terrorism },
        //                       new SqlParameter { ParameterName = "@Insurance", Value = Insurance },
        //                       new SqlParameter { ParameterName = "@Salesby", Value = Salesby },
        //                       new SqlParameter { ParameterName = "@Serviceby", Value = Serviceby },
        //                       new SqlParameter { ParameterName = "@location", Value = location },
        //                       new SqlParameter { ParameterName = "@Support", Value = Support },
        //                       new SqlParameter { ParameterName = "@Policy_Endorsement", Value = Policy_Endorsement },
        //                       new SqlParameter { ParameterName = "@RFormat", Value = "F1" },
        //                       new SqlParameter { ParameterName = "@InvNo", Value = "CHIX" },
        //                       new SqlParameter { ParameterName = "@ReportId", Value = "CHIX" },
        //                       new SqlParameter { ParameterName = "@DocName", Value = "Care Health Insurance Co. Ltd." }
        //                       );
        //        }
        //    }
        //}

        //public static void InsertTransaction(Microsoft.Office.Interop.Excel.Workbook WB, string TranID, string RDate, string Insurance, string Salesby, string Serviceby, string location, string Support, string Rmonth)
        //{
        //    //Chola MS General Insurance
        //    SQLProcs sql = new SQLProcs();
        //    Microsoft.Office.Interop.Excel.Worksheet wks = (Microsoft.Office.Interop.Excel.Worksheet)WB.Worksheets[1];
        //    Microsoft.Office.Interop.Excel.Range lastCell = wks.Cells.SpecialCells(Microsoft.Office.Interop.Excel.XlCellType.xlCellTypeLastCell, Type.Missing);
        //    int lastrow = lastCell.Row; var Terrorism = ""; var Policy_Endorsement = ""; var InsuredType = ""; var Policy_Type = ""; var Reward_Pcnt = ""; var Reward_Amt = "";
        //    for (int i = 2; i < lastrow; i++)
        //    {
        //        string InsuredName = ((Microsoft.Office.Interop.Excel.Range)wks.Cells[i, 11]).Value;
        //        if (InsuredName != null && InsuredName != "" && InsuredName != " ")
        //        {
        //            InsuredName = InsuredName.Replace("\n", "").TrimStart();
        //            InsuredName = Regex.Match(InsuredName, @"\D+").Value;
        //            var Itype = ((Microsoft.Office.Interop.Excel.Range)wks.Cells[i, 10]).Value;
        //            var Itype1 = ((Microsoft.Office.Interop.Excel.Range)wks.Cells[i, 47]).Value;
        //            if ((Itype == null || Itype == "") && (Itype1 != null && Itype1 != ""))
        //            {
        //                InsuredType = Itype1;
        //            }
        //            else if ((Itype1 == null || Itype1 == "") && (Itype != null && Itype != ""))
        //            {
        //                InsuredType = Itype;
        //            }
        //            else
        //            {
        //                InsuredType = Itype;
        //            }
        //            string PolicyNo = ((Microsoft.Office.Interop.Excel.Range)wks.Cells[i, 5]).Value.Replace("\n", "").TrimStart();
        //            var Presult = PolicyNo.Substring(PolicyNo.LastIndexOf('/') + 1);
        //            if (PolicyNo.Contains("/"))
        //            {
        //                var resultString = Regex.Match(Presult, @"\d+").Value;
        //                if (Convert.ToDouble(resultString) > 0)
        //                {
        //                    Policy_Endorsement = "Endorsement";
        //                }
        //                else
        //                {
        //                    Policy_Endorsement = "Policy";
        //                }
        //            }

        //            var Endo_Effective_Date = ((Microsoft.Office.Interop.Excel.Range)wks.Cells[i, 4]).Value;
        //            //var Effective_Date = ((Microsoft.Office.Interop.Excel.Range)wks.Cells[i, 12]).Value;
        //            var END_Date = ((Microsoft.Office.Interop.Excel.Range)wks.Cells[i, 43]).Value;
        //            int ENdolen = Convert.ToString(Endo_Effective_Date).Length;
        //            if (ENdolen > 11)
        //            {
        //                Endo_Effective_Date = Endo_Effective_Date.ToString("dd/MM/yyyy");
        //            }
        //            //int Efflen = Convert.ToString(Effective_Date).Length;
        //            //if (Efflen > 11)
        //            //{
        //            //    Effective_Date = Effective_Date.ToString("dd/MM/yyyy");
        //            //}
        //            int ENDlen = Convert.ToString(END_Date).Length;
        //            if (ENDlen > 11)
        //            {
        //                END_Date = END_Date.ToString("dd/MM/yyyy");
        //            }
        //            //var END_Date = ((Microsoft.Office.Interop.Excel.Range)wks.Cells[i, 5]).Value;
        //            //Endo_Effective_Date = Endo_Effective_Date.ToString("dd/MM/yyyy");
        //            //Effective_Date = Effective_Date.ToString("dd/MM/yyyy");
        //            //END_Date = END_Date.ToString("dd/MM/yyyy");
        //            //Policy_Endorsement = ((Microsoft.Office.Interop.Excel.Range)wks.Cells[i, 11]).Value.Replace("\n", "").TrimStart();

        //            var Premium_Amt = Convert.ToString(((Microsoft.Office.Interop.Excel.Range)wks.Cells[i, 19]).Value);
        //            var Revenue_Amt = Convert.ToString(((Microsoft.Office.Interop.Excel.Range)wks.Cells[i, 21]).Value);
        //            var TPRevenue_Amt = Convert.ToString(((Microsoft.Office.Interop.Excel.Range)wks.Cells[i, 46]).Value);
        //            var Revenue_Pcnt = Convert.ToString(((Microsoft.Office.Interop.Excel.Range)wks.Cells[i, 20]).Text);
        //            Revenue_Pcnt = Revenue_Pcnt.Replace("\n", "").Replace("%", "").Replace(",", "").TrimStart();
        //            Terrorism = Convert.ToString(((Microsoft.Office.Interop.Excel.Range)wks.Cells[i, 45]).Value);

        //            var Reward_Pcnt1 = Convert.ToString(((Microsoft.Office.Interop.Excel.Range)wks.Cells[i, 22]).Value);
        //            var Reward_Pcnt2 = Convert.ToString(((Microsoft.Office.Interop.Excel.Range)wks.Cells[i, 24]).Value);
        //            var Reward_Pcnt3 = Convert.ToString(((Microsoft.Office.Interop.Excel.Range)wks.Cells[i, 26]).Value);

        //            var Reward_Amt1 = Convert.ToString(((Microsoft.Office.Interop.Excel.Range)wks.Cells[i, 23]).Value);
        //            var Reward_Amt2 = Convert.ToString(((Microsoft.Office.Interop.Excel.Range)wks.Cells[i, 25]).Value);
        //            var Reward_Amt3 = Convert.ToString(((Microsoft.Office.Interop.Excel.Range)wks.Cells[i, 27]).Value);

        //            if (Reward_Pcnt1 == "" || Reward_Pcnt1 == " " || Reward_Pcnt1 == null)
        //            {
        //                Reward_Pcnt1 = 0;
        //            }
        //            if (Reward_Pcnt2 == "" || Reward_Pcnt2 == " " || Reward_Pcnt2 == null)
        //            {
        //                Reward_Pcnt2 = 0;
        //            }
        //            if (Reward_Pcnt3 == "" || Reward_Pcnt3 == " " || Reward_Pcnt3 == null)
        //            {
        //                Reward_Pcnt3 = 0;
        //            }
        //            if (Reward_Amt1 == "" || Reward_Amt1 == " " || Reward_Amt1 == null)
        //            {
        //                Reward_Amt1 = 0;
        //            }
        //            if (Reward_Amt2 == "" || Reward_Amt2 == " " || Reward_Amt2 == null)
        //            {
        //                Reward_Amt2 = 0;
        //            }
        //            if (Reward_Amt3 == "" || Reward_Amt3 == " " || Reward_Amt3 == null)
        //            {
        //                Reward_Amt3 = 0;
        //            }

        //            Reward_Pcnt = Convert.ToString(Convert.ToDouble(Reward_Pcnt1) + Convert.ToDouble(Reward_Pcnt2) + Convert.ToDouble(Reward_Pcnt3));
        //            Reward_Amt = Convert.ToString(Convert.ToDouble(Reward_Amt1) + Convert.ToDouble(Reward_Amt2) + Convert.ToDouble(Reward_Amt3));

        //            Policy_Type = Convert.ToString(((Microsoft.Office.Interop.Excel.Range)wks.Cells[i, 6]).Value);

        //            if (InsuredName.Contains("LIMITED") || InsuredName.Contains("LTD") || InsuredName.Contains(".COM"))
        //            {
        //                InsuredType = "Corporate";
        //            }
        //            else
        //            {
        //                InsuredType = "Retail";
        //            }
        //            if (Premium_Amt == "" || Premium_Amt == " " || Premium_Amt == null)
        //            {
        //                Premium_Amt = 0;
        //            }
        //            if (Revenue_Amt == "" || Revenue_Amt == " " || Revenue_Amt == null)
        //            {
        //                Revenue_Amt = 0;
        //            }
        //            if (Revenue_Pcnt == "" || Revenue_Pcnt == " " || Revenue_Pcnt == null)
        //            {
        //                Revenue_Pcnt = "0";
        //            }
        //            if (Terrorism == "" || Terrorism == " " || Terrorism == null)
        //            {
        //                Terrorism = "0";
        //            }
        //            if (TPRevenue_Amt == "" || TPRevenue_Amt == " " || TPRevenue_Amt == null)
        //            {
        //                TPRevenue_Amt = "0";
        //            }
        //            sql.ExecuteSQLNonQuery("SP_CholaTransactions",
        //                       new SqlParameter { ParameterName = "@Imode", Value = 1 },
        //                       new SqlParameter { ParameterName = "@RDate", Value = RDate },
        //                       new SqlParameter { ParameterName = "@Rmonth", Value = Rmonth },
        //                       new SqlParameter { ParameterName = "@ClientName", Value = InsuredName },
        //                       new SqlParameter { ParameterName = "@Itype", Value = InsuredType },
        //                       new SqlParameter { ParameterName = "@Policy_No", Value = PolicyNo },
        //                       new SqlParameter { ParameterName = "@Revenue_Pcnt", Value = Revenue_Pcnt },
        //                       new SqlParameter { ParameterName = "@Endo_Effective_Date", Value = Endo_Effective_Date },
        //                       //new SqlParameter { ParameterName = "@Effective_Date", Value = Effective_Date },
        //                       new SqlParameter { ParameterName = "@END_Date", Value = END_Date },
        //                       new SqlParameter { ParameterName = "@Policy_Type", Value = Policy_Type },
        //                       new SqlParameter { ParameterName = "@Premium_Amt", Value = Premium_Amt },
        //                       new SqlParameter { ParameterName = "@TPRevenue_Amt", Value = TPRevenue_Amt },
        //                       new SqlParameter { ParameterName = "@Reward_Pcnt", Value = Reward_Pcnt },
        //                       new SqlParameter { ParameterName = "@Reward_Amt", Value = Reward_Amt },
        //                       new SqlParameter { ParameterName = "@TranID", Value = TranID },
        //                       new SqlParameter { ParameterName = "@Revenue_Amt", Value = Revenue_Amt },
        //                       new SqlParameter { ParameterName = "@Terrorism", Value = Terrorism },
        //                       new SqlParameter { ParameterName = "@Insurance", Value = Insurance },
        //                       new SqlParameter { ParameterName = "@Salesby", Value = Salesby },
        //                       new SqlParameter { ParameterName = "@Serviceby", Value = Serviceby },
        //                       new SqlParameter { ParameterName = "@location", Value = location },
        //                       new SqlParameter { ParameterName = "@Support", Value = Support },
        //                       new SqlParameter { ParameterName = "@Policy_Endorsement", Value = Policy_Endorsement },
        //                       new SqlParameter { ParameterName = "@RFormat", Value = "F1" },
        //                       new SqlParameter { ParameterName = "@InvNo", Value = "CMGX" },
        //                       new SqlParameter { ParameterName = "@ReportId", Value = "CMGX" },
        //                       new SqlParameter { ParameterName = "@DocName", Value = "Chola MS General Insurance Co. Ltd." }
        //                       );
        //        }
        //    }
        //}

        //public static void InsertTransaction(Microsoft.Office.Interop.Excel.Workbook WB, string TranID, string RDate, string Insurance, string Salesby, string Serviceby, string location, string Support, string Rmonth)
        //{

        //    SQLProcs sql = new SQLProcs();
        //    Microsoft.Office.Interop.Excel.Worksheet wks = (Microsoft.Office.Interop.Excel.Worksheet)WB.Worksheets[1];
        //    Microsoft.Office.Interop.Excel.Range lastCell = wks.Cells.SpecialCells(Microsoft.Office.Interop.Excel.XlCellType.xlCellTypeLastCell, Type.Missing);
        //    int lastrow = lastCell.Row; var Terrorism = ""; var Policy_Endorsement = ""; var InsuredType = ""; var location1 = "";
        //    for (int i = 2; i <= lastrow; i++)
        //    {
        //        string InsuredName = ((Microsoft.Office.Interop.Excel.Range)wks.Cells[i, 9]).Value.Replace("\n", "").TrimStart();
        //        if (InsuredName == null || InsuredName == "")
        //        {
        //            InsuredName = ((Microsoft.Office.Interop.Excel.Range)wks.Cells[i, 9]).Value.Replace("\n", "").TrimStart();
        //        }
        //        string PolicyNo = ((Microsoft.Office.Interop.Excel.Range)wks.Cells[i, 6]).Value.Replace("\n", "").TrimStart();
        //        if (PolicyNo == null || PolicyNo == "")
        //        {
        //            PolicyNo = ((Microsoft.Office.Interop.Excel.Range)wks.Cells[i, 5]).Value.Replace("\n", "").TrimStart();
        //        }
        //        if (location1 == null || location1 == "")
        //        {
        //            location1 = ((Microsoft.Office.Interop.Excel.Range)wks.Cells[i, 15]).Value.Replace("\n", "").TrimStart();
        //        }

        //        var bustype = ((Microsoft.Office.Interop.Excel.Range)wks.Cells[i, 4]).Value;


        //        //var Endo_Effective_Date = Convert.ToString(((Microsoft.Office.Interop.Excel.Range)wks.Cells[i, 4]).Value);//.Replace("\n", "").TrimStart();
        //        //string END_Date = Convert.ToString(((Microsoft.Office.Interop.Excel.Range)wks.Cells[i, 6]).Value);//.Replace("\n", "").TrimStart();
        //        //var Endo_Effective_Date = ((Microsoft.Office.Interop.Excel.Range)wks.Cells[i, 27]).Value;
        //        var PE = ((Microsoft.Office.Interop.Excel.Range)wks.Cells[i, 6]).Value;
        //        //var END_Date = ((Microsoft.Office.Interop.Excel.Range)wks.Cells[i, 28]).Value;
        //        //Endo_Effective_Date = Endo_Effective_Date.ToString("dd/MM/yyyy");
        //        //END_Date = END_Date.ToString("dd/MM/yyyy");
        //        var Premium_Amt = Convert.ToString(((Microsoft.Office.Interop.Excel.Range)wks.Cells[i, 12]).Value);
        //        var Revenue_Amt = Convert.ToString(((Microsoft.Office.Interop.Excel.Range)wks.Cells[i, 14]).Value);
        //        var Revenue_Pct = Convert.ToString(((Microsoft.Office.Interop.Excel.Range)wks.Cells[i, 13]).Text);
        //        Revenue_Pct = Revenue_Pct.Replace("%", "").TrimStart();
        //        var TP_Amt = "";// Convert.ToString(((Microsoft.Office.Interop.Excel.Range)wks.Cells[i, 32]).Value);
        //                        //Terrorism = Convert.ToString(((Microsoft.Office.Interop.Excel.Range)wks.Cells[i, 42]).Value);
        //                        //var RewardOD_Pct = Convert.ToString(((Microsoft.Office.Interop.Excel.Range)wks.Cells[i, 46]).Value);
        //                        //var RewardTP_Pct = Convert.ToString(((Microsoft.Office.Interop.Excel.Range)wks.Cells[i, 46]).Value);
        //                        // var IRDARewardAmt = Convert.ToString(((Microsoft.Office.Interop.Excel.Range)wks.Cells[i, 48]).Value);
        //        var Policy_Type = ((Microsoft.Office.Interop.Excel.Range)wks.Cells[i, 3]).Value.Replace("\n", "").Replace(",", "").TrimStart();
        //        if (bustype.Contains("Group"))
        //        {
        //            InsuredType = "corporate";
        //        }
        //        else
        //        {
        //            InsuredType = "Retail";
        //        }
        //        if (Premium_Amt == "" || Premium_Amt == " ")
        //        {
        //            Premium_Amt = 0;
        //        }
        //        if (Revenue_Amt == "" || Revenue_Amt == " ")
        //        {
        //            Revenue_Amt = 0;
        //        }
        //        if (Terrorism == "" || Terrorism == " " || Terrorism == null)
        //        {
        //            Terrorism = "0";
        //        }
        //        if (PE == "" || PE == null)
        //        {
        //            Policy_Endorsement = "Policy";
        //        }
        //        else
        //        {
        //            Policy_Endorsement = "Endoresement";
        //        }
        //        if (TP_Amt == "" || TP_Amt == " " || Terrorism == null)
        //        {
        //            TP_Amt = "0";
        //        }
        //        if (Revenue_Pct == "" || Revenue_Pct == " ")
        //        {
        //            Revenue_Pct = "0";
        //        }
        //        //if (RewardOD_Pct == "" || RewardOD_Pct == " " || Terrorism == null)
        //        //{
        //        //    RewardOD_Pct = "0";
        //        //}
        //        //if (RewardTP_Pct == "" || RewardTP_Pct == " " || Terrorism == null)
        //        //{
        //        //    RewardTP_Pct = "0";
        //        //}
        //        //if (IRDARewardAmt == "" || IRDARewardAmt == " " || Terrorism == null)
        //        //{
        //        //    IRDARewardAmt = "0";
        //        //}
        //        sql.ExecuteSQLNonQuery("SP_AdityaBirlaTransactions",
        //                   new SqlParameter { ParameterName = "@Imode", Value = 1 },
        //                   new SqlParameter { ParameterName = "@RDate", Value = RDate },
        //                   new SqlParameter { ParameterName = "@Rmonth", Value = Rmonth },
        //                   new SqlParameter { ParameterName = "@ClientName", Value = InsuredName },
        //                   new SqlParameter { ParameterName = "@Itype", Value = InsuredType },
        //                   new SqlParameter { ParameterName = "@Policy_No", Value = PolicyNo },
        //                   new SqlParameter { ParameterName = "@Endo_Effective_Date", Value = null },
        //                   new SqlParameter { ParameterName = "@END_Date", Value = null },
        //                   new SqlParameter { ParameterName = "@Policy_Type", Value = Policy_Type },
        //                   new SqlParameter { ParameterName = "@Premium_Amt", Value = Premium_Amt },
        //                   new SqlParameter { ParameterName = "@TranID", Value = TranID },
        //                   new SqlParameter { ParameterName = "@Revenue_Amt", Value = Revenue_Amt },
        //                   new SqlParameter { ParameterName = "@TP_Amt", Value = TP_Amt },
        //                   new SqlParameter { ParameterName = "@Terrorism", Value = Terrorism },
        //                   new SqlParameter { ParameterName = "@Effective_Date", Value = null },
        //                   new SqlParameter { ParameterName = "@Revenue_Pct", Value = Revenue_Pct },
        //                   new SqlParameter { ParameterName = "@Insurance", Value = Insurance },
        //                   new SqlParameter { ParameterName = "@Salesby", Value = Salesby },
        //                   new SqlParameter { ParameterName = "@Serviceby", Value = Serviceby },
        //                   new SqlParameter { ParameterName = "@location", Value = location1 },
        //                   new SqlParameter { ParameterName = "@Support", Value = Support },
        //                   new SqlParameter { ParameterName = "@Policy_Endorsement", Value = Policy_Endorsement },
        //                   new SqlParameter { ParameterName = "@RFormat", Value = "F1" },
        //                   new SqlParameter { ParameterName = "@InvNo", Value = "ABHI" },
        //                   new SqlParameter { ParameterName = "@ReportId", Value = "ABHX" },
        //                   new SqlParameter { ParameterName = "@DocName", Value = "Aditya Birla Health Insurance Co. Ltd." }
        //                   );


        //    }

        //}

        //public static void InsertTransaction(Microsoft.Office.Interop.Excel.Workbook WB, string TranID, string RDate, string Insurance, string Salesby, string Serviceby, string location, string Support, string Rmonth)
        //{
        //    //Edelweiss General Insurance Co. Ltd.
        //    SQLProcs sql = new SQLProcs();
        //    Microsoft.Office.Interop.Excel.Worksheet wks = (Microsoft.Office.Interop.Excel.Worksheet)WB.Worksheets[1];
        //    Microsoft.Office.Interop.Excel.Range lastCell = wks.Cells.SpecialCells(Microsoft.Office.Interop.Excel.XlCellType.xlCellTypeLastCell, Type.Missing);
        //    int lastrow = lastCell.Row; var Terrorism = ""; var Policy_Endorsement = ""; var InsuredType = ""; var Policy_Type = ""; var New_Renewal = "";
        //    for (int i = 2; i < lastrow; i++)
        //    {
        //        string InsuredName = ((Microsoft.Office.Interop.Excel.Range)wks.Cells[i, 7]).Value;
        //        if (InsuredName != null && InsuredName != "" && InsuredName != " ")
        //        {
        //            InsuredName = InsuredName.Replace("\n", "").TrimStart();
        //            string PolicyNo = Convert.ToString(((Microsoft.Office.Interop.Excel.Range)wks.Cells[i, 9]).Value).Replace("\n", "").TrimStart();
        //            var Client_N_E = ((Microsoft.Office.Interop.Excel.Range)wks.Cells[i, 3]).Value.Replace("\n", "").TrimStart();
        //            Policy_Endorsement = ((Microsoft.Office.Interop.Excel.Range)wks.Cells[i, 4]).Value.Replace("\n", "").TrimStart();
        //            if(Policy_Endorsement == "Policy Booking")
        //            {
        //                Policy_Endorsement = "Policy";
        //            }
        //            if (Client_N_E == "New")
        //            {
        //                Client_N_E = "New Client";
        //                New_Renewal = "New Policy";
        //            }
        //            else
        //            {
        //                Client_N_E = "Existing Client";
        //                New_Renewal = "Renewal Policy";
        //            }
        //            var Endo_Effective_Date = ((Microsoft.Office.Interop.Excel.Range)wks.Cells[i, 12]).Value;
        //            var Effective_Date = ((Microsoft.Office.Interop.Excel.Range)wks.Cells[i, 10]).Value;
        //            var END_Date = ((Microsoft.Office.Interop.Excel.Range)wks.Cells[i, 11]).Value;
        //            int ENdolen = Convert.ToString(Endo_Effective_Date).Length;
        //            if (ENdolen > 11)
        //            {
        //                Endo_Effective_Date = Endo_Effective_Date.ToString("dd/MM/yyyy");
        //            }
        //            int Efflen = Convert.ToString(Effective_Date).Length;
        //            if (Efflen > 11)
        //            {
        //                Effective_Date = Effective_Date.ToString("dd/MM/yyyy");
        //            }
        //            int ENDlen = Convert.ToString(END_Date).Length;
        //            if (ENDlen > 11)
        //            {
        //                END_Date = END_Date.ToString("dd/MM/yyyy");
        //            }

        //            var Premium_Amt = Convert.ToString(((Microsoft.Office.Interop.Excel.Range)wks.Cells[i, 16]).Value);
        //            var Revenue_Amt = Convert.ToString(((Microsoft.Office.Interop.Excel.Range)wks.Cells[i, 17]).Value);
        //            var Revenue_Pcnt = Convert.ToString(((Microsoft.Office.Interop.Excel.Range)wks.Cells[i, 14]).Text);
        //            Revenue_Pcnt = Revenue_Pcnt.Replace("\n", "").Replace("%", "").Replace(",", "").TrimStart();

        //            Policy_Type = Convert.ToString(((Microsoft.Office.Interop.Excel.Range)wks.Cells[i, 6]).Value);

        //            if (InsuredName.ToUpper().Contains("LIMITED") || InsuredName.ToUpper().Contains("LTD") || InsuredName.ToUpper().Contains("INTERNATIONAL"))
        //            {
        //                InsuredType = "Corporate";
        //            }
        //            else
        //            {
        //                InsuredType = "Retail";
        //            }
        //            if (Premium_Amt == "" || Premium_Amt == " " || Premium_Amt == null)
        //            {
        //                Premium_Amt = 0;
        //            }
        //            if (Revenue_Amt == "" || Revenue_Amt == " " || Revenue_Amt == null)
        //            {
        //                Revenue_Amt = 0;
        //            }
        //            if (Revenue_Pcnt == "" || Revenue_Pcnt == " " || Revenue_Pcnt == null)
        //            {
        //                Revenue_Pcnt = "0";
        //            }
        //            if (Terrorism == "" || Terrorism == " " || Terrorism == null)
        //            {
        //                Terrorism = "0";
        //            }
        //            sql.ExecuteSQLNonQuery("SP_EdelweissTransactions",
        //                       new SqlParameter { ParameterName = "@Imode", Value = 1 },
        //                       new SqlParameter { ParameterName = "@RDate", Value = RDate },
        //                       new SqlParameter { ParameterName = "@Rmonth", Value = Rmonth },
        //                       new SqlParameter { ParameterName = "@ClientName", Value = InsuredName },
        //                       new SqlParameter { ParameterName = "@Itype", Value = InsuredType },
        //                       new SqlParameter { ParameterName = "@Policy_No", Value = PolicyNo },
        //                       new SqlParameter { ParameterName = "@Revenue_Pcnt", Value = Revenue_Pcnt },
        //                       new SqlParameter { ParameterName = "@Client_N_E", Value = Client_N_E },
        //                       new SqlParameter { ParameterName = "@New_Renewal", Value = New_Renewal },
        //                       new SqlParameter { ParameterName = "@Endo_Effective_Date", Value = Endo_Effective_Date },
        //                       new SqlParameter { ParameterName = "@Effective_Date", Value = Effective_Date },
        //                       new SqlParameter { ParameterName = "@END_Date", Value = END_Date },
        //                       new SqlParameter { ParameterName = "@Policy_Type", Value = Policy_Type },
        //                       new SqlParameter { ParameterName = "@Premium_Amt", Value = Premium_Amt },
        //                       new SqlParameter { ParameterName = "@TranID", Value = TranID },
        //                       new SqlParameter { ParameterName = "@Revenue_Amt", Value = Revenue_Amt },
        //                       new SqlParameter { ParameterName = "@Terrorism", Value = Terrorism },
        //                       new SqlParameter { ParameterName = "@Insurance", Value = Insurance },
        //                       new SqlParameter { ParameterName = "@Salesby", Value = Salesby },
        //                       new SqlParameter { ParameterName = "@Serviceby", Value = Serviceby },
        //                       new SqlParameter { ParameterName = "@location", Value = location },
        //                       new SqlParameter { ParameterName = "@Support", Value = Support },
        //                       new SqlParameter { ParameterName = "@Policy_Endorsement", Value = Policy_Endorsement },
        //                       new SqlParameter { ParameterName = "@RFormat", Value = "F1" },
        //                       new SqlParameter { ParameterName = "@InvNo", Value = "EGIX" },
        //                       new SqlParameter { ParameterName = "@ReportId", Value = "EGIX" },
        //                       new SqlParameter { ParameterName = "@DocName", Value = "Edelweiss General Insurance Co. Ltd." }
        //                       );
        //        }
        //    }
        //}

        //public static void InsertTransaction(Microsoft.Office.Interop.Excel.Workbook WB, string TranID, string RDate, string Insurance, string Salesby, string Serviceby, string location, string Support, string Rmonth)
        //{
        //    //HDFC Ergo General Insurance Co. Ltd.
        //    SQLProcs sql = new SQLProcs();
        //    Microsoft.Office.Interop.Excel.Worksheet wks = (Microsoft.Office.Interop.Excel.Worksheet)WB.Worksheets[1];
        //    Microsoft.Office.Interop.Excel.Range lastCell = wks.Cells.SpecialCells(Microsoft.Office.Interop.Excel.XlCellType.xlCellTypeLastCell, Type.Missing);
        //    int lastrow = lastCell.Row; var Terrorism = ""; var Policy_Endorsement = ""; var InsuredType = ""; var Policy_Type = ""; var New_Renewal = "";var Premium_Amt = "";
        //    for (int i = 2; i <= lastrow; i++)
        //    {
        //        string InsuredName = ((Microsoft.Office.Interop.Excel.Range)wks.Cells[i, 10]).Value;
        //        if (InsuredName != null && InsuredName != "" && InsuredName != " ")
        //        {
        //            InsuredName = InsuredName.Replace("\n", "").TrimStart();
        //            string PolicyNo = Convert.ToString(((Microsoft.Office.Interop.Excel.Range)wks.Cells[i, 9]).Value).Replace("\n", "").Replace("'", "").TrimStart();
        //            var Client_N_E = ((Microsoft.Office.Interop.Excel.Range)wks.Cells[i, 5]).Value.Replace("\n", "").TrimStart();
        //            Policy_Endorsement = ((Microsoft.Office.Interop.Excel.Range)wks.Cells[i, 6]).Value.Replace("\n", "").TrimStart();
        //            if (Client_N_E == "Renewal")
        //            {
        //                Client_N_E = "Existing Client";
        //                New_Renewal = "Renewal Policy";
        //            }
        //            else
        //            {
        //                Client_N_E = "New Client";
        //                New_Renewal = "New Policy";
        //            }
        //            var Endo_Effective_Date = ((Microsoft.Office.Interop.Excel.Range)wks.Cells[i, 3]).Value;
        //            var Effective_Date = ((Microsoft.Office.Interop.Excel.Range)wks.Cells[i, 11]).Value;
        //            var END_Date = ((Microsoft.Office.Interop.Excel.Range)wks.Cells[i, 12]).Value;
        //            int ENdolen = Convert.ToString(Endo_Effective_Date).Length;
        //            if (ENdolen > 11)
        //            {
        //                Endo_Effective_Date = Endo_Effective_Date.ToString("dd/MM/yyyy");
        //            }
        //            int Efflen = Convert.ToString(Effective_Date).Length;
        //            if (Efflen > 11)
        //            {
        //                Effective_Date = Effective_Date.ToString("dd/MM/yyyy");
        //            }
        //            int ENDlen = Convert.ToString(END_Date).Length;
        //            if (ENDlen > 11)
        //            {
        //                END_Date = END_Date.ToString("dd/MM/yyyy");
        //            }

        //            var PA = Convert.ToString(((Microsoft.Office.Interop.Excel.Range)wks.Cells[i, 35]).Value);
        //            var PA1 = Convert.ToString(((Microsoft.Office.Interop.Excel.Range)wks.Cells[i, 32]).Value);
        //            var Revenue_Amt = Convert.ToString(((Microsoft.Office.Interop.Excel.Range)wks.Cells[i, 44]).Value);
        //            var Revenue_Pcnt = Convert.ToString(((Microsoft.Office.Interop.Excel.Range)wks.Cells[i, 28]).Text);
        //            Revenue_Pcnt = Revenue_Pcnt.Replace("\n", "").Replace("%", "").Replace(",", "").TrimStart();
        //            var TP1 = Convert.ToString(((Microsoft.Office.Interop.Excel.Range)wks.Cells[i, 34]).Value);
        //            var TP = Convert.ToString(((Microsoft.Office.Interop.Excel.Range)wks.Cells[i, 33]).Value);

        //            Policy_Type = Convert.ToString(((Microsoft.Office.Interop.Excel.Range)wks.Cells[i, 4]).Value);

        //            if (InsuredName.ToUpper().Contains("LIMITED") || InsuredName.ToUpper().Contains("LTD") || InsuredName.ToUpper().Contains("INTERNATIONAL"))
        //            {
        //                InsuredType = "Corporate";
        //            }
        //            else
        //            {
        //                InsuredType = "Retail";
        //            }

        //            if (PA == "" || PA == " " || PA == null)
        //            {
        //                PA = "0";
        //            }
        //            if (PA1 == "" || PA1 == " " || PA1 == null)
        //            {
        //                PA1 = "0";
        //            }
        //            if (TP1 == "" || TP1 == " " || TP1 == null)
        //            {
        //                TP1 = "0";
        //            }
        //            if (TP == "" || TP == " " || TP == null)
        //            {
        //                TP = "0";
        //            }
        //            if (Policy_Type.ToUpper().Contains("GOODS CARRYING COMPREHENSIVE POLICY") || Policy_Type.ToUpper().Contains("PRIVATE CAR COMPREHENSIVE POLICY") 
        //                || Policy_Type.ToUpper().Contains("PRIVATE CAR PACKAGE POLICY BUNDLED"))
        //            {
        //                Premium_Amt = PA1; Terrorism = TP;
        //            }
        //            else
        //            {
        //                Premium_Amt = PA; Terrorism = TP1;
        //            }
        //            if (Premium_Amt == "" || Premium_Amt == " " || Premium_Amt == null)
        //            {
        //                Premium_Amt = "0";
        //            }
        //            if (Revenue_Amt == "" || Revenue_Amt == " " || Revenue_Amt == null)
        //            {
        //                Revenue_Amt = 0;
        //            }
        //            if (Revenue_Pcnt == "" || Revenue_Pcnt == " " || Revenue_Pcnt == null)
        //            {
        //                Revenue_Pcnt = "0";
        //            }
        //            if (Terrorism == "" || Terrorism == " " || Terrorism == null)
        //            {
        //                Terrorism = "0";
        //            }
        //            sql.ExecuteSQLNonQuery("SP_HDFCTransactions",
        //                       new SqlParameter { ParameterName = "@Imode", Value = 1 },
        //                       new SqlParameter { ParameterName = "@RDate", Value = RDate },
        //                       new SqlParameter { ParameterName = "@Rmonth", Value = Rmonth },
        //                       new SqlParameter { ParameterName = "@ClientName", Value = InsuredName },
        //                       new SqlParameter { ParameterName = "@Itype", Value = InsuredType },
        //                       new SqlParameter { ParameterName = "@Policy_No", Value = PolicyNo },
        //                       new SqlParameter { ParameterName = "@Revenue_Pcnt", Value = Revenue_Pcnt },
        //                       new SqlParameter { ParameterName = "@Client_N_E", Value = Client_N_E },
        //                       new SqlParameter { ParameterName = "@New_Renewal", Value = New_Renewal },
        //                       new SqlParameter { ParameterName = "@Endo_Effective_Date", Value = Endo_Effective_Date },
        //                       new SqlParameter { ParameterName = "@Effective_Date", Value = Effective_Date },
        //                       new SqlParameter { ParameterName = "@END_Date", Value = END_Date },
        //                       new SqlParameter { ParameterName = "@Policy_Type", Value = Policy_Type },
        //                       new SqlParameter { ParameterName = "@Premium_Amt", Value = Premium_Amt },
        //                       new SqlParameter { ParameterName = "@TranID", Value = TranID },
        //                       new SqlParameter { ParameterName = "@Revenue_Amt", Value = Revenue_Amt },
        //                       new SqlParameter { ParameterName = "@Terrorism", Value = Terrorism },
        //                       new SqlParameter { ParameterName = "@Insurance", Value = Insurance },
        //                       new SqlParameter { ParameterName = "@Salesby", Value = Salesby },
        //                       new SqlParameter { ParameterName = "@Serviceby", Value = Serviceby },
        //                       new SqlParameter { ParameterName = "@location", Value = location },
        //                       new SqlParameter { ParameterName = "@Support", Value = Support },
        //                       new SqlParameter { ParameterName = "@Policy_Endorsement", Value = Policy_Endorsement },
        //                       new SqlParameter { ParameterName = "@RFormat", Value = "F1" },
        //                       new SqlParameter { ParameterName = "@InvNo", Value = "HEGX" },
        //                       new SqlParameter { ParameterName = "@ReportId", Value = "HEGX" },
        //                       new SqlParameter { ParameterName = "@DocName", Value = "HDFC Ergo General Insurance Co. Ltd." }
        //                       );
        //        }
        //    }
        //}

        //public static void InsertTransaction(Microsoft.Office.Interop.Excel.Workbook WB, string TranID, string RDate, string Insurance, string Salesby, string Serviceby, string location, string Support, string Rmonth)
        //{
        //    //Iffco Tokio General Insurance Co.Ltd.
        //    SQLProcs sql = new SQLProcs();
        //    Microsoft.Office.Interop.Excel.Worksheet wks = (Microsoft.Office.Interop.Excel.Worksheet)WB.Worksheets[1];
        //    Microsoft.Office.Interop.Excel.Range lastCell = wks.Cells.SpecialCells(Microsoft.Office.Interop.Excel.XlCellType.xlCellTypeLastCell, Type.Missing);
        //    int lastrow = lastCell.Row; var Terrorism = ""; var Policy_Endorsement = ""; var InsuredType = ""; var Policy_Type = ""; var New_Renewal = ""; var Premium_Amt = "";
        //    for (int i = 8; i <= lastrow; i++)
        //    {
        //        string InsuredName = ((Microsoft.Office.Interop.Excel.Range)wks.Cells[i, 19]).Value;
        //        Policy_Endorsement = ((Microsoft.Office.Interop.Excel.Range)wks.Cells[i, 11]).Value;
        //        if ((InsuredName != null && InsuredName != "" && InsuredName != " ") && Policy_Endorsement != "PYMT" && Policy_Endorsement != null)
        //        {
        //            Policy_Endorsement = Policy_Endorsement.Replace("\n", "").TrimStart();
        //            InsuredName = InsuredName.Replace("\n", "").Replace(".", "").TrimStart();
        //            string PolicyNo = Convert.ToString(((Microsoft.Office.Interop.Excel.Range)wks.Cells[i, 13]).Value).Replace("\n", "").TrimStart();
        //            var ENDNO = Convert.ToString(((Microsoft.Office.Interop.Excel.Range)wks.Cells[i, 9]).Value);
        //            if(Convert.ToDecimal(ENDNO) > 1)
        //            {
        //                PolicyNo = PolicyNo + "/" + Convert.ToString(Convert.ToDecimal(ENDNO) - 1);
        //            }
        //            else
        //            {
        //                PolicyNo = PolicyNo + "/0";
        //            }
        //            var Endo_Effective_Date = ((Microsoft.Office.Interop.Excel.Range)wks.Cells[i, 10]).Value;
        //            var Effective_Date = ((Microsoft.Office.Interop.Excel.Range)wks.Cells[i, 21]).Value;
        //            var END_Date = ((Microsoft.Office.Interop.Excel.Range)wks.Cells[i, 22]).Value;
        //            int ENdolen = Convert.ToString(Endo_Effective_Date).Length;
        //            if (ENdolen > 11)
        //            {
        //                Endo_Effective_Date = Endo_Effective_Date.ToString("dd/MM/yyyy");
        //            }
        //            int Efflen = Convert.ToString(Effective_Date).Length;
        //            if (Efflen > 11)
        //            {
        //                Effective_Date = Effective_Date.ToString("dd/MM/yyyy");
        //            }
        //            int ENDlen = Convert.ToString(END_Date).Length;
        //            if (ENDlen > 11)
        //            {
        //                END_Date = END_Date.ToString("dd/MM/yyyy");
        //            }

        //            var PA = Convert.ToString(((Microsoft.Office.Interop.Excel.Range)wks.Cells[i, 23]).Value);
        //            var PA1 = Convert.ToString(((Microsoft.Office.Interop.Excel.Range)wks.Cells[i, 25]).Value);
        //            var Revenue_Amt = Convert.ToString(((Microsoft.Office.Interop.Excel.Range)wks.Cells[i, 33]).Value);
        //            //var Revenue_Pcnt = Convert.ToString(((Microsoft.Office.Interop.Excel.Range)wks.Cells[i, 28]).Text);
        //            //Revenue_Pcnt = Revenue_Pcnt.Replace("\n", "").Replace("%", "").Replace(",", "").TrimStart();
        //            Terrorism = Convert.ToString(((Microsoft.Office.Interop.Excel.Range)wks.Cells[i, 26]).Value);

        //            Policy_Type = Convert.ToString(((Microsoft.Office.Interop.Excel.Range)wks.Cells[i, 16]).Value);

        //            if (InsuredName.ToUpper().Contains("LIMITED") || InsuredName.ToUpper().Contains("LTD") || InsuredName.ToUpper().Contains("SERVICES") 
        //                || InsuredName.ToUpper().Contains("SOLUTIONS") || InsuredName.ToUpper().Contains("TECHNOLOGIES"))
        //            {
        //                InsuredType = "Corporate";
        //            }
        //            else
        //            {
        //                InsuredType = "Retail";
        //            }
        //            if (Policy_Endorsement == "ENDT")
        //            {
        //                Policy_Endorsement = "Endorsement";
        //            }
        //            else
        //            {
        //                Policy_Endorsement = "Policy";
        //            }
        //            if (PA == "" || PA == " " || PA == null)
        //            {
        //                PA = "0";
        //            }
        //            if (PA1 == "" || PA1 == " " || PA1 == null)
        //            {
        //                PA1 = "0";
        //            }
        //            if (Policy_Type.ToLower().Contains("commercial vehicle") || Policy_Type.ToLower().Contains("commercial vehicle tp only") || Policy_Type.ToLower().Contains("motor"))
        //            {
        //                Premium_Amt = PA1; 
        //            }
        //            else
        //            {
        //                Premium_Amt = PA; 
        //            }
        //            if (Premium_Amt == "" || Premium_Amt == " " || Premium_Amt == null)
        //            {
        //                Premium_Amt = "0";
        //            }
        //            if (Revenue_Amt == "" || Revenue_Amt == " " || Revenue_Amt == null)
        //            {
        //                Revenue_Amt = 0;
        //            }
        //            if (Terrorism == "" || Terrorism == " " || Terrorism == null)
        //            {
        //                Terrorism = "0";
        //            }
        //            sql.ExecuteSQLNonQuery("SP_IffcoTransactions",
        //                       new SqlParameter { ParameterName = "@Imode", Value = 1 },
        //                       new SqlParameter { ParameterName = "@RDate", Value = RDate },
        //                       new SqlParameter { ParameterName = "@Rmonth", Value = Rmonth },
        //                       new SqlParameter { ParameterName = "@ClientName", Value = InsuredName },
        //                       new SqlParameter { ParameterName = "@Itype", Value = InsuredType },
        //                       new SqlParameter { ParameterName = "@Policy_No", Value = PolicyNo },
        //                       //new SqlParameter { ParameterName = "@Revenue_Pcnt", Value = Revenue_Pcnt },
        //                       new SqlParameter { ParameterName = "@Endo_Effective_Date", Value = Endo_Effective_Date },
        //                       new SqlParameter { ParameterName = "@Effective_Date", Value = Effective_Date },
        //                       new SqlParameter { ParameterName = "@END_Date", Value = END_Date },
        //                       new SqlParameter { ParameterName = "@Policy_Type", Value = Policy_Type },
        //                       new SqlParameter { ParameterName = "@Premium_Amt", Value = Premium_Amt },
        //                       new SqlParameter { ParameterName = "@TranID", Value = TranID },
        //                       new SqlParameter { ParameterName = "@Revenue_Amt", Value = Revenue_Amt },
        //                       new SqlParameter { ParameterName = "@Terrorism", Value = Terrorism },
        //                       new SqlParameter { ParameterName = "@Insurance", Value = Insurance },
        //                       new SqlParameter { ParameterName = "@Salesby", Value = Salesby },
        //                       new SqlParameter { ParameterName = "@Serviceby", Value = Serviceby },
        //                       new SqlParameter { ParameterName = "@location", Value = location },
        //                       new SqlParameter { ParameterName = "@Support", Value = Support },
        //                       new SqlParameter { ParameterName = "@Policy_Endorsement", Value = Policy_Endorsement },
        //                       new SqlParameter { ParameterName = "@RFormat", Value = "F1" },
        //                       new SqlParameter { ParameterName = "@InvNo", Value = "ITGX" },
        //                       new SqlParameter { ParameterName = "@ReportId", Value = "ITGX" },
        //                       new SqlParameter { ParameterName = "@DocName", Value = "Iffco Tokio General Insurance Co.Ltd." }
        //                       );
        //        }
        //    }
        //}

        //public static void InsertTransaction(Microsoft.Office.Interop.Excel.Workbook WB, string TranID, string RDate, string Insurance, string Salesby, string Serviceby, string location, string Support, string Rmonth)
        //{
        //    //Liberty Videocon General Insurance Co. Ltd.
        //    SQLProcs sql = new SQLProcs();
        //    Microsoft.Office.Interop.Excel.Worksheet wks = (Microsoft.Office.Interop.Excel.Worksheet)WB.Worksheets[1];
        //    Microsoft.Office.Interop.Excel.Range lastCell = wks.Cells.SpecialCells(Microsoft.Office.Interop.Excel.XlCellType.xlCellTypeLastCell, Type.Missing);
        //    int lastrow = lastCell.Row; var Terrorism = ""; var Policy_Endorsement = ""; var InsuredType = ""; var Policy_Type = ""; var New_Renewal = ""; var Premium_Amt = "";
        //    for (int i = 2; i <= lastrow; i++)
        //    {
        //        string InsuredName = ((Microsoft.Office.Interop.Excel.Range)wks.Cells[i, 29]).Value;
        //        if (InsuredName != null && InsuredName != "" && InsuredName != " ")
        //        {
        //            InsuredName = InsuredName.Replace("\n", "").TrimStart();
        //            string PolicyNo = Convert.ToString(((Microsoft.Office.Interop.Excel.Range)wks.Cells[i, 7]).Value);
        //            if(PolicyNo == null || PolicyNo == "" || PolicyNo == " ")
        //            {
        //                PolicyNo = Convert.ToString(((Microsoft.Office.Interop.Excel.Range)wks.Cells[i, 8]).Value).TrimStart();
        //            }
        //            else
        //            {
        //                PolicyNo = PolicyNo.TrimStart();
        //            }
        //            var Client_N_E = ((Microsoft.Office.Interop.Excel.Range)wks.Cells[i, 31]).Value.Replace("\n", "").TrimStart();
        //            Policy_Endorsement = ((Microsoft.Office.Interop.Excel.Range)wks.Cells[i, 32]).Value.Replace("\n", "").TrimStart();
        //            if (Client_N_E == "New Business")
        //            {
        //                Client_N_E = "New Client";
        //                New_Renewal = "New Policy";
        //            }
        //            else
        //            {
        //                Client_N_E = "Existing Client";
        //                New_Renewal = "Renewal Policy";
        //            }
        //            var Endo_Effective_Date = ((Microsoft.Office.Interop.Excel.Range)wks.Cells[i, 27]).Value;
        //            var Effective_Date = ((Microsoft.Office.Interop.Excel.Range)wks.Cells[i, 25]).Value;
        //            var END_Date = ((Microsoft.Office.Interop.Excel.Range)wks.Cells[i, 26]).Value;

        //            //END_Date = END_Date.Substring(0, END_Date.LastIndexOf(" "));
        //            if (Endo_Effective_Date != null)
        //            {
        //                int ENdolen = Convert.ToString(Endo_Effective_Date).Length;
        //                if (ENdolen > 11)
        //                {
        //                    Endo_Effective_Date = Endo_Effective_Date.ToString("dd/MM/yyyy");
        //                }
        //            }
        //            int Efflen = Convert.ToString(Effective_Date).Length;
        //            if (Efflen == 16)
        //            {
        //                Effective_Date = Effective_Date.Substring(0, Effective_Date.LastIndexOf(" "));
        //            }
        //            else if (Efflen > 16)
        //            {
        //                Effective_Date = Effective_Date.ToString("dd/MM/yyyy");
        //            }
        //            int ENDlen = Convert.ToString(END_Date).Length;
        //            if (ENDlen == 16)
        //            {
        //                END_Date = END_Date.Substring(0, END_Date.LastIndexOf(" "));
        //            }
        //            else if (ENDlen > 16)
        //            {
        //                END_Date = END_Date.ToString("dd/MM/yyyy");
        //            }

        //            Premium_Amt = Convert.ToString(((Microsoft.Office.Interop.Excel.Range)wks.Cells[i, 18]).Value).Replace("(", "").Replace(")", "").Replace(",", "").TrimStart();
        //            var Revenue_Amt = Convert.ToString(((Microsoft.Office.Interop.Excel.Range)wks.Cells[i, 23]).Value).Replace("(", "").Replace(")", "").Replace(",", "").TrimStart();
        //            var Revenue_Pcnt = Convert.ToString(((Microsoft.Office.Interop.Excel.Range)wks.Cells[i, 19]).Text);
        //            Revenue_Pcnt = Revenue_Pcnt.Replace("\n", "").Replace("%", "").Replace(",", "").TrimStart();
        //            Terrorism = Convert.ToString(((Microsoft.Office.Interop.Excel.Range)wks.Cells[i, 17]).Value).Replace("-", "").Replace("(", "").Replace(")", "").Replace(",", "").TrimStart();

        //            Policy_Type = Convert.ToString(((Microsoft.Office.Interop.Excel.Range)wks.Cells[i, 14]).Value);

        //            if (InsuredName.ToUpper().Contains("LIMITED") || InsuredName.ToUpper().Contains("LTD") || InsuredName.ToUpper().Contains("INTERNATIONAL"))
        //            {
        //                InsuredType = "Corporate";
        //            }
        //            else
        //            {
        //                InsuredType = "Retail";
        //            }
        //            if (Premium_Amt == "" || Premium_Amt == " " || Premium_Amt == null)
        //            {
        //                Premium_Amt = "0";
        //            }
        //            if (Revenue_Amt == "" || Revenue_Amt == " " || Revenue_Amt == null)
        //            {
        //                Revenue_Amt = 0;
        //            }
        //            if (Revenue_Pcnt == "" || Revenue_Pcnt == " " || Revenue_Pcnt == null)
        //            {
        //                Revenue_Pcnt = "0";
        //            }
        //            if (Terrorism == "" || Terrorism == " " || Terrorism == null)
        //            {
        //                Terrorism = "0";
        //            }
        //            sql.ExecuteSQLNonQuery("SP_LibertyTransactions",
        //                       new SqlParameter { ParameterName = "@Imode", Value = 1 },
        //                       new SqlParameter { ParameterName = "@RDate", Value = RDate },
        //                       new SqlParameter { ParameterName = "@Rmonth", Value = Rmonth },
        //                       new SqlParameter { ParameterName = "@ClientName", Value = InsuredName },
        //                       new SqlParameter { ParameterName = "@Itype", Value = InsuredType },
        //                       new SqlParameter { ParameterName = "@Policy_No", Value = PolicyNo },
        //                       new SqlParameter { ParameterName = "@Revenue_Pcnt", Value = Revenue_Pcnt },
        //                       new SqlParameter { ParameterName = "@Client_N_E", Value = Client_N_E },
        //                       new SqlParameter { ParameterName = "@New_Renewal", Value = New_Renewal },
        //                       new SqlParameter { ParameterName = "@Endo_Effective_Date", Value = Endo_Effective_Date },
        //                       new SqlParameter { ParameterName = "@Effective_Date", Value = Effective_Date },
        //                       new SqlParameter { ParameterName = "@END_Date", Value = END_Date },
        //                       new SqlParameter { ParameterName = "@Policy_Type", Value = Policy_Type },
        //                       new SqlParameter { ParameterName = "@Premium_Amt", Value = Premium_Amt },
        //                       new SqlParameter { ParameterName = "@TranID", Value = TranID },
        //                       new SqlParameter { ParameterName = "@Revenue_Amt", Value = Revenue_Amt },
        //                       new SqlParameter { ParameterName = "@Terrorism", Value = Terrorism },
        //                       new SqlParameter { ParameterName = "@Insurance", Value = Insurance },
        //                       new SqlParameter { ParameterName = "@Salesby", Value = Salesby },
        //                       new SqlParameter { ParameterName = "@Serviceby", Value = Serviceby },
        //                       new SqlParameter { ParameterName = "@location", Value = location },
        //                       new SqlParameter { ParameterName = "@Support", Value = Support },
        //                       new SqlParameter { ParameterName = "@Policy_Endorsement", Value = Policy_Endorsement },
        //                       new SqlParameter { ParameterName = "@RFormat", Value = "F1" },
        //                       new SqlParameter { ParameterName = "@InvNo", Value = "LVGI" },
        //                       new SqlParameter { ParameterName = "@ReportId", Value = "LVGI" },
        //                       new SqlParameter { ParameterName = "@DocName", Value = "Liberty Videocon General Insurance Co.Ltd." }
        //                       );
        //        }
        //    }
        //}

        //public static void InsertTransaction(Microsoft.Office.Interop.Excel.Workbook WB, string TranID, string RDate, string Insurance, string Salesby, string Serviceby, string location, string Support, string Rmonth)
        //{
        //    //Raheja QBE General Insurance Co.Ltd.
        //    SQLProcs sql = new SQLProcs();
        //    Microsoft.Office.Interop.Excel.Worksheet wks = (Microsoft.Office.Interop.Excel.Worksheet)WB.Worksheets[1];
        //    Microsoft.Office.Interop.Excel.Range lastCell = wks.Cells.SpecialCells(Microsoft.Office.Interop.Excel.XlCellType.xlCellTypeLastCell, Type.Missing);
        //    int lastrow = lastCell.Row; var Terrorism = ""; var Policy_Endorsement = ""; var InsuredType = ""; var Policy_Type = ""; var New_Renewal = ""; var Premium_Amt = "";
        //    for (int i = 2; i <= lastrow; i++)
        //    {
        //        string InsuredName = "";
        //        string In1 = ((Microsoft.Office.Interop.Excel.Range)wks.Cells[i, 4]).Value;
        //        string In2 = ((Microsoft.Office.Interop.Excel.Range)wks.Cells[i, 5]).Value;
        //        if ((In1 != null && In1 != "" && In1 != " ") || (In2 != null && In2 != "" && In2 != " "))
        //        {
        //            if (In2 != null && In1 != null)
        //            {
        //                InsuredName = In1.Replace("\n", "").TrimStart() + " " + In2.TrimStart();
        //            }
        //            else if (In2 != null && In1 == null)
        //            {
        //                InsuredName = In2.Replace("\n", "").TrimStart();
        //            }
        //            else
        //            {
        //                InsuredName = In1.Replace("\n", "").TrimStart();
        //            }
        //            var BRcode = Convert.ToString(((Microsoft.Office.Interop.Excel.Range)wks.Cells[i, 1]).Value);
        //            string EndosNo = Convert.ToString(((Microsoft.Office.Interop.Excel.Range)wks.Cells[i, 7]).Value);
        //            bool result = EndosNo.Any(x => !char.IsLetter(x));
        //            string PolicyNo = Convert.ToString(((Microsoft.Office.Interop.Excel.Range)wks.Cells[i, 6]).Value);
        //            if (PolicyNo != null && PolicyNo != "" && PolicyNo != " ")
        //            {
        //                if(result == true)
        //                {
        //                    if(Convert.ToDecimal(EndosNo) > 1)
        //                    {
        //                        if (EndosNo.Length == 1)
        //                        {
        //                            EndosNo = "0000" + Convert.ToString(Convert.ToDecimal(EndosNo) - 1);
        //                        }
        //                        else if (EndosNo.Length == 2)
        //                        {
        //                            EndosNo = "000" + Convert.ToString(Convert.ToDecimal(EndosNo) - 1);
        //                        }
        //                        else if (EndosNo.Length == 3)
        //                        {
        //                            EndosNo = "00" + Convert.ToString(Convert.ToDecimal(EndosNo) - 1);
        //                        }
        //                        else if (EndosNo.Length == 3)
        //                        {
        //                            EndosNo = "0" + Convert.ToString(Convert.ToDecimal(EndosNo) - 1);
        //                        }
        //                        PolicyNo = "0" + BRcode + PolicyNo + EndosNo;
        //                    }
        //                    else
        //                    {
        //                        PolicyNo = "0" + BRcode + PolicyNo + "00000";
        //                    }
        //                }
        //                else
        //                {
        //                    PolicyNo = "0" + BRcode + PolicyNo + EndosNo;
        //                }                 
        //            }

        //            var Client_N_E = ((Microsoft.Office.Interop.Excel.Range)wks.Cells[i, 3]).Value.Replace("\n", "").TrimStart();
        //            //Policy_Endorsement = ((Microsoft.Office.Interop.Excel.Range)wks.Cells[i, 32]).Value.Replace("\n", "").TrimStart();
        //            if (Client_N_E == "New Business" || Client_N_E == "New")
        //            {
        //                Client_N_E = "New Client";
        //                New_Renewal = "New Policy";
        //            }
        //            else
        //            {
        //                Client_N_E = "Existing Client";
        //                New_Renewal = "Renewal Policy";
        //            }
        //            var Endo_Effective_Date = ((Microsoft.Office.Interop.Excel.Range)wks.Cells[i, 27]).Value;
        //            var Effective_Date = ((Microsoft.Office.Interop.Excel.Range)wks.Cells[i, 13]).Value;
        //            var END_Date = ((Microsoft.Office.Interop.Excel.Range)wks.Cells[i, 14]).Value;

        //            //END_Date = END_Date.Substring(0, END_Date.LastIndexOf(" "));
        //            if (Endo_Effective_Date != null)
        //            {
        //                int ENdolen = Convert.ToString(Endo_Effective_Date).Length;
        //                if (ENdolen > 11)
        //                {
        //                    Endo_Effective_Date = Endo_Effective_Date.ToString("dd/MM/yyyy");
        //                }
        //                else if(ENdolen <= 8)
        //                {
        //                    //DateTime date = Convert.ToDateTime(Endo_Effective_Date, CultureInfo.InvariantCulture);
        //                    //var d1 = date.ToString("yyyy-MM-dd");
        //                    var year = Convert.ToString(Endo_Effective_Date).Substring(0, 4);
        //                    string month = Convert.ToString(Endo_Effective_Date).Substring(4);
        //                    month = month.Substring(0, 2);
        //                    string date = Convert.ToString(Endo_Effective_Date).Substring(6);
        //                    Endo_Effective_Date = Convert.ToString(date + "/" + month + "/" + year);
        //                }
        //            }
        //            int Efflen = Convert.ToString(Effective_Date).Length;
        //            if (Efflen == 16)
        //            {
        //                Effective_Date = Effective_Date.Substring(0, Effective_Date.LastIndexOf(" "));
        //            }
        //            else if (Efflen > 16)
        //            {
        //                Effective_Date = Effective_Date.ToString("dd/MM/yyyy");
        //            }
        //            int ENDlen = Convert.ToString(END_Date).Length;
        //            if (ENDlen == 16)
        //            {
        //                END_Date = END_Date.Substring(0, END_Date.LastIndexOf(" "));
        //            }
        //            else if (ENDlen > 16)
        //            {
        //                END_Date = END_Date.ToString("dd/MM/yyyy");
        //            }

        //            var PA = Convert.ToString(((Microsoft.Office.Interop.Excel.Range)wks.Cells[i, 17]).Value);
        //            var PA1 = Convert.ToString(((Microsoft.Office.Interop.Excel.Range)wks.Cells[i, 28]).Value);
        //            var Revenue_Amt = Convert.ToString(((Microsoft.Office.Interop.Excel.Range)wks.Cells[i, 18]).Value);
        //            //var Revenue_Pcnt = Convert.ToString(((Microsoft.Office.Interop.Excel.Range)wks.Cells[i, 19]).Text);
        //            //Revenue_Pcnt = Revenue_Pcnt.Replace("\n", "").Replace("%", "").Replace(",", "").TrimStart();
        //            Terrorism = Convert.ToString(((Microsoft.Office.Interop.Excel.Range)wks.Cells[i, 29]).Value);
        //            var Reward_Pcnt = Convert.ToString(((Microsoft.Office.Interop.Excel.Range)wks.Cells[i, 50]).Value);
        //            var Reward_Amt = Convert.ToString(((Microsoft.Office.Interop.Excel.Range)wks.Cells[i, 51]).Value);
        //            Policy_Type = Convert.ToString(((Microsoft.Office.Interop.Excel.Range)wks.Cells[i, 10]).Value);

        //            if (InsuredName.ToUpper().Contains("LIMITED") || InsuredName.ToUpper().Contains("LTD") || InsuredName.ToUpper().Contains("INSURANCE"))
        //            {
        //                InsuredType = "Corporate";
        //            }
        //            else
        //            {
        //                InsuredType = "Retail";
        //            }
        //            if (PA == "" || PA == " " || PA == null)
        //            {
        //                PA = "0";
        //            }
        //            else
        //            {
        //                PA = PA.Replace(",", "").TrimStart();
        //            }
        //            if (PA1 == "" || PA1 == " " || PA1 == null)
        //            {
        //                PA1 = "0";
        //            }
        //            else
        //            {
        //                PA1 = PA1.Replace(",", "").TrimStart();
        //            }
        //            if (Policy_Type.ToUpper().Contains("MVA"))
        //            {
        //                Premium_Amt = PA1;
        //            }
        //            else
        //            {
        //                Premium_Amt = PA; 
        //            }
        //            if (Premium_Amt == "" || Premium_Amt == " " || Premium_Amt == null)
        //            {
        //                Premium_Amt = "0";
        //            }
        //            if (Revenue_Amt == "" || Revenue_Amt == " " || Revenue_Amt == null)
        //            {
        //                Revenue_Amt = 0;
        //            }
        //            if (Reward_Pcnt == "" || Reward_Pcnt == " " || Reward_Pcnt == null)
        //            {
        //                Reward_Pcnt = "0";
        //            }
        //            if (Reward_Amt == "" || Reward_Amt == " " || Reward_Amt == null)
        //            {
        //                Reward_Amt = "0";
        //            }
        //            else
        //            {
        //                Reward_Amt = Reward_Amt.Replace(",", "").TrimStart();
        //            }
        //            if (Terrorism == "" || Terrorism == " " || Terrorism == null)
        //            {
        //                Terrorism = "0";
        //            }
        //            sql.ExecuteSQLNonQuery("SP_RahejaTransactions",
        //                       new SqlParameter { ParameterName = "@Imode", Value = 1 },
        //                       new SqlParameter { ParameterName = "@RDate", Value = RDate },
        //                       new SqlParameter { ParameterName = "@Rmonth", Value = Rmonth },
        //                       new SqlParameter { ParameterName = "@ClientName", Value = InsuredName },
        //                       new SqlParameter { ParameterName = "@Itype", Value = InsuredType },
        //                       new SqlParameter { ParameterName = "@Policy_No", Value = PolicyNo },
        //                       new SqlParameter { ParameterName = "@Reward_Pcnt", Value = Reward_Pcnt },
        //                       new SqlParameter { ParameterName = "@Reward_Amt", Value = Reward_Amt },
        //                       new SqlParameter { ParameterName = "@Client_N_E", Value = Client_N_E },
        //                       new SqlParameter { ParameterName = "@New_Renewal", Value = New_Renewal },
        //                       new SqlParameter { ParameterName = "@Endo_Effective_Date", Value = Endo_Effective_Date },
        //                       new SqlParameter { ParameterName = "@Effective_Date", Value = Effective_Date },
        //                       new SqlParameter { ParameterName = "@END_Date", Value = END_Date },
        //                       new SqlParameter { ParameterName = "@Policy_Type", Value = Policy_Type },
        //                       new SqlParameter { ParameterName = "@Premium_Amt", Value = Premium_Amt },
        //                       new SqlParameter { ParameterName = "@TranID", Value = TranID },
        //                       new SqlParameter { ParameterName = "@Revenue_Amt", Value = Revenue_Amt },
        //                       new SqlParameter { ParameterName = "@Terrorism", Value = Terrorism },
        //                       new SqlParameter { ParameterName = "@Insurance", Value = Insurance },
        //                       new SqlParameter { ParameterName = "@Salesby", Value = Salesby },
        //                       new SqlParameter { ParameterName = "@Serviceby", Value = Serviceby },
        //                       new SqlParameter { ParameterName = "@location", Value = location },
        //                       new SqlParameter { ParameterName = "@Support", Value = Support },
        //                       new SqlParameter { ParameterName = "@Policy_Endorsement", Value = Policy_Endorsement },
        //                       new SqlParameter { ParameterName = "@RFormat", Value = "F1" },
        //                       new SqlParameter { ParameterName = "@InvNo", Value = "RQGX" },
        //                       new SqlParameter { ParameterName = "@ReportId", Value = "RQGX" },
        //                       new SqlParameter { ParameterName = "@DocName", Value = "Raheja QBE General Insurance Co.Ltd." }
        //                       );
        //        }
        //    }
        //}
        public static void InsertTransaction(Microsoft.Office.Interop.Excel.Workbook WB, string TranID, string RDate, string Insurance, string Salesby, string Serviceby, string location, string Support, string Rmonth)
        {
            //Reliance General Insurance Co. Ltd.
            SQLProcs sql = new SQLProcs();
            Microsoft.Office.Interop.Excel.Worksheet wks = (Microsoft.Office.Interop.Excel.Worksheet)WB.Worksheets[1];
            Microsoft.Office.Interop.Excel.Range lastCell = wks.Cells.SpecialCells(Microsoft.Office.Interop.Excel.XlCellType.xlCellTypeLastCell, Type.Missing);
            int lastrow = lastCell.Row; var Terrorism = ""; var Policy_Endorsement = ""; var InsuredType = ""; var Policy_Type = ""; var New_Renewal = ""; var Revenue_Amt = "";
            int K = 0;
            for (int i = 2; i <= lastrow; i++)
            {
                string InsuredName = ((Microsoft.Office.Interop.Excel.Range)wks.Cells[i, 12]).Value;
                if(InsuredName == null)
                {
                    K++;
                    if (K > 10)
                    {
                        i = lastrow + 1;
                    }
                }
                if (InsuredName != null && InsuredName != "" && InsuredName != " ")
                {
                    InsuredName = InsuredName.Replace("\n", "").TrimStart();
                    string PolicyNo = Convert.ToString(((Microsoft.Office.Interop.Excel.Range)wks.Cells[i, 10]).Value).Replace("\n", "").TrimStart();
                    Policy_Endorsement = Convert.ToString(((Microsoft.Office.Interop.Excel.Range)wks.Cells[i, 11]).Value).Replace("\n", "").TrimStart();
                    PolicyNo = PolicyNo + "/" + Policy_Endorsement;
                    if (Policy_Endorsement == "0")
                    {
                        Policy_Endorsement = "Policy";
                    }
                    else
                    {
                        Policy_Endorsement = "Endorsement";
                    }

                    var Premium_Amt = Convert.ToString(((Microsoft.Office.Interop.Excel.Range)wks.Cells[i, 14]).Value).Replace(",", "").TrimStart();
                    Terrorism = Convert.ToString(((Microsoft.Office.Interop.Excel.Range)wks.Cells[i, 19]).Value).Replace(",", "").TrimStart();
                    var RA1 = Convert.ToString(((Microsoft.Office.Interop.Excel.Range)wks.Cells[i, 16]).Value).Replace(",", "").TrimStart();
                    var RA2 = Convert.ToString(((Microsoft.Office.Interop.Excel.Range)wks.Cells[i, 21]).Value).Replace(",", "").TrimStart();
                    var Revenue_Pcnt = Convert.ToString(((Microsoft.Office.Interop.Excel.Range)wks.Cells[i, 15]).Text);
                    Revenue_Pcnt = Revenue_Pcnt.Replace("\n", "").Replace("%", "").Replace(",", "").TrimStart();
                    var Reward_Pcnt = Convert.ToString(((Microsoft.Office.Interop.Excel.Range)wks.Cells[i, 17]).Value);
                    var Reward_Amt = Convert.ToString(((Microsoft.Office.Interop.Excel.Range)wks.Cells[i, 18]).Value);

                    Policy_Type = Convert.ToString(((Microsoft.Office.Interop.Excel.Range)wks.Cells[i, 9]).Value);
                    Policy_Type = Convert.ToString(Regex.Match(Policy_Type, @"\d+").Value);

                    if (InsuredName.ToUpper().Contains("LIMITED") || InsuredName.ToUpper().Contains("LTD") || InsuredName.ToUpper().Contains("INTERNATIONAL"))
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
                    if (RA1 == "" || RA1 == " " || RA1 == null)
                    {
                        RA1 = 0;
                    }
                    if (RA2 == "" || RA2 == " " || RA2 == null)
                    {
                        RA2 = 0;
                    }
                    Revenue_Amt = Convert.ToString(Convert.ToDouble(RA1) + Convert.ToDouble(RA2));
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
                    sql.ExecuteSQLNonQuery("SP_RelianceTransactions",
                               new SqlParameter { ParameterName = "@Imode", Value = 1 },
                               new SqlParameter { ParameterName = "@RDate", Value = RDate },
                               new SqlParameter { ParameterName = "@Rmonth", Value = Rmonth },
                               new SqlParameter { ParameterName = "@ClientName", Value = InsuredName },
                               new SqlParameter { ParameterName = "@Itype", Value = InsuredType },
                               new SqlParameter { ParameterName = "@Policy_No", Value = PolicyNo },
                               new SqlParameter { ParameterName = "@Revenue_Pcnt", Value = Revenue_Pcnt },
                               new SqlParameter { ParameterName = "@Policy_Type", Value = Policy_Type },
                               new SqlParameter { ParameterName = "@Premium_Amt", Value = Premium_Amt },
                               new SqlParameter { ParameterName = "@TranID", Value = TranID },
                               new SqlParameter { ParameterName = "@Reward_Pcnt", Value = Reward_Pcnt },
                               new SqlParameter { ParameterName = "@Reward_Amt", Value = Reward_Amt },
                               new SqlParameter { ParameterName = "@Revenue_Amt", Value = Revenue_Amt },
                               new SqlParameter { ParameterName = "@Terrorism", Value = Terrorism },
                               new SqlParameter { ParameterName = "@Insurance", Value = Insurance },
                               new SqlParameter { ParameterName = "@Salesby", Value = Salesby },
                               new SqlParameter { ParameterName = "@Serviceby", Value = Serviceby },
                               new SqlParameter { ParameterName = "@location", Value = location },
                               new SqlParameter { ParameterName = "@Support", Value = Support },
                               new SqlParameter { ParameterName = "@Policy_Endorsement", Value = Policy_Endorsement },
                               new SqlParameter { ParameterName = "@RFormat", Value = "F1" },
                               new SqlParameter { ParameterName = "@InvNo", Value = "RGIX" },
                               new SqlParameter { ParameterName = "@ReportId", Value = "RGIX" },
                               new SqlParameter { ParameterName = "@DocName", Value = "Reliance General Insurance Co. Ltd." }
                               );
                }
            }
        }

        //public static void InsertTransaction(Microsoft.Office.Interop.Excel.Workbook WB, string TranID, string RDate, string Insurance, string Salesby, string Serviceby, string location, string Support, string Rmonth)
        //{
        //    //Royal Sundaram General Insurance Co Ltd.
        //    SQLProcs sql = new SQLProcs();
        //    Microsoft.Office.Interop.Excel.Worksheet wks = (Microsoft.Office.Interop.Excel.Worksheet)WB.Worksheets[1];
        //    Microsoft.Office.Interop.Excel.Range lastCell = wks.Cells.SpecialCells(Microsoft.Office.Interop.Excel.XlCellType.xlCellTypeLastCell, Type.Missing);
        //    int lastrow = lastCell.Row; var Policy_Endorsement = ""; var InsuredType = ""; var Policy_Type = ""; var New_Renewal = "";
        //    for (int i = 2; i <= lastrow; i++)
        //    {
        //        var Terrorism = ""; var Premium_Amt = "";
        //        string InsuredName = ((Microsoft.Office.Interop.Excel.Range)wks.Cells[i, 10]).Value;
        //        if (InsuredName != null && InsuredName != "" && InsuredName != " ")
        //        {
        //            InsuredName = InsuredName.Replace("\n", "").TrimStart();
        //            InsuredType = ((Microsoft.Office.Interop.Excel.Range)wks.Cells[i, 19]).Value.Replace("\n", "").TrimStart();
        //            string PolicyNo = Convert.ToString(((Microsoft.Office.Interop.Excel.Range)wks.Cells[i, 7]).Value).Replace("\n", "").Replace("'", "").TrimStart();
        //            var Client_N_E = ((Microsoft.Office.Interop.Excel.Range)wks.Cells[i, 8]).Value.Replace("\n", "").TrimStart();
        //            var Endono = Convert.ToString(((Microsoft.Office.Interop.Excel.Range)wks.Cells[i, 14]).Value).Replace("\n", "").TrimStart();
        //            //PolicyNo = PolicyNo + "/" + Endono;
        //            if (Convert.ToDecimal(Endono) > 0)
        //            {
        //                PolicyNo = PolicyNo + "/" + Endono;
        //                Policy_Endorsement = "Policy";
        //            }
        //            else
        //            {
        //                Policy_Endorsement = "Endorsement";
        //            }
        //            if (Client_N_E == "Renewal")
        //            {
        //                Client_N_E = "Existing Client";
        //                New_Renewal = "Renewal Policy";
        //            }
        //            else
        //            {
        //                Client_N_E = "New Client";
        //                New_Renewal = "New Policy";
        //            }
        //            var Endo_Effective_Date = ((Microsoft.Office.Interop.Excel.Range)wks.Cells[i, 11]).Value;
        //            var Effective_Date = ((Microsoft.Office.Interop.Excel.Range)wks.Cells[i, 17]).Value;
        //            int ENdolen = Convert.ToString(Endo_Effective_Date).Length;
        //            if (ENdolen > 11)
        //            {
        //                Endo_Effective_Date = Endo_Effective_Date.ToString("dd/MM/yyyy");
        //            }
        //            int Efflen = Convert.ToString(Effective_Date).Length;
        //            if (Efflen > 11)
        //            {
        //                Effective_Date = Effective_Date.ToString("dd/MM/yyyy");
        //            }
        //            Policy_Type = Convert.ToString(((Microsoft.Office.Interop.Excel.Range)wks.Cells[i, 18]).Value);
        //            var Ridercode = Convert.ToString(((Microsoft.Office.Interop.Excel.Range)wks.Cells[i, 22]).Value);
        //            if (Ridercode == null || Ridercode == "" || Ridercode == " " || Ridercode.Contains("ADD-ON"))
        //            {
        //                Premium_Amt = Convert.ToString(((Microsoft.Office.Interop.Excel.Range)wks.Cells[i, 26]).Value);
        //            }
        //            else if (Ridercode.Contains("TP-COMP"))
        //            {
        //                //if (Policy_Type.ToUpper().Contains("VPD") || Policy_Type.ToUpper().Contains("VPB"))
        //                //{
        //                Terrorism = Convert.ToString(((Microsoft.Office.Interop.Excel.Range)wks.Cells[i, 26]).Value);
        //                //}
        //                //else
        //                //{

        //                //}
        //            }
        //            var Revenue_Amt = Convert.ToString(((Microsoft.Office.Interop.Excel.Range)wks.Cells[i, 33]).Value);
        //            var Revenue_Pcnt = Convert.ToString(((Microsoft.Office.Interop.Excel.Range)wks.Cells[i, 32]).Text);
        //            Revenue_Pcnt = Revenue_Pcnt.Replace("\n", "").Replace("%", "").Replace(",", "").TrimStart();


        //            if (Premium_Amt == "" || Premium_Amt == " " || Premium_Amt == null)
        //            {
        //                Premium_Amt = "0";
        //            }
        //            if (Revenue_Amt == "" || Revenue_Amt == " " || Revenue_Amt == null)
        //            {
        //                Revenue_Amt = 0;
        //            }
        //            if (Revenue_Pcnt == "" || Revenue_Pcnt == " " || Revenue_Pcnt == null)
        //            {
        //                Revenue_Pcnt = "0";
        //            }
        //            if (Terrorism == "" || Terrorism == " " || Terrorism == null)
        //            {
        //                Terrorism = "0";
        //            }
        //            sql.ExecuteSQLNonQuery("SP_RoyalTransactions",
        //                       new SqlParameter { ParameterName = "@Imode", Value = 1 },
        //                       new SqlParameter { ParameterName = "@RDate", Value = RDate },
        //                       new SqlParameter { ParameterName = "@Rmonth", Value = Rmonth },
        //                       new SqlParameter { ParameterName = "@ClientName", Value = InsuredName },
        //                       new SqlParameter { ParameterName = "@Itype", Value = InsuredType },
        //                       new SqlParameter { ParameterName = "@Policy_No", Value = PolicyNo },
        //                       new SqlParameter { ParameterName = "@Revenue_Pcnt", Value = Revenue_Pcnt },
        //                       new SqlParameter { ParameterName = "@Client_N_E", Value = Client_N_E },
        //                       new SqlParameter { ParameterName = "@New_Renewal", Value = New_Renewal },
        //                       new SqlParameter { ParameterName = "@Endo_Effective_Date", Value = Endo_Effective_Date },
        //                       new SqlParameter { ParameterName = "@Effective_Date", Value = Effective_Date },
        //                       //new SqlParameter { ParameterName = "@END_Date", Value = END_Date },
        //                       new SqlParameter { ParameterName = "@Policy_Type", Value = Policy_Type },
        //                       new SqlParameter { ParameterName = "@Premium_Amt", Value = Premium_Amt },
        //                       new SqlParameter { ParameterName = "@TranID", Value = TranID },
        //                       new SqlParameter { ParameterName = "@Revenue_Amt", Value = Revenue_Amt },
        //                       new SqlParameter { ParameterName = "@Terrorism", Value = Terrorism },
        //                       new SqlParameter { ParameterName = "@Insurance", Value = Insurance },
        //                       new SqlParameter { ParameterName = "@Salesby", Value = Salesby },
        //                       new SqlParameter { ParameterName = "@Serviceby", Value = Serviceby },
        //                       new SqlParameter { ParameterName = "@location", Value = location },
        //                       new SqlParameter { ParameterName = "@Support", Value = Support },
        //                       new SqlParameter { ParameterName = "@Policy_Endorsement", Value = Policy_Endorsement },
        //                       new SqlParameter { ParameterName = "@RFormat", Value = "F1" },
        //                       new SqlParameter { ParameterName = "@InvNo", Value = "RSGX" },
        //                       new SqlParameter { ParameterName = "@ReportId", Value = "RSGX" },
        //                       new SqlParameter { ParameterName = "@DocName", Value = "Royal Sundaram General Insurance Co Ltd." }
        //                       );
        //        }
        //    }
        //}

        //public static void InsertTransaction(Microsoft.Office.Interop.Excel.Workbook WB, string TranID, string RDate, string Insurance, string Salesby, string Serviceby, string location, string Support, string Rmonth)
        //{
        //    //SBI General Insurance Co. Ltd.
        //    SQLProcs sql = new SQLProcs();
        //    Microsoft.Office.Interop.Excel.Worksheet wks = (Microsoft.Office.Interop.Excel.Worksheet)WB.Worksheets[1];
        //    Microsoft.Office.Interop.Excel.Range lastCell = wks.Cells.SpecialCells(Microsoft.Office.Interop.Excel.XlCellType.xlCellTypeLastCell, Type.Missing);
        //    int lastrow = lastCell.Row; var Terrorism = ""; var Policy_Endorsement = ""; var InsuredType = ""; var Policy_Type = ""; var New_Renewal = ""; var Revenue_Amt = "";
        //    for (int i = 3; i <= lastrow; i++)
        //    {
        //        string InsuredName = ((Microsoft.Office.Interop.Excel.Range)wks.Cells[i, 11]).Value;
        //        if (InsuredName != null && InsuredName != "" && InsuredName != " ")
        //        {
        //            InsuredName = InsuredName.Replace("\n", "").TrimStart();
        //            InsuredType = ((Microsoft.Office.Interop.Excel.Range)wks.Cells[i, 8]).Value;
        //            string PolicyNo = Convert.ToString(((Microsoft.Office.Interop.Excel.Range)wks.Cells[i, 12]).Value).Replace("\n", "").TrimStart();
        //            var Endorsementno = Convert.ToString(((Microsoft.Office.Interop.Excel.Range)wks.Cells[i, 15]).Value);

        //            if (Endorsementno != "" && Endorsementno != " " && Endorsementno != null)
        //            {
        //                PolicyNo = PolicyNo + "/" + Endorsementno;
        //                if (Convert.ToDecimal(Endorsementno) > 0)
        //                {
        //                    Policy_Endorsement = "Endorsement";
        //                }
        //                else
        //                {
        //                    Policy_Endorsement = "Policy";
        //                }
        //            }
        //            else
        //            {
        //                Policy_Endorsement = "Policy";
        //            }
        //            var Endo_Effective_Date = ((Microsoft.Office.Interop.Excel.Range)wks.Cells[i, 2]).Value;
        //            var Effective_Date = ((Microsoft.Office.Interop.Excel.Range)wks.Cells[i, 13]).Value;
        //            var END_Date = ((Microsoft.Office.Interop.Excel.Range)wks.Cells[i, 14]).Value;
        //            int ENdolen = Convert.ToString(Endo_Effective_Date).Length;
        //            if (ENdolen > 11)
        //            {
        //                Endo_Effective_Date = Endo_Effective_Date.ToString("dd/MM/yyyy");
        //            }
        //            int Efflen = Convert.ToString(Effective_Date).Length;
        //            if (Efflen > 11)
        //            {
        //                Effective_Date = Effective_Date.ToString("dd/MM/yyyy");
        //            }
        //            int ENDlen = Convert.ToString(END_Date).Length;
        //            if (ENDlen > 11)
        //            {
        //                END_Date = END_Date.ToString("dd/MM/yyyy");
        //            }
        //            var Premium_Amt = Convert.ToString(((Microsoft.Office.Interop.Excel.Range)wks.Cells[i, 18]).Value).Replace(",", "").Replace("(", "").Replace(")", "").TrimStart();
        //            Terrorism = Convert.ToString(((Microsoft.Office.Interop.Excel.Range)wks.Cells[i, 23]).Value).Replace(",", "").Replace("(", "").Replace(")", "").TrimStart();
        //            Revenue_Amt = Convert.ToString(((Microsoft.Office.Interop.Excel.Range)wks.Cells[i, 26]).Value).Replace(",", "").Replace("(", "").Replace(")", "").Replace("-", "").TrimStart();
        //            var Revenue_Pcnt = Convert.ToString(((Microsoft.Office.Interop.Excel.Range)wks.Cells[i, 21]).Text);
        //            Revenue_Pcnt = Revenue_Pcnt.Replace("\n", "").Replace("%", "").Replace(",", "").TrimStart();

        //            Policy_Type = Convert.ToString(((Microsoft.Office.Interop.Excel.Range)wks.Cells[i, 7]).Value);

        //            if (Premium_Amt == "" || Premium_Amt == " " || Premium_Amt == null)
        //            {
        //                Premium_Amt = 0;
        //            }
        //            if (Revenue_Amt == "" || Revenue_Amt == " " || Revenue_Amt == null)
        //            {
        //                Revenue_Amt = "0";
        //            }
        //            if (Revenue_Pcnt == "" || Revenue_Pcnt == " " || Revenue_Pcnt == null)
        //            {
        //                Revenue_Pcnt = "0";
        //            }
        //            if (Terrorism == "" || Terrorism == " " || Terrorism == null)
        //            {
        //                Terrorism = "0";
        //            }
        //            sql.ExecuteSQLNonQuery("SP_SBITransactions",
        //                       new SqlParameter { ParameterName = "@Imode", Value = 1 },
        //                       new SqlParameter { ParameterName = "@RDate", Value = RDate },
        //                       new SqlParameter { ParameterName = "@Rmonth", Value = Rmonth },
        //                       new SqlParameter { ParameterName = "@ClientName", Value = InsuredName },
        //                       new SqlParameter { ParameterName = "@Itype", Value = InsuredType },
        //                       new SqlParameter { ParameterName = "@Policy_No", Value = PolicyNo },
        //                       new SqlParameter { ParameterName = "@Revenue_Pcnt", Value = Revenue_Pcnt },
        //                       new SqlParameter { ParameterName = "@Policy_Type", Value = Policy_Type },
        //                       new SqlParameter { ParameterName = "@Premium_Amt", Value = Premium_Amt },
        //                       new SqlParameter { ParameterName = "@TranID", Value = TranID },
        //                       new SqlParameter { ParameterName = "@Revenue_Amt", Value = Revenue_Amt },
        //                       new SqlParameter { ParameterName = "@Terrorism", Value = Terrorism },
        //                       new SqlParameter { ParameterName = "@Insurance", Value = Insurance },
        //                       new SqlParameter { ParameterName = "@Salesby", Value = Salesby },
        //                       new SqlParameter { ParameterName = "@Serviceby", Value = Serviceby },
        //                       new SqlParameter { ParameterName = "@location", Value = location },
        //                       new SqlParameter { ParameterName = "@Support", Value = Support },
        //                       new SqlParameter { ParameterName = "@Policy_Endorsement", Value = Policy_Endorsement },
        //                       new SqlParameter { ParameterName = "@RFormat", Value = "F1" },
        //                       new SqlParameter { ParameterName = "@InvNo", Value = "SBIX" },
        //                       new SqlParameter { ParameterName = "@ReportId", Value = "SBIX" },
        //                       new SqlParameter { ParameterName = "@DocName", Value = "SBI General Insurance Co. Ltd." }
        //                       );
        //        }
        //    }
        //}
    }
}
