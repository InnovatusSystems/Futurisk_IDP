using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using NPOI.SS.UserModel;
using System.Text;
using System.Threading.Tasks;
using System.Text.RegularExpressions;


namespace Smartreader_DLL
{
    public class RoyalInsurance
    {
        public static void InsertTransaction(Microsoft.Office.Interop.Excel.Workbook WB, string TranID, string RDate, string Insurance, string Salesby, string Serviceby, string location, string Support, string Rmonth, string strconn)
        {
            //Royal Sundaram General Insurance Co Ltd.
            SQLProcs sql = new SQLProcs();
            Microsoft.Office.Interop.Excel.Worksheet wks = (Microsoft.Office.Interop.Excel.Worksheet)WB.Worksheets[1];
            Microsoft.Office.Interop.Excel.Range lastCell = wks.Cells.SpecialCells(Microsoft.Office.Interop.Excel.XlCellType.xlCellTypeLastCell, Type.Missing);
            int lastrow = lastCell.Row; 
            for (int i = 2; i <= lastrow; i++)
            {
                var Terrorism = ""; var Premium_Amt = ""; var Policy_Endorsement = ""; var InsuredType = ""; var Policy_Type = ""; var New_Renewal = "";
                string InsuredName = ((Microsoft.Office.Interop.Excel.Range)wks.Cells[i, 10]).Value;
                if (InsuredName != null && InsuredName != "" && InsuredName != " ")
                {
                    InsuredName = InsuredName.Replace("\n", "").TrimStart();
                    InsuredType = ((Microsoft.Office.Interop.Excel.Range)wks.Cells[i, 19]).Value.Replace("\n", "").TrimStart();
                    string PolicyNo = Convert.ToString(((Microsoft.Office.Interop.Excel.Range)wks.Cells[i, 7]).Value).Replace("\n", "").Replace("'", "").TrimStart();
                    var Client_N_E = ((Microsoft.Office.Interop.Excel.Range)wks.Cells[i, 8]).Value.Replace("\n", "").TrimStart();
                    var Endono = Convert.ToString(((Microsoft.Office.Interop.Excel.Range)wks.Cells[i, 14]).Value).Replace("\n", "").TrimStart();
                    //PolicyNo = PolicyNo + "/" + Endono;
                    if (Convert.ToDecimal(Endono) > 0)
                    {
                        PolicyNo = PolicyNo + "/" + Endono;
                        Policy_Endorsement = "Policy";
                    }
                    else
                    {
                        Policy_Endorsement = "Endorsement";
                    }
                    if (Client_N_E == "Renewal")
                    {
                        Client_N_E = "Existing Client";
                        New_Renewal = "Renewal Policy";
                    }
                    else
                    {
                        Client_N_E = "New Client";
                        New_Renewal = "New Policy";
                    }
                    var Endo_Effective_Date = ((Microsoft.Office.Interop.Excel.Range)wks.Cells[i, 11]).Value;
                    var Effective_Date = ((Microsoft.Office.Interop.Excel.Range)wks.Cells[i, 17]).Value;
                    int ENdolen = Convert.ToString(Endo_Effective_Date).Length;
                    if (ENdolen > 11)
                    {
                        Endo_Effective_Date = Endo_Effective_Date.ToString("dd/MM/yyyy");
                    }
                    int Efflen = Convert.ToString(Effective_Date).Length;
                    if (Efflen > 11)
                    {
                        Effective_Date = Effective_Date.ToString("dd/MM/yyyy");
                    }
                    Policy_Type = Convert.ToString(((Microsoft.Office.Interop.Excel.Range)wks.Cells[i, 18]).Value);
                    var Ridercode = Convert.ToString(((Microsoft.Office.Interop.Excel.Range)wks.Cells[i, 22]).Value);
                    if (Ridercode == null || Ridercode == "" || Ridercode == " " || Ridercode.Contains("ADD-ON"))
                    {
                        Premium_Amt = Convert.ToString(((Microsoft.Office.Interop.Excel.Range)wks.Cells[i, 26]).Value);
                    }
                    else if (Ridercode.Contains("TP-COMP"))
                    {
                        //if (Policy_Type.ToUpper().Contains("VPD") || Policy_Type.ToUpper().Contains("VPB"))
                        //{
                        Terrorism = Convert.ToString(((Microsoft.Office.Interop.Excel.Range)wks.Cells[i, 26]).Value);
                        //}
                        //else
                        //{

                        //}
                    }
                    var Revenue_Amt = Convert.ToString(((Microsoft.Office.Interop.Excel.Range)wks.Cells[i, 33]).Value);
                    var Revenue_Pcnt = Convert.ToString(((Microsoft.Office.Interop.Excel.Range)wks.Cells[i, 32]).Text);
                    Revenue_Pcnt = Revenue_Pcnt.Replace("\n", "").Replace("%", "").Replace(",", "").TrimStart();


                    if (Premium_Amt == "" || Premium_Amt == " " || Premium_Amt == null)
                    {
                        Premium_Amt = "0";
                    }
                    if (Revenue_Amt == "" || Revenue_Amt == " " || Revenue_Amt == null)
                    {
                        Revenue_Amt = 0;
                    }
                    if (Revenue_Pcnt == "" || Revenue_Pcnt == " " || Revenue_Pcnt == null)
                    {
                        Revenue_Pcnt = "0";
                    }
                    if (Terrorism == "" || Terrorism == " " || Terrorism == null)
                    {
                        Terrorism = "0";
                    }
                    sql.ExecuteSQLNonQuery(strconn,"SP_RoyalTransactions",
                               new SqlParameter { ParameterName = "@Imode", Value = 1 },
                               new SqlParameter { ParameterName = "@RDate", Value = RDate },
                               new SqlParameter { ParameterName = "@Rmonth", Value = Rmonth },
                               new SqlParameter { ParameterName = "@ClientName", Value = InsuredName },
                               new SqlParameter { ParameterName = "@Itype", Value = InsuredType },
                               new SqlParameter { ParameterName = "@Policy_No", Value = PolicyNo },
                               new SqlParameter { ParameterName = "@Revenue_Pcnt", Value = Revenue_Pcnt },
                               new SqlParameter { ParameterName = "@Client_N_E", Value = Client_N_E },
                               new SqlParameter { ParameterName = "@New_Renewal", Value = New_Renewal },
                               new SqlParameter { ParameterName = "@Endo_Effective_Date", Value = Endo_Effective_Date },
                               new SqlParameter { ParameterName = "@Effective_Date", Value = Effective_Date },
                               //new SqlParameter { ParameterName = "@END_Date", Value = END_Date },
                               new SqlParameter { ParameterName = "@Policy_Type", Value = Policy_Type },
                               new SqlParameter { ParameterName = "@Premium_Amt", Value = Premium_Amt },
                               new SqlParameter { ParameterName = "@TranID", Value = TranID },
                               new SqlParameter { ParameterName = "@Revenue_Amt", Value = Revenue_Amt },
                               new SqlParameter { ParameterName = "@Terrorism", Value = Terrorism },
                               new SqlParameter { ParameterName = "@Insurance", Value = Insurance },
                               new SqlParameter { ParameterName = "@Salesby", Value = Salesby },
                               new SqlParameter { ParameterName = "@Serviceby", Value = Serviceby },
                               new SqlParameter { ParameterName = "@location", Value = location },
                               new SqlParameter { ParameterName = "@Support", Value = Support },
                               new SqlParameter { ParameterName = "@Policy_Endorsement", Value = Policy_Endorsement },
                               new SqlParameter { ParameterName = "@RFormat", Value = "F1" },
                               new SqlParameter { ParameterName = "@InvNo", Value = "RSGX" },
                               new SqlParameter { ParameterName = "@ReportId", Value = "RSGX" },
                               new SqlParameter { ParameterName = "@DocName", Value = "Royal Sundaram General Insurance Co Ltd." }
                               );
                }
            }
        }
    }
}
