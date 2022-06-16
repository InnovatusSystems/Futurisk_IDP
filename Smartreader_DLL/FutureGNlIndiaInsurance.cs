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
    public class FutureGNlIndiaInsurance
    {
        public static void InsertTransaction(Microsoft.Office.Interop.Excel.Workbook WB, string TranID, string RDate, string Insurance, string Salesby, string Serviceby, string location, string Support, string Rmonth, string strconn)
        {
            //Future Generali India Insurance Co. Ltd. Excel
            SQLProcs sql = new SQLProcs();
            Microsoft.Office.Interop.Excel.Worksheet wks = (Microsoft.Office.Interop.Excel.Worksheet)WB.Worksheets[1];
            Microsoft.Office.Interop.Excel.Range lastCell = wks.Cells.SpecialCells(Microsoft.Office.Interop.Excel.XlCellType.xlCellTypeLastCell, Type.Missing);
            int lastrow = lastCell.Row;
            for (int i = 2; i <= lastrow; i++)
            {
                var Terrorism = ""; var Policy_Endorsement = ""; var InsuredType = ""; var Policy_Type = ""; var New_Renewal = ""; var Revenue_Amt = "";
                string InsuredName = ((Microsoft.Office.Interop.Excel.Range)wks.Cells[i, 39]).Value;
                if (InsuredName != null && InsuredName != "" && InsuredName != " ")
                {
                    InsuredName = InsuredName.Replace("\n", "").TrimStart();
                    InsuredType = Convert.ToString(((Microsoft.Office.Interop.Excel.Range)wks.Cells[i, 36]).Value);
                    string PolicyNo = Convert.ToString(((Microsoft.Office.Interop.Excel.Range)wks.Cells[i, 3]).Value);
                    if (PolicyNo == null || PolicyNo == "")
                    {
                        PolicyNo = Convert.ToString(((Microsoft.Office.Interop.Excel.Range)wks.Cells[i, 2]).Value).Replace("\n", "").TrimStart();
                    }
                    else
                    {
                        PolicyNo = PolicyNo.Replace("\n", "").TrimStart();
                    }
                    if (InsuredType == "Corporate Broking")
                    {
                        InsuredType = "Corporate";
                    }
                    else
                    {
                        InsuredType = "Retail";
                    }

                    var Client_N_E = ((Microsoft.Office.Interop.Excel.Range)wks.Cells[i, 13]).Value.Replace("\n", "").TrimStart();
                    if (Client_N_E == "Endorsement Issue")
                    {
                        Policy_Endorsement = "Endorsement";
                    }
                    else
                    {
                        Policy_Endorsement = "Policy";
                    }
                    if (Client_N_E == "New Business Issue")
                    {
                        Client_N_E = "New Client";
                        if (Policy_Endorsement != "Endorsement")
                        {
                            New_Renewal = "New Policy";
                        }
                    }
                    else
                    {
                        Client_N_E = "Existing Client";
                        if (Policy_Endorsement != "Endorsement")
                        {
                            New_Renewal = "Renewal Policy";
                        }
                    }
                    var Endo_Effective_Date = ((Microsoft.Office.Interop.Excel.Range)wks.Cells[i, 31]).Value;
                    var Effective_Date = ((Microsoft.Office.Interop.Excel.Range)wks.Cells[i, 32]).Value;
                    var END_Date = ((Microsoft.Office.Interop.Excel.Range)wks.Cells[i, 33]).Value;
                    int ENdolen = Convert.ToString(Endo_Effective_Date).Length;
                    if (ENdolen == 8)
                    {
                        string date = Convert.ToString(Endo_Effective_Date);
                        var yy = date.Substring(0, 4);
                        var MM = date.Substring(4, 2);
                        var DD = date.Substring(6, 2);
                        Endo_Effective_Date = DD + "/" + MM + "/" + yy;
                    }
                    int Efflen = Convert.ToString(Effective_Date).Length;
                    if (Efflen == 8)
                    {
                        string date = Convert.ToString(Effective_Date);
                        var yy = date.Substring(0, 4);
                        var MM = date.Substring(4, 2);
                        var DD = date.Substring(6, 2);
                        Effective_Date = DD + "/" + MM + "/" + yy;
                        //Effective_Date = Effective_Date.ToString("dd/MM/yyyy");
                    }
                    int ENDlen = Convert.ToString(END_Date).Length;
                    if (ENDlen == 8)
                    {
                        string date = Convert.ToString(END_Date);
                        var yy = date.Substring(0, 4);
                        var MM = date.Substring(4, 2);
                        var DD = date.Substring(6, 2);
                        END_Date = DD + "/" + MM + "/" + yy;
                        //END_Date = END_Date.ToString("dd/MM/yyyy");
                    }
                    var Premium_Amt = Convert.ToString(((Microsoft.Office.Interop.Excel.Range)wks.Cells[i, 22]).Value).Replace(",", "").Replace("(", "").Replace(")", "").TrimStart();
                    Terrorism = Convert.ToString(((Microsoft.Office.Interop.Excel.Range)wks.Cells[i, 45]).Value).Replace(",", "").Replace("(", "").Replace(")", "").TrimStart();
                    Revenue_Amt = Convert.ToString(((Microsoft.Office.Interop.Excel.Range)wks.Cells[i, 23]).Value).Replace(",", "").Replace("(", "").Replace(")", "").Replace("-", "").TrimStart();
                    var Revenue_Pcnt = Convert.ToString(((Microsoft.Office.Interop.Excel.Range)wks.Cells[i, 56]).Text);
                    Revenue_Pcnt = Revenue_Pcnt.Replace("\n", "").Replace("%", "").Replace(",", "").TrimStart();
                    var RewardAmt = Convert.ToString(((Microsoft.Office.Interop.Excel.Range)wks.Cells[i, 75]).Value).Replace(",", "").Replace("(", "").Replace(")", "").TrimStart();
                    //Policy_Type = Convert.ToString(((Microsoft.Office.Interop.Excel.Range)wks.Cells[i, 7]).Value);
                    var offlocation = Convert.ToString(((Microsoft.Office.Interop.Excel.Range)wks.Cells[i, 18]).Value).Replace("\n", "").TrimStart();
                    if (Premium_Amt == "" || Premium_Amt == " " || Premium_Amt == null)
                    {
                        Premium_Amt = 0;
                    }
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
                    if (RewardAmt == "" || RewardAmt == " " || RewardAmt == null)
                    {
                        RewardAmt = "0";
                    }
                    sql.ExecuteSQLNonQuery(strconn,"SP_FutureGNlIndiaTransaction",
                               new SqlParameter { ParameterName = "@Imode", Value = 1 },
                               new SqlParameter { ParameterName = "@RDate", Value = RDate },
                               new SqlParameter { ParameterName = "@Rmonth", Value = Rmonth },
                               new SqlParameter { ParameterName = "@ClientName", Value = InsuredName },
                               new SqlParameter { ParameterName = "@Itype", Value = InsuredType },
                               new SqlParameter { ParameterName = "@Policy_No", Value = PolicyNo },
                               new SqlParameter { ParameterName = "@Endo_Effective_Date", Value = Endo_Effective_Date },
                               new SqlParameter { ParameterName = "@Effective_Date", Value = Effective_Date },
                               new SqlParameter { ParameterName = "@END_Date", Value = END_Date },
                               new SqlParameter { ParameterName = "@Client_N_E", Value = Client_N_E },
                               new SqlParameter { ParameterName = "@New_Renewal", Value = New_Renewal },
                               new SqlParameter { ParameterName = "@Revenue_Pct", Value = Revenue_Pcnt },
                               new SqlParameter { ParameterName = "@Policy_Type", Value = Policy_Type },
                               new SqlParameter { ParameterName = "@Premium_Amt", Value = Premium_Amt },
                               new SqlParameter { ParameterName = "@RewardAmt", Value = RewardAmt },
                               new SqlParameter { ParameterName = "@TranID", Value = TranID },
                               new SqlParameter { ParameterName = "@Revenue_Amt", Value = Revenue_Amt },
                               new SqlParameter { ParameterName = "@Terrorism", Value = Terrorism },
                               new SqlParameter { ParameterName = "@Insurance", Value = Insurance },
                               new SqlParameter { ParameterName = "@Salesby", Value = Salesby },
                               new SqlParameter { ParameterName = "@Serviceby", Value = Serviceby },
                               new SqlParameter { ParameterName = "@location", Value = offlocation },
                               new SqlParameter { ParameterName = "@Support", Value = Support },
                               new SqlParameter { ParameterName = "@Policy_Endorsement", Value = Policy_Endorsement },
                               new SqlParameter { ParameterName = "@RFormat", Value = "F1" },
                               new SqlParameter { ParameterName = "@InvNo", Value = "FGIX" },
                               new SqlParameter { ParameterName = "@ReportId", Value = "FGIX" },
                               new SqlParameter { ParameterName = "@DocName", Value = "Future General India Insurance Co.Ltd." }
                               );
                }
            }
        }
    }
}
