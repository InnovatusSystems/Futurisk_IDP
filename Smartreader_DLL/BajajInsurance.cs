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
    public class BajajInsurance
    {
        public static void InsertTransaction(Microsoft.Office.Interop.Excel.Workbook WB, string TranID, string RDate, string Insurance, string Salesby, string Serviceby, string location, string Support, string Rmonth, string strconn)
        {
            //Bajaj insurance
            SQLProcs sql = new SQLProcs();
            Microsoft.Office.Interop.Excel.Worksheet wks = (Microsoft.Office.Interop.Excel.Worksheet)WB.Worksheets[1];
            Microsoft.Office.Interop.Excel.Range lastCell = wks.Cells.SpecialCells(Microsoft.Office.Interop.Excel.XlCellType.xlCellTypeLastCell, Type.Missing);
            int lastrow = lastCell.Row; 
            for (int i = 2; i <= lastrow; i++)
            {
                var Terrorism = ""; var Policy_Endorsement = ""; var InsuredType = ""; var Policy_Type = "";
                string InsuredName = ((Microsoft.Office.Interop.Excel.Range)wks.Cells[i, 26]).Value;
                if (InsuredName != null && InsuredName != "" && InsuredName != " ")
                {
                    InsuredName = InsuredName.Replace("\n", "").TrimStart();
                    string PolicyNo = ((Microsoft.Office.Interop.Excel.Range)wks.Cells[i, 2]).Value.Replace("\n", "").TrimStart();
                    var Endo_Effective_Date = ((Microsoft.Office.Interop.Excel.Range)wks.Cells[i, 3]).Value;
                    var Effective_Date = ((Microsoft.Office.Interop.Excel.Range)wks.Cells[i, 35]).Value;
                    int ENdolen = Convert.ToString(Endo_Effective_Date).Length;
                    if (ENdolen > 11)
                    {
                        Endo_Effective_Date = Endo_Effective_Date.ToString("dd/MM/yyyy");
                    }
                    if (Effective_Date != null)
                    {
                        int Efflen = Convert.ToString(Effective_Date).Length;
                        if (Efflen > 11)
                        {
                            Effective_Date = Effective_Date.ToString("dd/MM/yyyy");
                        }
                    }
                    var Presult = PolicyNo.Substring(PolicyNo.LastIndexOf('-') + 1);
                    if (PolicyNo.Contains("-"))
                    {
                        var resultString = Regex.Match(Presult, @"\d+").Value;
                        if (Convert.ToDouble(resultString) > 0)
                        {
                            if (PolicyNo.Contains("EE") || PolicyNo.Contains("ER"))
                            {
                                Policy_Endorsement = "Endorsement";
                            }
                            else
                            {
                                Policy_Endorsement = "Policy";
                            }
                        }
                        else
                        {
                            Policy_Endorsement = "Policy";
                        }
                    }
                    var Premium_Amt = Convert.ToString(((Microsoft.Office.Interop.Excel.Range)wks.Cells[i, 9]).Value);
                    var Revenue_Amt = Convert.ToString(((Microsoft.Office.Interop.Excel.Range)wks.Cells[i, 15]).Value);
                    var Revenue_Pcnt = Convert.ToString(((Microsoft.Office.Interop.Excel.Range)wks.Cells[i, 5]).Text);
                    Revenue_Pcnt = Revenue_Pcnt.Replace("\n", "").Replace("%", "").Replace(",", "").TrimStart();
                    var TP1 = Convert.ToString(((Microsoft.Office.Interop.Excel.Range)wks.Cells[i, 11]).Value);
                    var TP = Convert.ToString(((Microsoft.Office.Interop.Excel.Range)wks.Cells[i, 10]).Value);

                    Policy_Type = Convert.ToString(((Microsoft.Office.Interop.Excel.Range)wks.Cells[i, 19]).Value);

                    if (InsuredName.Contains("LIMITED") || InsuredName.Contains("LTD") || InsuredName.Contains(".COM"))
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
                    if (Revenue_Amt == "" || Revenue_Amt == " " || Revenue_Amt == null)
                    {
                        Revenue_Amt = 0;
                    }
                    if (Revenue_Pcnt == "" || Revenue_Pcnt == " " || Revenue_Pcnt == null)
                    {
                        Revenue_Pcnt = "0";
                    }
                    if (TP1 == "" || TP1 == " " || TP1 == null)
                    {
                        TP1 = "0";
                    }
                    if (TP == "" || TP == " " || TP == null)
                    {
                        TP = "0";
                    }
                    if (TP == "0" && TP1 != "0")
                    {
                        Terrorism = TP1;
                    }
                    else if (TP != "0" && TP1 == "0")
                    {
                        Terrorism = TP;
                    }
                    if (Terrorism == "" || Terrorism == " " || Terrorism == null)
                    {
                        Terrorism = "0";
                    }
                    sql.ExecuteSQLNonQuery(strconn,"SP_BajajTransactions",
                               new SqlParameter { ParameterName = "@Imode", Value = 1 },
                               new SqlParameter { ParameterName = "@RDate", Value = RDate },
                               new SqlParameter { ParameterName = "@Rmonth", Value = Rmonth },
                               new SqlParameter { ParameterName = "@ClientName", Value = InsuredName },
                               new SqlParameter { ParameterName = "@Itype", Value = InsuredType },
                               new SqlParameter { ParameterName = "@Policy_No", Value = PolicyNo },
                               new SqlParameter { ParameterName = "@Revenue_Pcnt", Value = Revenue_Pcnt },
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
                               new SqlParameter { ParameterName = "@InvNo", Value = "BAGX" },
                               new SqlParameter { ParameterName = "@ReportId", Value = "BAGX" },
                               new SqlParameter { ParameterName = "@DocName", Value = "Bajaj Allianz General Insurance Co. Ltd." }
                               );
                }
            }
        }
    }
}
