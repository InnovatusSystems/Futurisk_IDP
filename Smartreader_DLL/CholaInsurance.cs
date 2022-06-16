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
    public class CholaInsurance
    {
        public static void InsertTransaction(Microsoft.Office.Interop.Excel.Workbook WB, string TranID, string RDate, string Insurance, string Salesby, string Serviceby, string location, string Support, string Rmonth, string strconn)
        {
            //Chola MS General Insurance
            SQLProcs sql = new SQLProcs();
            Microsoft.Office.Interop.Excel.Worksheet wks = (Microsoft.Office.Interop.Excel.Worksheet)WB.Worksheets[1];
            Microsoft.Office.Interop.Excel.Range lastCell = wks.Cells.SpecialCells(Microsoft.Office.Interop.Excel.XlCellType.xlCellTypeLastCell, Type.Missing);
            int lastrow = lastCell.Row;
            for (int i = 2; i <= lastrow; i++)
            {
                var Terrorism = ""; var Policy_Endorsement = ""; var InsuredType = ""; var Policy_Type = ""; var Reward_Pcnt = ""; var Reward_Amt = "";
                string InsuredName = ((Microsoft.Office.Interop.Excel.Range)wks.Cells[i, 11]).Value;
                if (InsuredName != null && InsuredName != "" && InsuredName != " ")
                {
                    InsuredName = InsuredName.Replace("\n", "").TrimStart();
                    InsuredName = Regex.Match(InsuredName, @"\D+").Value;
                    var Itype = ((Microsoft.Office.Interop.Excel.Range)wks.Cells[i, 10]).Value;
                    var Itype1 = ((Microsoft.Office.Interop.Excel.Range)wks.Cells[i, 47]).Value;
                    if ((Itype == null || Itype == "") && (Itype1 != null && Itype1 != ""))
                    {
                        InsuredType = Itype1;
                    }
                    else if ((Itype1 == null || Itype1 == "") && (Itype != null && Itype != ""))
                    {
                        InsuredType = Itype;
                    }
                    else
                    {
                        InsuredType = Itype;
                    }
                    string PolicyNo = ((Microsoft.Office.Interop.Excel.Range)wks.Cells[i, 5]).Value.Replace("\n", "").TrimStart();
                    var Presult = PolicyNo.Substring(PolicyNo.LastIndexOf('/') + 1);
                    if (PolicyNo.Contains("/"))
                    {
                        var resultString = Regex.Match(Presult, @"\d+").Value;
                        if (Convert.ToDouble(resultString) > 0)
                        {
                            Policy_Endorsement = "Endorsement";
                        }
                        else
                        {
                            Policy_Endorsement = "Policy";
                        }
                    }

                    var Endo_Effective_Date = ((Microsoft.Office.Interop.Excel.Range)wks.Cells[i, 4]).Value;
                    //var Effective_Date = ((Microsoft.Office.Interop.Excel.Range)wks.Cells[i, 12]).Value;
                    var END_Date = ((Microsoft.Office.Interop.Excel.Range)wks.Cells[i, 43]).Value;
                    int ENdolen = Convert.ToString(Endo_Effective_Date).Length;
                    if (ENdolen > 11)
                    {
                        Endo_Effective_Date = Endo_Effective_Date.ToString("dd/MM/yyyy");
                    }
                    //int Efflen = Convert.ToString(Effective_Date).Length;
                    //if (Efflen > 11)
                    //{
                    //    Effective_Date = Effective_Date.ToString("dd/MM/yyyy");
                    //}
                    int ENDlen = Convert.ToString(END_Date).Length;
                    if (ENDlen > 11)
                    {
                        END_Date = END_Date.ToString("dd/MM/yyyy");
                    }
                    //var END_Date = ((Microsoft.Office.Interop.Excel.Range)wks.Cells[i, 5]).Value;
                    //Endo_Effective_Date = Endo_Effective_Date.ToString("dd/MM/yyyy");
                    //Effective_Date = Effective_Date.ToString("dd/MM/yyyy");
                    //END_Date = END_Date.ToString("dd/MM/yyyy");
                    //Policy_Endorsement = ((Microsoft.Office.Interop.Excel.Range)wks.Cells[i, 11]).Value.Replace("\n", "").TrimStart();

                    var Premium_Amt = Convert.ToString(((Microsoft.Office.Interop.Excel.Range)wks.Cells[i, 19]).Value);
                    var Revenue_Amt = Convert.ToString(((Microsoft.Office.Interop.Excel.Range)wks.Cells[i, 21]).Value);
                    var TPRevenue_Amt = Convert.ToString(((Microsoft.Office.Interop.Excel.Range)wks.Cells[i, 46]).Value);
                    var Revenue_Pcnt = Convert.ToString(((Microsoft.Office.Interop.Excel.Range)wks.Cells[i, 20]).Text);
                    Revenue_Pcnt = Revenue_Pcnt.Replace("\n", "").Replace("%", "").Replace(",", "").TrimStart();
                    Terrorism = Convert.ToString(((Microsoft.Office.Interop.Excel.Range)wks.Cells[i, 45]).Value);

                    var Reward_Pcnt1 = Convert.ToString(((Microsoft.Office.Interop.Excel.Range)wks.Cells[i, 22]).Value);
                    var Reward_Pcnt2 = Convert.ToString(((Microsoft.Office.Interop.Excel.Range)wks.Cells[i, 24]).Value);
                    var Reward_Pcnt3 = Convert.ToString(((Microsoft.Office.Interop.Excel.Range)wks.Cells[i, 26]).Value);

                    var Reward_Amt1 = Convert.ToString(((Microsoft.Office.Interop.Excel.Range)wks.Cells[i, 23]).Value);
                    var Reward_Amt2 = Convert.ToString(((Microsoft.Office.Interop.Excel.Range)wks.Cells[i, 25]).Value);
                    var Reward_Amt3 = Convert.ToString(((Microsoft.Office.Interop.Excel.Range)wks.Cells[i, 27]).Value);

                    if (Reward_Pcnt1 == "" || Reward_Pcnt1 == " " || Reward_Pcnt1 == null)
                    {
                        Reward_Pcnt1 = 0;
                    }
                    if (Reward_Pcnt2 == "" || Reward_Pcnt2 == " " || Reward_Pcnt2 == null)
                    {
                        Reward_Pcnt2 = 0;
                    }
                    if (Reward_Pcnt3 == "" || Reward_Pcnt3 == " " || Reward_Pcnt3 == null)
                    {
                        Reward_Pcnt3 = 0;
                    }
                    if (Reward_Amt1 == "" || Reward_Amt1 == " " || Reward_Amt1 == null)
                    {
                        Reward_Amt1 = 0;
                    }
                    if (Reward_Amt2 == "" || Reward_Amt2 == " " || Reward_Amt2 == null)
                    {
                        Reward_Amt2 = 0;
                    }
                    if (Reward_Amt3 == "" || Reward_Amt3 == " " || Reward_Amt3 == null)
                    {
                        Reward_Amt3 = 0;
                    }

                    Reward_Pcnt = Convert.ToString(Convert.ToDouble(Reward_Pcnt1) + Convert.ToDouble(Reward_Pcnt2) + Convert.ToDouble(Reward_Pcnt3));
                    Reward_Amt = Convert.ToString(Convert.ToDouble(Reward_Amt1) + Convert.ToDouble(Reward_Amt2) + Convert.ToDouble(Reward_Amt3));

                    Policy_Type = Convert.ToString(((Microsoft.Office.Interop.Excel.Range)wks.Cells[i, 6]).Value);

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
                    if (Terrorism == "" || Terrorism == " " || Terrorism == null)
                    {
                        Terrorism = "0";
                    }
                    if (TPRevenue_Amt == "" || TPRevenue_Amt == " " || TPRevenue_Amt == null)
                    {
                        TPRevenue_Amt = "0";
                    }
                    sql.ExecuteSQLNonQuery(strconn,"SP_CholaTransactions",
                               new SqlParameter { ParameterName = "@Imode", Value = 1 },
                               new SqlParameter { ParameterName = "@RDate", Value = RDate },
                               new SqlParameter { ParameterName = "@Rmonth", Value = Rmonth },
                               new SqlParameter { ParameterName = "@ClientName", Value = InsuredName },
                               new SqlParameter { ParameterName = "@Itype", Value = InsuredType },
                               new SqlParameter { ParameterName = "@Policy_No", Value = PolicyNo },
                               new SqlParameter { ParameterName = "@Revenue_Pcnt", Value = Revenue_Pcnt },
                               new SqlParameter { ParameterName = "@Endo_Effective_Date", Value = Endo_Effective_Date },
                               //new SqlParameter { ParameterName = "@Effective_Date", Value = Effective_Date },
                               new SqlParameter { ParameterName = "@END_Date", Value = END_Date },
                               new SqlParameter { ParameterName = "@Policy_Type", Value = Policy_Type },
                               new SqlParameter { ParameterName = "@Premium_Amt", Value = Premium_Amt },
                               new SqlParameter { ParameterName = "@TPRevenue_Amt", Value = TPRevenue_Amt },
                               new SqlParameter { ParameterName = "@Reward_Pcnt", Value = Reward_Pcnt },
                               new SqlParameter { ParameterName = "@Reward_Amt", Value = Reward_Amt },
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
                               new SqlParameter { ParameterName = "@InvNo", Value = "CMGX" },
                               new SqlParameter { ParameterName = "@ReportId", Value = "CMGX" },
                               new SqlParameter { ParameterName = "@DocName", Value = "Chola MS General Insurance Co. Ltd." }
                               );
                }
            }
        }
    }
}
