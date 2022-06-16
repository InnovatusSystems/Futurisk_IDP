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
    public class StarHealthInsurence1
    {
        public static void InsertTransaction(Microsoft.Office.Interop.Excel.Workbook WB, string TranID, string RDate, string Insurance, string Salesby, string Serviceby, string location, string Support, string Rmonth, string strconn)
        {
            SQLProcs sql = new SQLProcs();
            Microsoft.Office.Interop.Excel.Worksheet wks = (Microsoft.Office.Interop.Excel.Worksheet)WB.Worksheets[1];
            Microsoft.Office.Interop.Excel.Range lastCell = wks.Cells.SpecialCells(Microsoft.Office.Interop.Excel.XlCellType.xlCellTypeLastCell, Type.Missing);
            int lastrow = lastCell.Row; 
            for (int i = 2; i <= lastrow; i++)
            {
                var Terrorism = ""; var Policy_Endorsement = ""; var InsuredType = "";
                string InsuredName = ((Microsoft.Office.Interop.Excel.Range)wks.Cells[i, 3]).Value;
                if (InsuredName != null && InsuredName != "" && InsuredName != " ")
                {
                    InsuredName = InsuredName.Replace("\n", "").TrimStart();
                    var Client_N_E = ((Microsoft.Office.Interop.Excel.Range)wks.Cells[i, 18]).Value.Replace("\n", "").TrimStart();
                    if (Client_N_E == "FRESH")
                    {
                        Client_N_E = "New";
                    }
                    string PolicyNo = ((Microsoft.Office.Interop.Excel.Range)wks.Cells[i, 2]).Value.Replace("\n", "").TrimStart();
                    var Endo_Effective_Date = ((Microsoft.Office.Interop.Excel.Range)wks.Cells[i, 13]).Value;
                    var Effective_Date = ((Microsoft.Office.Interop.Excel.Range)wks.Cells[i, 11]).Value;
                    var END_Date = ((Microsoft.Office.Interop.Excel.Range)wks.Cells[i, 12]).Value;
                    int ENdolen = Convert.ToString(Endo_Effective_Date).Length;
                    int Efflen = Convert.ToString(Effective_Date).Length;
                    int ENdlen = Convert.ToString(END_Date).Length;
                    if (ENdolen > 11)
                    {
                        Endo_Effective_Date = Endo_Effective_Date.ToString("dd/MM/yyyy");
                    }
                    if (Efflen > 11)
                    {
                        Effective_Date = Effective_Date.ToString("dd/MM/yyyy");
                    }
                    if (ENdlen > 11)
                    {
                        END_Date = END_Date.ToString("dd/MM/yyyy");
                    }

                    var Premium_Amt = Convert.ToString(((Microsoft.Office.Interop.Excel.Range)wks.Cells[i, 4]).Value);
                    var Revenue_Amt = Convert.ToString(((Microsoft.Office.Interop.Excel.Range)wks.Cells[i, 25]).Value);
                    var Revenue_Pcnt = Convert.ToString(((Microsoft.Office.Interop.Excel.Range)wks.Cells[i, 14]).Value);
                    //Terrorism = Convert.ToString(((Microsoft.Office.Interop.Excel.Range)wks.Cells[i, 13]).Value);

                    var Policy_Type = ((Microsoft.Office.Interop.Excel.Range)wks.Cells[i, 19]).Value.Replace("\n", "").Replace(",", "").TrimStart();
                    var Presult = PolicyNo.Substring(PolicyNo.LastIndexOf('/') + 1);
                    if (Presult.Contains("0"))
                    {
                        Policy_Endorsement = "Policy";
                    }
                    else
                    {
                        Policy_Endorsement = "Endorsement";
                    }

                    //if (InsuredName.Contains("LIMITED") || InsuredName.Contains("LTD") || InsuredName.Contains(".COM"))
                    //{
                    InsuredType = "Corporate";
                    //}
                    //else
                    //{
                    //    InsuredType = "Retail";
                    //}
                    if (Premium_Amt == "" || Premium_Amt == " " || Terrorism == null)
                    {
                        Premium_Amt = 0;
                    }
                    if (Revenue_Amt == "" || Revenue_Amt == " " || Terrorism == null)
                    {
                        Revenue_Amt = 0;
                    }
                    if (Terrorism == "" || Terrorism == " " || Terrorism == null)
                    {
                        Terrorism = "0";
                    }
                    if (Revenue_Pcnt == "" || Revenue_Pcnt == " " || Terrorism == null)
                    {
                        Revenue_Pcnt = "0";
                    }
                    sql.ExecuteSQLNonQuery(strconn,"SP_StarHealthTransactions",
                               new SqlParameter { ParameterName = "@Imode", Value = 1 },
                               new SqlParameter { ParameterName = "@RDate", Value = RDate },
                               new SqlParameter { ParameterName = "@Rmonth", Value = Rmonth },
                               new SqlParameter { ParameterName = "@ClientName", Value = InsuredName },
                               new SqlParameter { ParameterName = "@Itype", Value = InsuredType },
                               new SqlParameter { ParameterName = "@Client_N_E", Value = Client_N_E },
                               new SqlParameter { ParameterName = "@Policy_No", Value = PolicyNo },
                               new SqlParameter { ParameterName = "@Endo_Effective_Date", Value = Endo_Effective_Date },
                               new SqlParameter { ParameterName = "@Effective_Date", Value = Effective_Date },
                               new SqlParameter { ParameterName = "@END_Date", Value = END_Date },
                               new SqlParameter { ParameterName = "@Policy_Type", Value = Policy_Type },
                               new SqlParameter { ParameterName = "@Premium_Amt", Value = Premium_Amt },
                               new SqlParameter { ParameterName = "@TranID", Value = TranID },
                               new SqlParameter { ParameterName = "@Revenue_Amt", Value = Revenue_Amt },
                               new SqlParameter { ParameterName = "@Terrorism", Value = Terrorism },
                               new SqlParameter { ParameterName = "@Revenue_Pcnt", Value = Revenue_Pcnt },
                               new SqlParameter { ParameterName = "@Insurance", Value = Insurance },
                               new SqlParameter { ParameterName = "@Salesby", Value = Salesby },
                               new SqlParameter { ParameterName = "@Serviceby", Value = Serviceby },
                               new SqlParameter { ParameterName = "@location", Value = location },
                               new SqlParameter { ParameterName = "@Support", Value = Support },
                               new SqlParameter { ParameterName = "@Policy_Endorsement", Value = Policy_Endorsement },
                               new SqlParameter { ParameterName = "@RFormat", Value = "F1" },
                               new SqlParameter { ParameterName = "@InvNo", Value = "STAR" },
                               new SqlParameter { ParameterName = "@ReportId", Value = "STA1" },
                               new SqlParameter { ParameterName = "@DocName", Value = "Star Health & Allied Insurance Co. Ltd." }
                               );
                }
            }
        }
        public static void InsertRetailTransaction(Microsoft.Office.Interop.Excel.Workbook WB, string TranID, string RDate, string Insurance, string Salesby, string Serviceby, string location, string Support, string Rmonth, string strconn)
        {
            //Star health Retail - F2
            SQLProcs sql = new SQLProcs();
            Microsoft.Office.Interop.Excel.Worksheet wks = (Microsoft.Office.Interop.Excel.Worksheet)WB.Worksheets[1];
            Microsoft.Office.Interop.Excel.Range lastCell = wks.Cells.SpecialCells(Microsoft.Office.Interop.Excel.XlCellType.xlCellTypeLastCell, Type.Missing);
            int lastrow = lastCell.Row; 
            for (int i = 2; i <= lastrow; i++)
            {
                var Terrorism = ""; var Policy_Endorsement = ""; var InsuredType = "";
                string InsuredName = ((Microsoft.Office.Interop.Excel.Range)wks.Cells[i, 15]).Value;
                if (InsuredName != null && InsuredName != "" && InsuredName != " ")
                {
                    InsuredName = InsuredName.Replace("\n", "").TrimStart();
                    string Client_N_E = ((Microsoft.Office.Interop.Excel.Range)wks.Cells[i, 11]).Value.Replace("\n", "").TrimStart();
                    if (Client_N_E.ToLower() == "fresh")
                    {
                        Client_N_E = "New";
                    }
                    string PolicyNo = ((Microsoft.Office.Interop.Excel.Range)wks.Cells[i, 10]).Value.Replace("\n", "").TrimStart();
                    var Endo_Effective_Date = ((Microsoft.Office.Interop.Excel.Range)wks.Cells[i, 3]).Value;
                    int ENdolen = Convert.ToString(Endo_Effective_Date).Length;
                    if (ENdolen > 11)
                    {
                        Endo_Effective_Date = Endo_Effective_Date.ToString("dd/MM/yyyy");
                    }

                    var Premium_Amt = Convert.ToString(((Microsoft.Office.Interop.Excel.Range)wks.Cells[i, 23]).Value);
                    var Revenue_Amt = Convert.ToString(((Microsoft.Office.Interop.Excel.Range)wks.Cells[i, 25]).Value);
                    var Revenue_Pcnt = Convert.ToString(((Microsoft.Office.Interop.Excel.Range)wks.Cells[i, 24]).Value).Replace("%", "").Replace("\n", "").Replace("0.", "").TrimStart();
                    //Terrorism = Convert.ToString(((Microsoft.Office.Interop.Excel.Range)wks.Cells[i, 13]).Value);

                    var Policy_Type = ((Microsoft.Office.Interop.Excel.Range)wks.Cells[i, 6]).Value.Replace("\n", "").Replace(",", "").TrimStart();
                    //InsuredType = ((Microsoft.Office.Interop.Excel.Range)wks.Cells[i, 5]).Value.Replace("\n", "").Replace(",", "").TrimStart();
                    InsuredType = "Retail";
                    if (PolicyNo.StartsWith("p"))
                    {
                        var Presult = PolicyNo.Substring(PolicyNo.LastIndexOf('/') + 1);

                        if (Convert.ToDouble(Presult) > 0)
                        {
                            Policy_Endorsement = "Policy";
                        }
                        else
                        {
                            Policy_Endorsement = "Endorsement";
                        }
                    }
                    else
                    {
                        Policy_Endorsement = "Policy";
                    }
                    if (Premium_Amt == "" || Premium_Amt == " " || Terrorism == null)
                    {
                        Premium_Amt = 0;
                    }
                    if (Revenue_Amt == "" || Revenue_Amt == " " || Terrorism == null)
                    {
                        Revenue_Amt = 0;
                    }
                    if (Terrorism == "" || Terrorism == " " || Terrorism == null)
                    {
                        Terrorism = "0";
                    }
                    if (Revenue_Pcnt == "" || Revenue_Pcnt == " " || Terrorism == null)
                    {
                        Revenue_Pcnt = "0";
                    }
                    sql.ExecuteSQLNonQuery(strconn,"SP_StarHealthTransactions",
                               new SqlParameter { ParameterName = "@Imode", Value = 1 },
                               new SqlParameter { ParameterName = "@RDate", Value = RDate },
                               new SqlParameter { ParameterName = "@Rmonth", Value = Rmonth },
                               new SqlParameter { ParameterName = "@ClientName", Value = InsuredName },
                               new SqlParameter { ParameterName = "@Itype", Value = InsuredType },
                               new SqlParameter { ParameterName = "@Client_N_E", Value = Client_N_E },
                               new SqlParameter { ParameterName = "@Policy_No", Value = PolicyNo },
                               new SqlParameter { ParameterName = "@Endo_Effective_Date", Value = Endo_Effective_Date },
                               //new SqlParameter { ParameterName = "@Effective_Date", Value = Effective_Date },
                               //new SqlParameter { ParameterName = "@END_Date", Value = END_Date },
                               new SqlParameter { ParameterName = "@Policy_Type", Value = Policy_Type },
                               new SqlParameter { ParameterName = "@Premium_Amt", Value = Premium_Amt },
                               new SqlParameter { ParameterName = "@TranID", Value = TranID },
                               new SqlParameter { ParameterName = "@Revenue_Amt", Value = Revenue_Amt },
                               new SqlParameter { ParameterName = "@Terrorism", Value = Terrorism },
                               new SqlParameter { ParameterName = "@Revenue_Pcnt", Value = Revenue_Pcnt },
                               new SqlParameter { ParameterName = "@Insurance", Value = Insurance },
                               new SqlParameter { ParameterName = "@Salesby", Value = Salesby },
                               new SqlParameter { ParameterName = "@Serviceby", Value = Serviceby },
                               new SqlParameter { ParameterName = "@location", Value = location },
                               new SqlParameter { ParameterName = "@Support", Value = Support },
                               new SqlParameter { ParameterName = "@Policy_Endorsement", Value = Policy_Endorsement },
                               new SqlParameter { ParameterName = "@RFormat", Value = "F1" },
                               new SqlParameter { ParameterName = "@InvNo", Value = "STAR" },
                               new SqlParameter { ParameterName = "@ReportId", Value = "STA1" },
                               new SqlParameter { ParameterName = "@DocName", Value = "Star Health & Allied Insurance Co. Ltd." }
                               );
                }
            }
        }
    }
}
