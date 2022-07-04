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
    public  class ExportCreditTransaction
    {
        public static void InsertTransaction(Microsoft.Office.Interop.Excel.Workbook WB, string TranID, string RDate, string Insurance, string Salesby, string Serviceby, string location, string Support, string Rmonth, string strconn)
        {
            //Export Credit Guarantee Corporation of India Ltd.
            SQLProcs sql = new SQLProcs();
            Microsoft.Office.Interop.Excel.Worksheet wks = (Microsoft.Office.Interop.Excel.Worksheet)WB.Worksheets[1];
            Microsoft.Office.Interop.Excel.Range lastCell = wks.Cells.SpecialCells(Microsoft.Office.Interop.Excel.XlCellType.xlCellTypeLastCell, Type.Missing);
            int lastrow = lastCell.Row; string InsuredName = ""; string PolicyNo = "";
            for (int i = 13; i <= lastrow; i++)
            {
                var Terrorism = ""; var Policy_Endorsement = ""; var InsuredType = ""; var Policy_Type = "";
                var New_Renewal = ""; var Revenue_Amt = "";
                string IN = ((Microsoft.Office.Interop.Excel.Range)wks.Cells[i, 2]).Value;
                if (IN != null && IN != "" && IN != " ")
                {
                    InsuredName = IN.Replace("\n", "").TrimStart();
                }
                string Pno = Convert.ToString(((Microsoft.Office.Interop.Excel.Range)wks.Cells[i, 3]).Value);
                if (Pno != null && Pno != "" && Pno != " ")
                {
                    Pno = Pno.Replace("\n", "").TrimStart();
                    PolicyNo = Regex.Match(Pno, @"\d+").Value;
                }
                var Endo_Effective_Date = ((Microsoft.Office.Interop.Excel.Range)wks.Cells[i, 4]).Value;
                if (Endo_Effective_Date != null && Convert.ToString(Endo_Effective_Date) != "" && Convert.ToString(Endo_Effective_Date) != " " && Convert.ToString(Endo_Effective_Date) != "कुल")
                {
                    InsuredType = "Corporate";

                    //var Endorsementno = Convert.ToString(((Microsoft.Office.Interop.Excel.Range)wks.Cells[i, 15]).Value);
                    //var offLocation = Convert.ToString(((Microsoft.Office.Interop.Excel.Range)wks.Cells[i, 12]).Value).Replace("\n", "").TrimStart();


                    //var Effective_Date = ((Microsoft.Office.Interop.Excel.Range)wks.Cells[i, 9]).Value;
                    //var END_Date = ((Microsoft.Office.Interop.Excel.Range)wks.Cells[i, 10]).Value;

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
                    //int ENDlen = Convert.ToString(END_Date).Length;
                    //if (ENDlen > 11)
                    //{
                    //    END_Date = END_Date.ToString("dd/MM/yyyy");
                    //}
                    //var Client_N_E = ((Microsoft.Office.Interop.Excel.Range)wks.Cells[i, 40]).Value.Replace("\n", "").TrimStart();

                    var Premium_Amt = Convert.ToString(((Microsoft.Office.Interop.Excel.Range)wks.Cells[i, 5]).Value).Replace(",", "").Replace("(", "").Replace(")", "").TrimStart();
                    //Terrorism = Convert.ToString(((Microsoft.Office.Interop.Excel.Range)wks.Cells[i, 23]).Value).Replace(",", "").Replace("(", "").Replace(")", "").TrimStart();
                    Revenue_Amt = Convert.ToString(((Microsoft.Office.Interop.Excel.Range)wks.Cells[i, 7]).Value).Replace(",", "").Replace("(", "").Replace(")", "").Replace("-", "").TrimStart();
                    var Revenue_Pcnt = Convert.ToString(((Microsoft.Office.Interop.Excel.Range)wks.Cells[i, 6]).Text);
                    Revenue_Pcnt = Revenue_Pcnt.Replace("\n", "").Replace("%", "").Replace(",", "").TrimStart();
                    //PolicyNo = Regex.Replace(PolicyNo, "[^a-zA-Z0-9_.]+", "", RegexOptions.Compiled);

                    //string endno = Convert.ToString(((Microsoft.Office.Interop.Excel.Range)wks.Cells[i, 5]).Value).Replace("\n", "").TrimStart();
                    //if (endno != null && endno != "" && endno != " " && endno != ":")
                    //{
                    //    Policy_Endorsement = "Endorsement";
                    //}
                    //else
                    //{
                    //    Policy_Endorsement = "Policy";
                    //}
                    //if (Client_N_E == "New Business")
                    //{
                    //    Client_N_E = "New Client";
                    //    if (Policy_Endorsement != "Endorsement" && InsuredType != "Retail")
                    //    {
                    //        New_Renewal = "New Policy";
                    //    }
                    //}
                    //else
                    //{
                    //    Client_N_E = "Existing Client";
                    //    if (Policy_Endorsement != "Endorsement" && InsuredType != "Retail")
                    //    {
                    //        New_Renewal = "Renewal Policy";
                    //    }
                    //}

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
                    sql.ExecuteSQLNonQuery(strconn,"SP_ExportCreditTransaction",
                               new SqlParameter { ParameterName = "@Imode", Value = 1 },
                               new SqlParameter { ParameterName = "@RDate", Value = RDate },
                               new SqlParameter { ParameterName = "@Rmonth", Value = Rmonth },
                               new SqlParameter { ParameterName = "@ClientName", Value = InsuredName },
                               new SqlParameter { ParameterName = "@Itype", Value = InsuredType },
                               new SqlParameter { ParameterName = "@Policy_No", Value = PolicyNo },
                               new SqlParameter { ParameterName = "@Revenue_Pct", Value = Revenue_Pcnt },
                               //new SqlParameter { ParameterName = "@Policy_Type", Value = Policy_Type },
                               new SqlParameter { ParameterName = "@Premium_Amt", Value = Premium_Amt },
                               new SqlParameter { ParameterName = "@Endo_Effective_Date", Value = Endo_Effective_Date },
                               //new SqlParameter { ParameterName = "@Effective_Date", Value = Effective_Date },
                               //new SqlParameter { ParameterName = "@END_Date", Value = END_Date },
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
                               new SqlParameter { ParameterName = "@InvNo", Value = "ECGX" },
                               new SqlParameter { ParameterName = "@ReportId", Value = "ECGX" },
                               new SqlParameter { ParameterName = "@DocName", Value = "Export Credit Guarantee Corporation of India Ltd." }
                               );
                }
            }
        }
    }
}
