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
    public  class NivaInsurance
    {
        public static void InsertTransaction(Microsoft.Office.Interop.Excel.Workbook WB, string TranID, string RDate, string Insurance, string Salesby, string Serviceby, string location, string Support, string Rmonth, string strconn)
        {
            //Niva Bupa Health Insurance Co. Ltd. 
            SQLProcs sql = new SQLProcs();
            Microsoft.Office.Interop.Excel.Worksheet wks = (Microsoft.Office.Interop.Excel.Worksheet)WB.Worksheets[1];
            Microsoft.Office.Interop.Excel.Range lastCell = wks.Cells.SpecialCells(Microsoft.Office.Interop.Excel.XlCellType.xlCellTypeLastCell, Type.Missing);
            int lastrow = lastCell.Row;
            for (int i = 2; i <= lastrow; i++)
            {
                var Terrorism = ""; var Policy_Endorsement = ""; var InsuredType = ""; var Policy_Type = ""; var New_Renewal = ""; var Revenue_Amt = ""; var Premium_Amt = "";
                string InsuredName = ((Microsoft.Office.Interop.Excel.Range)wks.Cells[i, 4]).Value;
                if (InsuredName != null && InsuredName != "" && InsuredName != " ")
                {
                    InsuredName = InsuredName.Replace("\n", "").TrimStart();
                    //InsuredType = ((Microsoft.Office.Interop.Excel.Range)wks.Cells[i, 8]).Value;
                    string PolicyNo = Convert.ToString(((Microsoft.Office.Interop.Excel.Range)wks.Cells[i, 1]).Value).Replace("\n", "").TrimStart();
                    //var Client_N_E = Convert.ToString(((Microsoft.Office.Interop.Excel.Range)wks.Cells[i, 8]).Value);
                    //var Endno = Convert.ToString(((Microsoft.Office.Interop.Excel.Range)wks.Cells[i, 7]).Value);
                    //if (Endno == null || Endno == "" || Endno == " ")
                    //{
                    //    Policy_Endorsement = "Policy";
                    //}
                    //else
                    //{
                    //    PolicyNo = PolicyNo + "/" + Endno;
                    //    Policy_Endorsement = "Endorsement";
                    //}

                    //var Endo_Effective_Date = ((Microsoft.Office.Interop.Excel.Range)wks.Cells[i, 18]).Value;
                    //var Effective_Date = ((Microsoft.Office.Interop.Excel.Range)wks.Cells[i, 17]).Value;
                    //var END_Date = ((Microsoft.Office.Interop.Excel.Range)wks.Cells[i, 19]).Value;
                    //if (Endo_Effective_Date != null && Endo_Effective_Date != "" && Endo_Effective_Date != " ")
                    //{
                    //    int ENdolen = Convert.ToString(Endo_Effective_Date).Length;
                    //    if (ENdolen > 11)
                    //    {
                    //        Endo_Effective_Date = Endo_Effective_Date.ToString("dd/MM/yyyy");
                    //    }
                    //}
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
                    Premium_Amt = Convert.ToString(((Microsoft.Office.Interop.Excel.Range)wks.Cells[i, 6]).Value);
                    if (Premium_Amt != null)
                    {
                        Premium_Amt = Premium_Amt.Replace(",", "").Replace("(", "").Replace(")", "").TrimStart();
                    }
                    //Premium_Amt = Convert.ToString(Convert.ToDecimal(PA1) + Convert.ToDecimal(PA2));
                    //Terrorism = Convert.ToString(((Microsoft.Office.Interop.Excel.Range)wks.Cells[i, 45]).Value).Replace(",", "").Replace("(", "").Replace(")", "").TrimStart();
                    Revenue_Amt = Convert.ToString(((Microsoft.Office.Interop.Excel.Range)wks.Cells[i, 7]).Value);
                    if (Revenue_Amt != null)
                    {
                        Revenue_Amt = Revenue_Amt.Replace(",", "").Replace("(", "").Replace(")", "").Replace("-", "").TrimStart();
                    }
                    var Revenue_Pcnt = Convert.ToString(((Microsoft.Office.Interop.Excel.Range)wks.Cells[i, 10]).Text);
                    if (Revenue_Pcnt != null)
                    {
                        Revenue_Pcnt = Revenue_Pcnt.Replace("\n", "").Replace("%", "").Replace(",", "").TrimStart();
                    }
                    Policy_Type = Convert.ToString(((Microsoft.Office.Interop.Excel.Range)wks.Cells[i, 19]).Value);
                    var offlocation = Convert.ToString(((Microsoft.Office.Interop.Excel.Range)wks.Cells[i, 17]).Value);
                    if (InsuredName.ToUpper().Contains("LIMITED") || InsuredName.ToUpper().Contains("LTD"))
                    {
                        InsuredType = "Corporate";
                    }
                    else
                    {
                        InsuredType = "Retail";
                    }

                    if (Premium_Amt == "" || Premium_Amt == " " || Premium_Amt == null)
                    {
                        Premium_Amt = "0";
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
                    sql.ExecuteSQLNonQuery(strconn,"SP_NivaExcelTransaction",
                               new SqlParameter { ParameterName = "@Imode", Value = 1 },
                               new SqlParameter { ParameterName = "@RDate", Value = RDate },
                               new SqlParameter { ParameterName = "@Rmonth", Value = Rmonth },
                               new SqlParameter { ParameterName = "@ClientName", Value = InsuredName },
                               new SqlParameter { ParameterName = "@Itype", Value = InsuredType },
                               new SqlParameter { ParameterName = "@Policy_No", Value = PolicyNo },
                               //new SqlParameter { ParameterName = "@Endo_Effective_Date", Value = Endo_Effective_Date },
                               //new SqlParameter { ParameterName = "@Effective_Date", Value = Effective_Date },
                               //new SqlParameter { ParameterName = "@END_Date", Value = END_Date },
                               new SqlParameter { ParameterName = "@Revenue_Pct", Value = Revenue_Pcnt },
                               new SqlParameter { ParameterName = "@Policy_Type", Value = Policy_Type },
                               new SqlParameter { ParameterName = "@Premium_Amt", Value = Premium_Amt },
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
                               new SqlParameter { ParameterName = "@InvNo", Value = "NBHX" },
                               new SqlParameter { ParameterName = "@ReportId", Value = "NBHX" },
                               new SqlParameter { ParameterName = "@DocName", Value = "Niva Bupa Health Insurance Co.Ltd." }
                               );
                }
            }
        }
    }
}
