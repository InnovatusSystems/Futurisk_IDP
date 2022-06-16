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
    public class AdityaInsurance
    {
        public static void InsertTransaction(Microsoft.Office.Interop.Excel.Workbook WB, string TranID, string RDate, string Insurance, string Salesby, string Serviceby, string location, string Support, string Rmonth, string strconn)
        {

            SQLProcs sql = new SQLProcs();
            Microsoft.Office.Interop.Excel.Worksheet wks = (Microsoft.Office.Interop.Excel.Worksheet)WB.Worksheets[1];
            Microsoft.Office.Interop.Excel.Range lastCell = wks.Cells.SpecialCells(Microsoft.Office.Interop.Excel.XlCellType.xlCellTypeLastCell, Type.Missing);
            int lastrow = lastCell.Row; 
            for (int i = 2; i <= lastrow; i++)
            {
                var Terrorism = ""; var Policy_Endorsement = ""; var InsuredType = ""; var location1 = "";
                string InsuredName = ((Microsoft.Office.Interop.Excel.Range)wks.Cells[i, 9]).Value.Replace("\n", "").TrimStart();
                if (InsuredName == null || InsuredName == "")
                {
                    InsuredName = ((Microsoft.Office.Interop.Excel.Range)wks.Cells[i, 9]).Value.Replace("\n", "").TrimStart();
                }
                string PolicyNo = ((Microsoft.Office.Interop.Excel.Range)wks.Cells[i, 6]).Value.Replace("\n", "").TrimStart();
                if (PolicyNo == null || PolicyNo == "")
                {
                    PolicyNo = ((Microsoft.Office.Interop.Excel.Range)wks.Cells[i, 5]).Value.Replace("\n", "").TrimStart();
                }
                if (location1 == null || location1 == "")
                {
                    location1 = ((Microsoft.Office.Interop.Excel.Range)wks.Cells[i, 15]).Value.Replace("\n", "").TrimStart();
                }

                var bustype = ((Microsoft.Office.Interop.Excel.Range)wks.Cells[i, 4]).Value;


                //var Endo_Effective_Date = Convert.ToString(((Microsoft.Office.Interop.Excel.Range)wks.Cells[i, 4]).Value);//.Replace("\n", "").TrimStart();
                //string END_Date = Convert.ToString(((Microsoft.Office.Interop.Excel.Range)wks.Cells[i, 6]).Value);//.Replace("\n", "").TrimStart();
                //var Endo_Effective_Date = ((Microsoft.Office.Interop.Excel.Range)wks.Cells[i, 27]).Value;
                var PE = ((Microsoft.Office.Interop.Excel.Range)wks.Cells[i, 6]).Value;
                //var END_Date = ((Microsoft.Office.Interop.Excel.Range)wks.Cells[i, 28]).Value;
                //Endo_Effective_Date = Endo_Effective_Date.ToString("dd/MM/yyyy");
                //END_Date = END_Date.ToString("dd/MM/yyyy");
                var Premium_Amt = Convert.ToString(((Microsoft.Office.Interop.Excel.Range)wks.Cells[i, 12]).Value);
                var Revenue_Amt = Convert.ToString(((Microsoft.Office.Interop.Excel.Range)wks.Cells[i, 14]).Value);
                var Revenue_Pct = Convert.ToString(((Microsoft.Office.Interop.Excel.Range)wks.Cells[i, 13]).Value);
                Revenue_Pct = Revenue_Pct.Replace("%", "").TrimStart();
                var TP_Amt = "";// Convert.ToString(((Microsoft.Office.Interop.Excel.Range)wks.Cells[i, 32]).Value);
                                //Terrorism = Convert.ToString(((Microsoft.Office.Interop.Excel.Range)wks.Cells[i, 42]).Value);
                                //var RewardOD_Pct = Convert.ToString(((Microsoft.Office.Interop.Excel.Range)wks.Cells[i, 46]).Value);
                                //var RewardTP_Pct = Convert.ToString(((Microsoft.Office.Interop.Excel.Range)wks.Cells[i, 46]).Value);
                                // var IRDARewardAmt = Convert.ToString(((Microsoft.Office.Interop.Excel.Range)wks.Cells[i, 48]).Value);
                var Policy_Type = ((Microsoft.Office.Interop.Excel.Range)wks.Cells[i, 3]).Value.Replace("\n", "").Replace(",", "").TrimStart();
                if (bustype.Contains("Group"))
                {
                    InsuredType = "corporate";
                }
                else
                {
                    InsuredType = "Retail";
                }
                if (Premium_Amt == "" || Premium_Amt == " ")
                {
                    Premium_Amt = 0;
                }
                if (Revenue_Amt == "" || Revenue_Amt == " ")
                {
                    Revenue_Amt = 0;
                }
                if (Terrorism == "" || Terrorism == " " || Terrorism == null)
                {
                    Terrorism = "0";
                }
                if (PE == "" || PE == null)
                {
                    Policy_Endorsement = "Policy";
                }
                else
                {
                    Policy_Endorsement = "Endoresement";
                }
                if (TP_Amt == "" || TP_Amt == " " || Terrorism == null)
                {
                    TP_Amt = "0";
                }
                if (Revenue_Pct == "" || Revenue_Pct == " ")
                {
                    Revenue_Pct = "0";
                }
                //if (RewardOD_Pct == "" || RewardOD_Pct == " " || Terrorism == null)
                //{
                //    RewardOD_Pct = "0";
                //}
                //if (RewardTP_Pct == "" || RewardTP_Pct == " " || Terrorism == null)
                //{
                //    RewardTP_Pct = "0";
                //}
                //if (IRDARewardAmt == "" || IRDARewardAmt == " " || Terrorism == null)
                //{
                //    IRDARewardAmt = "0";
                //}
                sql.ExecuteSQLNonQuery(strconn, "SP_AdityaBirlaTransactions",
                           new SqlParameter { ParameterName = "@Imode", Value = 1 },
                           new SqlParameter { ParameterName = "@RDate", Value = RDate },
                           new SqlParameter { ParameterName = "@Rmonth", Value = Rmonth },
                           new SqlParameter { ParameterName = "@ClientName", Value = InsuredName },
                           new SqlParameter { ParameterName = "@Itype", Value = InsuredType },
                           new SqlParameter { ParameterName = "@Policy_No", Value = PolicyNo },
                           new SqlParameter { ParameterName = "@Endo_Effective_Date", Value = null },
                           new SqlParameter { ParameterName = "@END_Date", Value = null },
                           new SqlParameter { ParameterName = "@Policy_Type", Value = Policy_Type },
                           new SqlParameter { ParameterName = "@Premium_Amt", Value = Premium_Amt },
                           new SqlParameter { ParameterName = "@TranID", Value = TranID },
                           new SqlParameter { ParameterName = "@Revenue_Amt", Value = Revenue_Amt },
                           new SqlParameter { ParameterName = "@TP_Amt", Value = TP_Amt },
                           new SqlParameter { ParameterName = "@Terrorism", Value = Terrorism },
                           new SqlParameter { ParameterName = "@Effective_Date", Value = null },
                           new SqlParameter { ParameterName = "@Revenue_Pct", Value = Revenue_Pct },
                           // new SqlParameter { ParameterName = "@RewardOD_Pct", Value = RewardOD_Pct },
                           //new SqlParameter { ParameterName = "@RewardTP_Pct", Value = RewardTP_Pct },
                           // new SqlParameter { ParameterName = "@IRDARewardAmt", Value = IRDARewardAmt },
                           new SqlParameter { ParameterName = "@Insurance", Value = Insurance },
                           new SqlParameter { ParameterName = "@Salesby", Value = Salesby },
                           new SqlParameter { ParameterName = "@Serviceby", Value = Serviceby },
                           new SqlParameter { ParameterName = "@location", Value = location1 },
                           new SqlParameter { ParameterName = "@Support", Value = Support },
                           new SqlParameter { ParameterName = "@Policy_Endorsement", Value = Policy_Endorsement },
                           new SqlParameter { ParameterName = "@RFormat", Value = "F1" },
                           new SqlParameter { ParameterName = "@InvNo", Value = "ABHI" },
                           new SqlParameter { ParameterName = "@ReportId", Value = "ABHX" },
                           new SqlParameter { ParameterName = "@DocName", Value = "Aditya Birla Health Insurance Co. Ltd." }
                           );


            }

        }
    }
}
