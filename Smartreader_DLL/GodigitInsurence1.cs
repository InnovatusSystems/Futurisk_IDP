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
    public class GodigitInsurence1
    {
        public static void InsertTransaction(Microsoft.Office.Interop.Excel.Workbook WB, string TranID, string RDate, string Insurance, string Salesby, string Serviceby, string location, string Support, string Rmonth, string strconn)
        {
            SQLProcs sql = new SQLProcs();
            Microsoft.Office.Interop.Excel.Worksheet wks = (Microsoft.Office.Interop.Excel.Worksheet)WB.Worksheets[1];
            Microsoft.Office.Interop.Excel.Range lastCell = wks.Cells.SpecialCells(Microsoft.Office.Interop.Excel.XlCellType.xlCellTypeLastCell, Type.Missing);
            int lastrow = lastCell.Row; var Terrorism = ""; var Policy_Endorsement = ""; var InsuredType = "";
            for (int i = 2; i <= lastrow; i++)
            {
                string InsuredName = ((Microsoft.Office.Interop.Excel.Range)wks.Cells[i, 30]).Value.Replace("\n", "").TrimStart();
                if (InsuredName == null || InsuredName == "")
                {
                    InsuredName = ((Microsoft.Office.Interop.Excel.Range)wks.Cells[i, 37]).Value.Replace("\n", "").TrimStart();
                }
                string PolicyNo = ((Microsoft.Office.Interop.Excel.Range)wks.Cells[i, 38]).Value.Replace("\n", "").TrimStart();
                if(PolicyNo == null || PolicyNo == "")
                {
                    PolicyNo = ((Microsoft.Office.Interop.Excel.Range)wks.Cells[i, 1]).Value.Replace("\n", "").TrimStart();
                }
                string Endo_Effective_Date = ((Microsoft.Office.Interop.Excel.Range)wks.Cells[i, 4]).Value.Replace("\n", "").TrimStart();
                var END_Date = ((Microsoft.Office.Interop.Excel.Range)wks.Cells[i, 6]).Value.Replace("\n", "").TrimStart();
                var Premium_Amt = ((Microsoft.Office.Interop.Excel.Range)wks.Cells[i, 31]).Value.Replace("\n", "").Replace(",", "").TrimStart();
                var Revenue_Amt = ((Microsoft.Office.Interop.Excel.Range)wks.Cells[i, 47]).Value.Replace("\n", "").Replace(",", "").TrimStart();
                var Revenue_Pct = ((Microsoft.Office.Interop.Excel.Range)wks.Cells[i, 43]).Value.Replace("\n", "").Replace(",", "").TrimStart();
                var TP_Amt = ((Microsoft.Office.Interop.Excel.Range)wks.Cells[i, 32]).Value.Replace("\n", "").Replace(",", "").TrimStart();
                Terrorism = ((Microsoft.Office.Interop.Excel.Range)wks.Cells[i, 41]).Value.Replace("\n", "").Replace(",", "").TrimStart();
                var RewardOD_Pct = ((Microsoft.Office.Interop.Excel.Range)wks.Cells[i, 41]).Value.Replace("\n", "").Replace(",", "").TrimStart();
                var RewardTP_Pct = ((Microsoft.Office.Interop.Excel.Range)wks.Cells[i, 46]).Value.Replace("\n", "").Replace(",", "").TrimStart();
                var IRDARewardAmt = ((Microsoft.Office.Interop.Excel.Range)wks.Cells[i, 48]).Value.Replace("\n", "").Replace(",", "").TrimStart();
                var Policy_Type = ((Microsoft.Office.Interop.Excel.Range)wks.Cells[i, 15]).Value.Replace("\n", "").Replace(",", "").TrimStart();
                if (InsuredName.Contains("LIMITED"))
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
                if (Terrorism == "" || Terrorism == " ")
                {
                    Terrorism = "0";
                }
                if (TP_Amt == "" || TP_Amt == " ")
                {
                    TP_Amt = "0";
                }
                if (Revenue_Pct == "" || Revenue_Pct == " ")
                {
                    Revenue_Pct = "0";
                }
                if (RewardOD_Pct == "" || RewardOD_Pct == " ")
                {
                    RewardOD_Pct = "0";
                }
                if (RewardTP_Pct == "" || RewardTP_Pct == " ")
                {
                    RewardTP_Pct = "0";
                }
                if (IRDARewardAmt == "" || IRDARewardAmt == " ")
                {
                    IRDARewardAmt = "0";
                }
                sql.ExecuteSQLNonQuery(strconn, "SP_Insert_Transactions",
                           new SqlParameter { ParameterName = "@Imode", Value = 1 },
                           new SqlParameter { ParameterName = "@RDate", Value = RDate },
                           new SqlParameter { ParameterName = "@Rmonth", Value = Rmonth },
                           new SqlParameter { ParameterName = "@ClientName", Value = InsuredName },
                           new SqlParameter { ParameterName = "@InsuredType", Value = InsuredType },
                           new SqlParameter { ParameterName = "@Policy_No", Value = PolicyNo },
                           new SqlParameter { ParameterName = "@Endo_Effective_Date", Value = Endo_Effective_Date },
                           new SqlParameter { ParameterName = "@END_Date", Value = END_Date },
                           new SqlParameter { ParameterName = "@Policy_Type", Value = Policy_Type },
                           new SqlParameter { ParameterName = "@Premium_Amt", Value = Premium_Amt },
                           new SqlParameter { ParameterName = "@TranID", Value = TranID },
                           new SqlParameter { ParameterName = "@Revenue_Amt", Value = Revenue_Amt },
                           new SqlParameter { ParameterName = "@Terrorism", Value = Terrorism },
                           new SqlParameter { ParameterName = "@TP_Amt", Value = TP_Amt },
                           new SqlParameter { ParameterName = "@Revenue_Pct", Value = Revenue_Pct },
                           new SqlParameter { ParameterName = "@RewardOD_Pct", Value = RewardOD_Pct },
                           new SqlParameter { ParameterName = "@RewardTP_Pct", Value = RewardTP_Pct },
                           new SqlParameter { ParameterName = "@IRDARewardAmt", Value = IRDARewardAmt },
                           new SqlParameter { ParameterName = "@Insurance", Value = Insurance },
                           new SqlParameter { ParameterName = "@Salesby", Value = Salesby },
                           new SqlParameter { ParameterName = "@Serviceby", Value = Serviceby },
                           new SqlParameter { ParameterName = "@location", Value = location },
                           new SqlParameter { ParameterName = "@Support", Value = Support },
                           new SqlParameter { ParameterName = "@Policy_Endorsement", Value = Policy_Endorsement },
                           new SqlParameter { ParameterName = "@RFormat", Value = "F1" },
                           new SqlParameter { ParameterName = "@InvNo", Value = "GGIC" },
                           new SqlParameter { ParameterName = "@ReportId", Value = "GGI1" },
                           new SqlParameter { ParameterName = "@DocName", Value = "Godigit General Insurance Co. Ltd." }
                           );
            }
        }
    }
}
