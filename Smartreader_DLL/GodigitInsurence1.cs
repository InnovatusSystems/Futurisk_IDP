using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using NPOI.SS.UserModel;
using System.Text;
using System.Threading.Tasks;
using System.Text.RegularExpressions;
using Syncfusion.XlsIO;
using System.IO;
using System.Globalization;

namespace Smartreader_DLL
{
    public class GodigitInsurence1
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

        public static void ExceltoDatatable(string TranID, string RDate, string Insurance, string Salesby, string Serviceby, string location, string Support, string Rmonth, string Filepath, string strconn)
        {
            SQLProcs sql = new SQLProcs();
            using (ExcelEngine excelEngine = new ExcelEngine())
            {
                IApplication application = excelEngine.Excel;
                application.DefaultVersion = Syncfusion.XlsIO.ExcelVersion.Xlsx;
                FileStream inputStream = new FileStream(Filepath, FileMode.Open, FileAccess.Read);
                Syncfusion.XlsIO.IWorkbook workbook = application.Workbooks.Open(inputStream);
                IWorksheet worksheet = workbook.Worksheets[0];

                //Read data from the worksheet and export to the DataTable.
                DataTable DT = worksheet.ExportDataTable(worksheet.UsedRange, ExcelExportDataTableOptions.ColumnNames);
                ///var Terrorism = ""; var InsuredType = ""; var Policy_Endorsement = "";

                DataTable RDT = copyDatatable(DT);

                //DataSet ds = new DataSet();

                //ds = 
                sql.ExecuteSQLNonQuery(strconn,"SP_GoDigitTransactions",
                                   new SqlParameter { ParameterName = "@Imode", Value = 9 },
                                   new SqlParameter { ParameterName = "@RDate", Value = RDate },
                                   new SqlParameter { ParameterName = "@Rmonth", Value = Rmonth },
                                   new SqlParameter { ParameterName = "@TranID", Value = TranID },
                                   new SqlParameter { ParameterName = "@Insurance", Value = Insurance },
                                   new SqlParameter { ParameterName = "@Salesby", Value = Salesby },
                                   new SqlParameter { ParameterName = "@Serviceby", Value = Serviceby },
                                   new SqlParameter { ParameterName = "@location", Value = location },
                                   new SqlParameter { ParameterName = "@Support", Value = Support },
                                   new SqlParameter { ParameterName = "@GodigitTansaction", Value = RDT },
                                   new SqlParameter { ParameterName = "@RFormat", Value = "F1" },
                                   new SqlParameter { ParameterName = "@InvNo", Value = "GGIC" },
                                   new SqlParameter { ParameterName = "@ReportId", Value = "GGI1" },
                                   new SqlParameter { ParameterName = "@DocName", Value = "Godigit General Insurance Co. Ltd." }
                                   );
            }
        }
        public static DataTable copyDatatable(DataTable dt)
        {
            DataTable dtcopy = new DataTable();
            DataRow dr = null;

            foreach (DataColumn column in dt.Columns)
            {
                if (column.ColumnName.ToLower() == "policy holder" || column.ColumnName.ToLower() == "insured person" || column.ColumnName.ToLower() == "master policy number" ||
                    column.ColumnName.ToLower() == "policy number" || column.ColumnName.ToLower() == "policy issue date" || column.ColumnName.ToLower() == "risk exp date" ||
                    column.ColumnName.ToLower() == "od premium" || column.ColumnName.ToLower() == "irda od amt" || column.ColumnName.ToLower() == "irda od %" || column.ColumnName.ToLower() == "tp premium" ||
                    column.ColumnName.ToLower() == "terrorism_premium" || column.ColumnName.ToLower() == "reward od %" || column.ColumnName.ToLower() == "reward tp%" ||
                    column.ColumnName.ToLower() == "irda reward amt" || column.ColumnName.ToLower() == "product name")
                {
                    dtcopy.Columns.Add(column.ColumnName);
                }
            }

            foreach (DataRow row in dt.Rows)
            {
                dr = dtcopy.NewRow();
                int j = 0; int k = 0;
                for (int i = 0; i < row.ItemArray.Length; i++)
                {
                    string name = dt.Columns[i].ColumnName.ToLower().ToString();
                    if (name == "policy holder" || name == "insured person" || name == "master policy number" ||
                    name == "policy number" || name == "policy issue date" || name == "risk exp date" ||
                    name == "od premium" || name == "irda od amt" || name == "irda od %" || name == "tp premium" ||
                    name == "terrorism_premium" || name == "reward od %" || name == "reward tp%" ||
                    name == "irda reward amt" || name == "product name")
                    {
                        if (k != 0)
                        {
                            j++;
                        }
                        k = 1;
                    }
                    if (name == "policy holder")
                    {
                        if (row[i].ToString() != "" && row[i].ToString() != null)
                        {
                            dr[j] = row[i].ToString();
                        }
                        else
                        {
                            dr[j] = "";
                        }
                    }
                    else if (name == "insured person")
                    {
                        dr[j] = row[i].ToString();
                    }
                    else if (name == "master policy number")
                    {
                        //string teamname = (row[i].ToString()).Substring(0, 3);
                        dr[j] = row[i].ToString();
                    }
                    else if (name == "policy number")
                    {
                        dr[j] = row[i].ToString();
                    }
                    else if (name == "policy issue date")
                    {
                        if (row[i].ToString() != "" && row[i].ToString() != null)
                        {
                            string strDate = row[i].ToString();
                            DateTime date = Convert.ToDateTime(strDate, CultureInfo.InvariantCulture);
                            dr[j] = date.ToString("yyyy-MM-dd");
                        }
                        else
                        {
                            dr[j] = row[i].ToString();
                        }
                    }
                    else if (name == "risk exp date")
                    {
                        if (row[i].ToString() != "" && row[i].ToString() != null)
                        {
                            string strDate = row[i].ToString();
                            DateTime date = Convert.ToDateTime(strDate, CultureInfo.InvariantCulture);
                            dr[j] = date.ToString("yyyy-MM-dd");
                        }
                        else
                        {
                            dr[j] = row[i].ToString();
                        }
                    }
                    else if (name == "od premium")
                    {
                        if (row[i].ToString() == "" || row[i].ToString() == null)
                        {
                            dr[j] = "0";
                        }
                        else
                        {
                            dr[j] = row[i].ToString();
                        }
                    }
                    else if (name == "irda od amt")
                    {
                        if (row[i].ToString() == "" || row[i].ToString() == null)
                        {
                            dr[j] = "0";
                        }
                        else
                        {
                            dr[j] = row[i].ToString();
                        }
                    }
                    else if (name == "irda od %")
                    {
                        if (row[i].ToString() == "" || row[i].ToString() == null)
                        {
                            dr[j] = "0";
                        }
                        else
                        {
                            dr[j] = row[i].ToString();
                        }
                    }
                    else if (name == "tp premium")
                    {
                        if (row[i].ToString() == "" || row[i].ToString() == null)
                        {
                            dr[j] = "0";
                        }
                        else
                        {
                            dr[j] = row[i].ToString();
                        }
                    }
                    else if (name == "terrorism_premium")
                    {
                        if (row[i].ToString() == "" || row[i].ToString() == null)
                        {
                            dr[j] = "0";
                        }
                        else
                        {
                            dr[j] = row[i].ToString();
                        }
                    }
                    else if (name == "reward od %")
                    {
                        if (row[i].ToString() == "" || row[i].ToString() == null)
                        {
                            dr[j] = "0";
                        }
                        else
                        {
                            dr[j] = row[i].ToString();
                        }
                    }
                    else if (name == "reward tp%")
                    {
                        if (row[i].ToString() == "" || row[i].ToString() == null)
                        {
                            dr[j] = "0";
                        }
                        else
                        {
                            dr[j] = row[i].ToString();
                        }
                    }
                    else if (name == "irda reward amt")
                    {
                        if (row[i].ToString() == "" || row[i].ToString() == null)
                        {
                            dr[j] = "0";
                        }
                        else
                        {
                            dr[j] = row[i].ToString();
                        }
                    }
                    else if (name == "product name")
                    {
                        if (row[i].ToString() == "" || row[i].ToString() == null)
                        {
                            dr[j] = "0";
                        }
                        else
                        {
                            dr[j] = row[i].ToString();
                        }
                    }
                    //else
                    //{
                    //    dr[i] = row[i].ToString();
                    //}
                }
                dtcopy.Rows.Add(dr);
            }
            return dtcopy;
        }
    }
}
