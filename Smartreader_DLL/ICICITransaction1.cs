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
    public class ICICITransaction1
    {
        public static void ExceltoDatatable1(string TranID, string RDate, string Insurance, string Salesby, string Serviceby, string location, string Support, string Rmonth, string Filepath, string strconn)
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

                DataTable RDT = copyDatatable1(DT);

                //DataSet ds = new DataSet();

                //ds = 
                sql.ExecuteSQLNonQuery(strconn,"SP_ICICITransactions",
                                   new SqlParameter { ParameterName = "@Imode", Value = 9 },
                                   new SqlParameter { ParameterName = "@RDate", Value = RDate },
                                   new SqlParameter { ParameterName = "@Rmonth", Value = Rmonth },
                                   new SqlParameter { ParameterName = "@TranID", Value = TranID },
                                   new SqlParameter { ParameterName = "@Insurance", Value = Insurance },
                                   new SqlParameter { ParameterName = "@Salesby", Value = Salesby },
                                   new SqlParameter { ParameterName = "@Serviceby", Value = Serviceby },
                                   new SqlParameter { ParameterName = "@location", Value = location },
                                   new SqlParameter { ParameterName = "@Support", Value = Support },
                                   new SqlParameter { ParameterName = "@ICICITransaction", Value = RDT },
                                   new SqlParameter { ParameterName = "@RFormat", Value = "F1" },
                                   new SqlParameter { ParameterName = "@InvNo", Value = "ILGI" },
                                   new SqlParameter { ParameterName = "@ReportId", Value = "ILG1" },
                                   new SqlParameter { ParameterName = "@DocName", Value = "ICICI General Insurance Co. Ltd." }
                                   );

            }
        }
        public static DataTable copyDatatable1(DataTable dt)
        {
            DataTable dtcopy = new DataTable();
            DataRow dr = null;

            foreach (DataColumn column in dt.Columns)
            {
                if (column.ColumnName.ToUpper() == "CUSTOMER_NAME" || column.ColumnName.ToUpper() == "POL_NUM_TXT" || column.ColumnName.ToUpper() == "POLICY_START_DATE" ||
                    column.ColumnName.ToUpper() == "POLICY_END_DATE" || column.ColumnName.ToUpper() == "PRODUCT_NAME" ||
                    column.ColumnName.ToUpper() == "PREMIUM_FOR_PAYOUTS" || column.ColumnName.ToUpper() == "TERRORISM_PREMIUM_AMOUNT" ||
                    column.ColumnName.ToUpper() == "COMMISSION_PAYOUTS_PERCENTAGE" || column.ColumnName.ToUpper() == "ACTUAL_COMMISSION_AMOUNT")
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
                    string name = dt.Columns[i].ColumnName.ToUpper().ToString();
                    if (name == "CUSTOMER_NAME" || name == "POL_NUM_TXT" || name == "POLICY_START_DATE" ||
                    name == "POLICY_END_DATE" || name == "PRODUCT_NAME" || name == "PREMIUM_FOR_PAYOUTS" || name == "TERRORISM_PREMIUM_AMOUNT" ||
                    name == "COMMISSION_PAYOUTS_PERCENTAGE" || name == "ACTUAL_COMMISSION_AMOUNT")
                    {
                        if (k != 0)
                        {
                            j++;
                        }
                        k = 1;
                    }
                    if (name == "CUSTOMER_NAME")
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
                    else if (name == "POL_NUM_TXT")
                    {
                        dr[j] = row[i].ToString();
                    }
                    else if (name == "POLICY_START_DATE")
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
                    else if (name == "POLICY_END_DATE")
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
                    else if (name == "PRODUCT_NAME")
                    {
                        dr[j] = row[i].ToString();
                    }
                    else if (name == "PREMIUM_FOR_PAYOUTS")
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
                    else if (name == "TERRORISM_PREMIUM_AMOUNT")
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
                    else if (name == "COMMISSION_PAYOUTS_PERCENTAGE")
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
                    else if (name == "ACTUAL_COMMISSION_AMOUNT")
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
                }
                dtcopy.Rows.Add(dr);
            }
            return dtcopy;
        }
        public static void InsertTransaction(Microsoft.Office.Interop.Excel.Workbook WB, string UserID,  string strconn)
        {
            SQLProcs sql = new SQLProcs();
            Microsoft.Office.Interop.Excel.Worksheet wks = (Microsoft.Office.Interop.Excel.Worksheet)WB.Worksheets[1];
            Microsoft.Office.Interop.Excel.Range lastCell = wks.Cells.SpecialCells(Microsoft.Office.Interop.Excel.XlCellType.xlCellTypeLastCell, Type.Missing);
            int lastrow = lastCell.Row; var Terrorism = "";
            for (int i = 2; i <= lastrow; i++)
            {
                string PolicyNo = ((Microsoft.Office.Interop.Excel.Range)wks.Cells[i, 14]).Value.Replace("\n", "").TrimStart();
                if (PolicyNo == "" || PolicyNo == " " || PolicyNo == null)
                {
                    var Revenue_Amt = Convert.ToString(((Microsoft.Office.Interop.Excel.Range)wks.Cells[i, 26]).Value);
                    Terrorism = Convert.ToString(((Microsoft.Office.Interop.Excel.Range)wks.Cells[i, 42]).Value);
                    if (Revenue_Amt == "" || Revenue_Amt == " " || Terrorism == null)
                    {
                        Revenue_Amt = 0;
                    }
                    if (Terrorism == "" || Terrorism == " " || Terrorism == null)
                    {
                        Terrorism = "0";
                    }
                    sql.ExecuteSQLNonQuery(strconn,"SP_ICICITransactions",
                               new SqlParameter { ParameterName = "@Imode", Value = 10 },
                               new SqlParameter { ParameterName = "@Policy_No", Value = PolicyNo },
                               new SqlParameter { ParameterName = "@Revenue_Amt", Value = Revenue_Amt },
                               new SqlParameter { ParameterName = "@Terrorism", Value = Terrorism },
                               new SqlParameter { ParameterName = "@RFormat", Value = "F1" },
                               new SqlParameter { ParameterName = "@InvNo", Value = "ILGI" },
                               new SqlParameter { ParameterName = "@ReportId", Value = "ILG1" },
                               new SqlParameter { ParameterName = "@UserId", Value = UserID }
                               );
                }
            }
        }
    }
}
