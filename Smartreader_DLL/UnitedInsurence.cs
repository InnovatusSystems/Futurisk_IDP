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
    public class UnitedInsurence
    {
        public static void DeleteRows(ISheet sheet)
        {
            var num = sheet.LastRowNum;
            int termsList1 = -1, termsList2 = -1;//, termsList3 = -1, termsList4 = -1;
            for (int rowIndex = sheet.LastRowNum; rowIndex >= 0; rowIndex--)
            {
                IRow row = sheet.GetRow(rowIndex);
                if (row == null) continue;
                ICell cell = row.GetCell(0);
                if (cell != null && cell.StringCellValue.Contains("Evaluation Only. Created with Aspose.PDF."))
                {
                    if (rowIndex != num)
                    {
                        sheet.ShiftRows(row.RowNum + 1, sheet.LastRowNum, -1);
                    }
                }
                ICell cell6 = row.GetCell(0);
                if (cell6 != null && cell6.StringCellValue.Contains("Dept Code"))
                {
                    if (rowIndex != 7 && rowIndex != 8)
                    {
                        termsList1 = rowIndex;
                    }
                }
                termsList2 = num - 1;

                if (termsList1 != -1 && termsList2 != -1)
                {

                    for (int list = termsList2; list >= termsList1; list--)
                    {
                        IRow row1 = sheet.GetRow(list);
                        sheet.ShiftRows(row1.RowNum + 1, sheet.LastRowNum, -1);
                    }
                    termsList2 = -1;
                    termsList1 = -1;
                }
            }
        }

        public static void InsertTransaction(Microsoft.Office.Interop.Excel.Workbook WB, string TranID, string RDate, string Insurance, string Salesby, string Serviceby, string location, string Support, string Rmonth, string strconn)
        {
            SQLProcs sql = new SQLProcs();
            Microsoft.Office.Interop.Excel.Worksheet wks = (Microsoft.Office.Interop.Excel.Worksheet)WB.Worksheets[1];
            Microsoft.Office.Interop.Excel.Range lastCell = wks.Cells.SpecialCells(Microsoft.Office.Interop.Excel.XlCellType.xlCellTypeLastCell, Type.Missing);
            int lastrow = lastCell.Row; 

            for (int i = 9; i < lastrow; i++)
            {
                var Terrorism = ""; var Policy_Endorsement = "";
                var InsuredName = ((Microsoft.Office.Interop.Excel.Range)wks.Cells[i, 4]).Value.Replace("\n", "").TrimStart();
                var InsuredType = ((Microsoft.Office.Interop.Excel.Range)wks.Cells[i, 9]).Value.Replace("\n", "").TrimStart();
                string PolicyNo = ((Microsoft.Office.Interop.Excel.Range)wks.Cells[i, 3]).Value.Replace("\n", "").TrimStart();
                string BillNo = ((Microsoft.Office.Interop.Excel.Range)wks.Cells[4, 7]).Value.Replace("\n", "").TrimStart();
                string LicenseNo = ((Microsoft.Office.Interop.Excel.Range)wks.Cells[6, 7]).Value.Replace("\n", "").TrimStart();
                string BRCode = ((Microsoft.Office.Interop.Excel.Range)wks.Cells[6, 2]).Value;
                string Endo_Effective_Date = ((Microsoft.Office.Interop.Excel.Range)wks.Cells[5, 7]).Value.Replace("\n", "").TrimStart();
                var Effective_Date = ((Microsoft.Office.Interop.Excel.Range)wks.Cells[5, 7]).Value.Replace("\n", "").TrimStart();
                var END_Date = ((Microsoft.Office.Interop.Excel.Range)wks.Cells[i, 5]).Value.Replace("\n", "").TrimStart();
                string Policy_Type = ((Microsoft.Office.Interop.Excel.Range)wks.Cells[i, 2]).Value.Replace("\n", "").TrimStart();
                var Premium_Amt = ((Microsoft.Office.Interop.Excel.Range)wks.Cells[i, 6]).Value.Replace("\n", "").Replace(",", "").TrimStart();
                var Revenue_Amt = ((Microsoft.Office.Interop.Excel.Range)wks.Cells[i, 8]).Value.Replace("\n", "").Replace(",", "").TrimStart();
                var Ineligible_Amt = ((Microsoft.Office.Interop.Excel.Range)wks.Cells[i, 7]).Value.Replace("\n", "").Replace(",", "").TrimStart();
                var DepCode = ((Microsoft.Office.Interop.Excel.Range)wks.Cells[i, 1]).Value.Replace("\n", "").TrimStart();
                var Presult = PolicyNo.Substring(PolicyNo.LastIndexOf('/') + 1);
                if (Policy_Type.Contains("Motor TP"))
                {
                    Terrorism = ((Microsoft.Office.Interop.Excel.Range)wks.Cells[i, 6]).Value.Replace("\n", "").Replace(",", "").TrimStart();
                    Premium_Amt = "0";
                }
                if (Presult.Contains("0"))
                {
                    Policy_Endorsement = "Policy";
                }
                else
                {
                    Policy_Endorsement = "Endorsement";
                }
                if (InsuredType == "Individual")
                {
                    InsuredType = "Retail";
                }
                if (Policy_Type == "Motor TP" || Policy_Type == "Motor")
                {
                    InsuredType = "Retail";
                }
                if (Endo_Effective_Date.Contains("Bill Date"))
                {
                    var BillDate = ((Microsoft.Office.Interop.Excel.Range)wks.Cells[5, 8]).Value;
                    if (BillDate != "" && BillDate != " ")
                    {
                        Endo_Effective_Date = ((Microsoft.Office.Interop.Excel.Range)wks.Cells[5, 8]).Value.Replace("\n", "").TrimStart();
                        Effective_Date = ((Microsoft.Office.Interop.Excel.Range)wks.Cells[5, 8]).Value.Replace("\n", "").TrimStart();
                    }
                }
                if (BillNo.Contains("Bill Number"))
                {
                    BillNo = ((Microsoft.Office.Interop.Excel.Range)wks.Cells[4, 8]).Value.Replace("\n", "").TrimStart();
                }
                if (LicenseNo.Contains("License Number"))
                {
                    LicenseNo = ((Microsoft.Office.Interop.Excel.Range)wks.Cells[6, 8]).Value.Replace("\n", "").TrimStart();
                }
                if (BRCode == "" || BRCode == " ")
                {
                    BRCode = ((Microsoft.Office.Interop.Excel.Range)wks.Cells[6, 3]).Value.Replace("\n", "").TrimStart();
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
                if (Ineligible_Amt == "" || Ineligible_Amt == " ")
                {
                    Ineligible_Amt = "0";
                }

                sql.ExecuteSQLNonQuery(strconn,"SP_Insert_Transactions",
                            new SqlParameter { ParameterName = "@Imode", Value = 1 },
                            new SqlParameter { ParameterName = "@RDate", Value = RDate },
                            new SqlParameter { ParameterName = "@Rmonth", Value = Rmonth },
                            new SqlParameter { ParameterName = "@ClientName", Value = InsuredName },
                            new SqlParameter { ParameterName = "@Itype", Value = InsuredType },
                            new SqlParameter { ParameterName = "@Policy_No", Value = PolicyNo },
                            new SqlParameter { ParameterName = "@Endo_Effective_Date", Value = Endo_Effective_Date },
                            new SqlParameter { ParameterName = "@Effective_Date", Value = Effective_Date },
                            new SqlParameter { ParameterName = "@END_Date", Value = END_Date },
                            new SqlParameter { ParameterName = "@Policy_Type", Value = Policy_Type },
                            new SqlParameter { ParameterName = "@Premium_Amt", Value = Premium_Amt },
                            new SqlParameter { ParameterName = "@TranID", Value = TranID },
                            new SqlParameter { ParameterName = "@Revenue_Amt", Value = Revenue_Amt },
                            new SqlParameter { ParameterName = "@Terrorism", Value = Terrorism },
                            new SqlParameter { ParameterName = "@Ineligible_Amt", Value = Ineligible_Amt },
                            new SqlParameter { ParameterName = "@DeptCode", Value = DepCode },
                            new SqlParameter { ParameterName = "@BillNo", Value = BillNo },
                            new SqlParameter { ParameterName = "@LicenseNo", Value = LicenseNo },
                            new SqlParameter { ParameterName = "@BRCode", Value = BRCode },
                            new SqlParameter { ParameterName = "@Insurance", Value = Insurance },
                            new SqlParameter { ParameterName = "@Salesby", Value = Salesby },
                            new SqlParameter { ParameterName = "@Serviceby", Value = Serviceby },
                            new SqlParameter { ParameterName = "@location", Value = location },
                            new SqlParameter { ParameterName = "@Support", Value = Support },
                            new SqlParameter { ParameterName = "@Policy_Endorsement", Value = Policy_Endorsement },
                            new SqlParameter { ParameterName = "@RFormat", Value = "F1" },
                            new SqlParameter { ParameterName = "@InvNo", Value = "UIIP" },
                            new SqlParameter { ParameterName = "@DocName", Value = "United India Insurance Co.Ltd." }
                            );
            }


            //DataSet ds = new DataSet();

            //ds = sql.SQLExecuteDataset(strconn,"SP_Insert_Transactions",
            //     new SqlParameter { ParameterName = "@Imode", Value = 4 },
            //     new SqlParameter { ParameterName = "@TranID", Value = TranID }
            //     );

            //if (ds != null && ds.Tables.Count > 0 && ds.Tables[0].Rows.Count > 0)
            //{
            //    BatchID = ds.Tables[0].Rows[0]["BatchID"].ToString();
            //}
            //else
            //{
            //    BatchID = "";
            //}

            //DataSet dsR = new DataSet();
            //dsR = sql.SQLExecuteDataset(strconn,"SP_Insert_Transactions",
            //     new SqlParameter { ParameterName = "@Imode", Value = 5 },
            //     new SqlParameter { ParameterName = "@BatchID", Value = BatchID },
            //     new SqlParameter { ParameterName = "@Filename", Value = Fileinfo.Filename },
            //     new SqlParameter { ParameterName = "@version", Value = LoginInfo.version },
            //     new SqlParameter { ParameterName = "@UserId", Value = LoginInfo.UserID }
            //     );
            //if (dsR != null && dsR.Tables.Count > 0 && dsR.Tables[0].Rows.Count > 0)
            //{
            //    NoRecord = dsR.Tables[0].Rows[0]["NoRecord"].ToString();
            //}
            //else
            //{
            //    NoRecord = "";
            //}
        }

        public static void InsertExcelTransaction(Microsoft.Office.Interop.Excel.Workbook WB, string TranID, string RDate, string Insurance, string Salesby, string Serviceby, string location, string Support, string Rmonth, string strconn)
        {
            //United India Insurance Excel
            SQLProcs sql = new SQLProcs();
            Microsoft.Office.Interop.Excel.Worksheet wks = (Microsoft.Office.Interop.Excel.Worksheet)WB.Worksheets[1];
            Microsoft.Office.Interop.Excel.Range lastCell = wks.Cells.SpecialCells(Microsoft.Office.Interop.Excel.XlCellType.xlCellTypeLastCell, Type.Missing);
            int lastrow = lastCell.Row; var Terrorism = ""; var Policy_Endorsement = ""; var InsuredType = ""; var Policy_Type = ""; var New_Renewal = ""; var Revenue_Amt = "";
            for (int i = 2; i <= lastrow; i++)
            {
                string InsuredName = ((Microsoft.Office.Interop.Excel.Range)wks.Cells[i, 15]).Value;
                if (InsuredName != null && InsuredName != "" && InsuredName != " ")
                {
                    InsuredName = InsuredName.Replace("\n", "").TrimStart();
                    //InsuredType = ((Microsoft.Office.Interop.Excel.Range)wks.Cells[i, 8]).Value;
                    string PolicyNo = Convert.ToString(((Microsoft.Office.Interop.Excel.Range)wks.Cells[i, 13]).Value).Replace("\n", "").TrimStart();
                    //var Endorsementno = Convert.ToString(((Microsoft.Office.Interop.Excel.Range)wks.Cells[i, 15]).Value);

                    var Endo_Effective_Date = ((Microsoft.Office.Interop.Excel.Range)wks.Cells[i, 5]).Value;
                    //var Effective_Date = ((Microsoft.Office.Interop.Excel.Range)wks.Cells[i, 5]).Value;
                    var END_Date = ((Microsoft.Office.Interop.Excel.Range)wks.Cells[i, 6]).Value;
                    int ENdolen = Convert.ToString(Endo_Effective_Date).Length;
                    if (ENdolen > 11)
                    {
                        string date = Convert.ToString(Endo_Effective_Date);
                        var yy = date.Substring(6, 4);
                        var MM = date.Substring(3, 2);
                        var DD = date.Substring(0, 2);
                        Endo_Effective_Date = DD + "/" + MM + "/" + yy;
                        //Endo_Effective_Date = Endo_Effective_Date.ToString("dd/MM/yyyy");
                    }
                    //int Efflen = Convert.ToString(Effective_Date).Length;
                    //if (Efflen > 11)
                    //{
                    //    Effective_Date = Effective_Date.ToString("dd/MM/yyyy");
                    //}
                    int ENDlen = Convert.ToString(END_Date).Length;
                    if (ENDlen > 11)
                    {
                        string date = Convert.ToString(END_Date);
                        var yy = date.Substring(6, 4);
                        var MM = date.Substring(3, 2);
                        var DD = date.Substring(0, 2);
                        END_Date = DD + "/" + MM + "/" + yy;
                        //END_Date = END_Date.ToString("dd/MM/yyyy");
                    }
                    var Premium_Amt = Convert.ToString(((Microsoft.Office.Interop.Excel.Range)wks.Cells[i, 8]).Value).Replace(",", "").Replace("(", "").Replace(")", "").TrimStart();
                    //Terrorism = Convert.ToString(((Microsoft.Office.Interop.Excel.Range)wks.Cells[i, 23]).Value).Replace(",", "").Replace("(", "").Replace(")", "").TrimStart();
                    Revenue_Amt = Convert.ToString(((Microsoft.Office.Interop.Excel.Range)wks.Cells[i, 10]).Value).Replace(",", "").Replace("(", "").Replace(")", "").Replace("-", "").TrimStart();
                    var Revenue_Pcnt = Convert.ToString(((Microsoft.Office.Interop.Excel.Range)wks.Cells[i, 7]).Text);
                    Revenue_Pcnt = Revenue_Pcnt.Replace("\n", "").Replace("%", "").Replace(",", "").TrimStart();

                    //Policy_Type = Convert.ToString(((Microsoft.Office.Interop.Excel.Range)wks.Cells[i, 7]).Value);
                    if (InsuredName.Contains("LIMITED") || InsuredName.Contains("LTD"))
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
                    sql.ExecuteSQLNonQuery(strconn,"SP_UnitedExcelTransaction",
                               new SqlParameter { ParameterName = "@Imode", Value = 1 },
                               new SqlParameter { ParameterName = "@RDate", Value = RDate },
                               new SqlParameter { ParameterName = "@Rmonth", Value = Rmonth },
                               new SqlParameter { ParameterName = "@ClientName", Value = InsuredName },
                               new SqlParameter { ParameterName = "@Itype", Value = InsuredType },
                               new SqlParameter { ParameterName = "@Policy_No", Value = PolicyNo },
                               new SqlParameter { ParameterName = "@Endo_Effective_Date", Value = Endo_Effective_Date },
                               //new SqlParameter { ParameterName = "@Effective_Date", Value = Effective_Date },
                               new SqlParameter { ParameterName = "@END_Date", Value = END_Date },
                               new SqlParameter { ParameterName = "@Revenue_Pct", Value = Revenue_Pcnt },
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
                               new SqlParameter { ParameterName = "@InvNo", Value = "UIIX" },
                               new SqlParameter { ParameterName = "@ReportId", Value = "UIIX" },
                               new SqlParameter { ParameterName = "@DocName", Value = "United India Insurance" }
                               );
                }
            }
        }


    }
}
