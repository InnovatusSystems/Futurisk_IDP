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
    public class NewIndiaPDFInsurancedll
    {
        public static string InsertTransaction1(Microsoft.Office.Interop.Excel.Workbook WB, string TranID, string RDate, string Insurance, string Salesby, string Serviceby, string location, string Support, string Rmonth, string strconn)
        {
            SQLProcs sql = new SQLProcs();
            Microsoft.Office.Interop.Excel.Worksheet wks = (Microsoft.Office.Interop.Excel.Worksheet)WB.Worksheets[2];
            Microsoft.Office.Interop.Excel.Range lastCell = wks.Cells.SpecialCells(Microsoft.Office.Interop.Excel.XlCellType.xlCellTypeLastCell, Type.Missing);
            int lastrow = lastCell.Row; string PolicyNo = ""; int J = 0;
            string Checknull = ((Microsoft.Office.Interop.Excel.Range)wks.Cells[1, 13]).Value;
            string CheckFormat = ((Microsoft.Office.Interop.Excel.Range)wks.Cells[1, 3]).Value;
            string FormatResult = "Not OK";
            if (Checknull != null && Checknull.Contains("Insured Name"))
            {
                J = 1;
            }
            if (CheckFormat != null)
            {
                CheckFormat = CheckFormat.Replace("\n", "").TrimStart();
                if (CheckFormat.Contains("Office Code") || CheckFormat.Contains("OfficeCode"))
                {
                    FormatResult = "OK";
                }
                else
                {
                    FormatResult = "Not OK";
                }
            }
            if (FormatResult == "OK")
            {
                for (int i = 2; i <= lastrow; i++)
                {
                    var Terrorism = ""; var Policy_Endorsement = ""; var InsuredType = ""; var Policy_Type = ""; var New_Renewal = ""; var Revenue_Amt = "";

                    string InsuredName = ((Microsoft.Office.Interop.Excel.Range)wks.Cells[i, 14 - J]).Value;
                    InsuredType = ((Microsoft.Office.Interop.Excel.Range)wks.Cells[i, 16 - J]).Value;
                    if (InsuredType != null && InsuredType != "" && InsuredType != " ")
                    {
                        InsuredName = InsuredName.Replace("\n", " ").TrimStart();
                        string Pno = Convert.ToString(((Microsoft.Office.Interop.Excel.Range)wks.Cells[i, 4]).Value);
                        //var Endorsementno = Convert.ToString(((Microsoft.Office.Interop.Excel.Range)wks.Cells[i, 15]).Value);
                        if (Pno != null && Pno != "" && Pno != " ")
                        {
                            PolicyNo = Pno.Replace("\n", "").TrimStart();
                        }
                        var Endo_Effective_Date = ((Microsoft.Office.Interop.Excel.Range)wks.Cells[i, 18 - J]).Value;
                        var Effective_Date = ((Microsoft.Office.Interop.Excel.Range)wks.Cells[i, 18 - J]).Value;
                        var END_Date = ((Microsoft.Office.Interop.Excel.Range)wks.Cells[i, 19 - J]).Value;
                        int ENdolen = Convert.ToString(Endo_Effective_Date).Length;
                        if (ENdolen > 11)
                        {
                            Endo_Effective_Date = Endo_Effective_Date.ToString("dd/MM/yyyy");
                        }
                        int Efflen = Convert.ToString(Effective_Date).Length;
                        if (Efflen > 11)
                        {
                            Effective_Date = Effective_Date.ToString("dd/MM/yyyy");
                        }
                        int ENDlen = Convert.ToString(END_Date).Length;
                        if (ENDlen > 11)
                        {
                            END_Date = END_Date.ToString("dd/MM/yyyy");
                        }
                        var Premium_Amt = Convert.ToString(((Microsoft.Office.Interop.Excel.Range)wks.Cells[i, 21 - J]).Value).Replace(",", "").Replace("(", "").Replace(")", "").Replace(".", "").TrimStart();
                        //Terrorism = Convert.ToString(((Microsoft.Office.Interop.Excel.Range)wks.Cells[i, 23]).Value).Replace(",", "").Replace("(", "").Replace(")", "").TrimStart();
                        Revenue_Amt = Convert.ToString(((Microsoft.Office.Interop.Excel.Range)wks.Cells[i, 22 - J]).Value).Replace(",", "").Replace("(", "").Replace(")", "").Replace(".", "").TrimStart();
                        var Revenue_Pcnt = Convert.ToString(((Microsoft.Office.Interop.Excel.Range)wks.Cells[i, 20 - J]).Text);
                        Revenue_Pcnt = Revenue_Pcnt.Replace("\n", "").Replace("%", "").Replace(",", "").TrimStart();
                        PolicyNo = Regex.Replace(PolicyNo, "[^a-zA-Z0-9_.]+", "", RegexOptions.Compiled);
                        string endno = Convert.ToString(((Microsoft.Office.Interop.Excel.Range)wks.Cells[i, 8]).Value).Replace("\n", "").TrimStart();
                        if (endno != null && endno != "" && endno != " " && endno != ":")
                        {
                            Policy_Endorsement = "Endorsement";
                        }
                        else
                        {
                            Policy_Endorsement = "Policy";
                        }
                        Policy_Type = Convert.ToString(((Microsoft.Office.Interop.Excel.Range)wks.Cells[i, 12]).Value);
                        //if (InsuredName.ToUpper().Contains("LIMITED") || InsuredName.ToUpper().Contains("LTD"))
                        if (InsuredType == "Organizational")
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
                        sql.ExecuteSQLNonQuery(strconn,"SP_NewIndiaPDFTransaction",
                                   new SqlParameter { ParameterName = "@Imode", Value = 1 },
                                   new SqlParameter { ParameterName = "@RDate", Value = RDate },
                                   new SqlParameter { ParameterName = "@Rmonth", Value = Rmonth },
                                   new SqlParameter { ParameterName = "@ClientName", Value = InsuredName },
                                   new SqlParameter { ParameterName = "@Itype", Value = InsuredType },
                                   new SqlParameter { ParameterName = "@Policy_No", Value = PolicyNo },
                                   new SqlParameter { ParameterName = "@Revenue_Pct", Value = Revenue_Pcnt },
                                   new SqlParameter { ParameterName = "@Policy_Type", Value = Policy_Type },
                                   new SqlParameter { ParameterName = "@Premium_Amt", Value = Premium_Amt },
                                   new SqlParameter { ParameterName = "@Endo_Effective_Date", Value = Endo_Effective_Date },
                                   new SqlParameter { ParameterName = "@Effective_Date", Value = Effective_Date },
                                   new SqlParameter { ParameterName = "@END_Date", Value = END_Date },
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
                                   new SqlParameter { ParameterName = "@InvNo", Value = "NIP1" },
                                   new SqlParameter { ParameterName = "@ReportId", Value = "NIP1" },
                                   new SqlParameter { ParameterName = "@DocName", Value = "New India Assurance Company Ltd." }
                                   );
                    }
                }
            }
            return FormatResult;
        }

        public static string InsertTransaction2(Microsoft.Office.Interop.Excel.Workbook WB, string TranID, string RDate, string Insurance, string Salesby, string Serviceby, string location, string Support, string Rmonth, string strconn)
        {
            //New India Assurance Company Limited.
            SQLProcs sql = new SQLProcs();
            Microsoft.Office.Interop.Excel.Worksheet wks = (Microsoft.Office.Interop.Excel.Worksheet)WB.Worksheets[2];
            Microsoft.Office.Interop.Excel.Range lastCell = wks.Cells.SpecialCells(Microsoft.Office.Interop.Excel.XlCellType.xlCellTypeLastCell, Type.Missing);
            int lastrow = lastCell.Row;
            string CheckFormat = ((Microsoft.Office.Interop.Excel.Range)wks.Cells[1, 3]).Value;
            string FormatResult = "Not OK";
            if (CheckFormat != null)
            {
                CheckFormat = CheckFormat.Replace("\n", "").TrimStart();
                if (CheckFormat.Contains("Dept"))
                {
                    FormatResult = "OK";
                }
                else
                {
                    FormatResult = "Not OK";
                }
            }
            if (FormatResult == "OK")
            {
                for (int i = 2; i <= lastrow; i++)
                {
                    var Terrorism = ""; var Policy_Endorsement = ""; var InsuredType = ""; var Policy_Type = ""; var New_Renewal = ""; var Revenue_Amt = "";
                    string InsuredName = ((Microsoft.Office.Interop.Excel.Range)wks.Cells[i, 9]).Value;
                    if (InsuredName != null && InsuredName != "" && InsuredName != " ")
                    {
                        InsuredName = InsuredName.Replace("\n", " ").TrimStart();
                        InsuredType = ((Microsoft.Office.Interop.Excel.Range)wks.Cells[i, 12]).Value;
                        string PolicyNo = Convert.ToString(((Microsoft.Office.Interop.Excel.Range)wks.Cells[i, 4]).Value).Replace("\n", "").TrimStart();
                        //var Endorsementno = Convert.ToString(((Microsoft.Office.Interop.Excel.Range)wks.Cells[i, 15]).Value);

                        var Endo_Effective_Date = ((Microsoft.Office.Interop.Excel.Range)wks.Cells[i, 15]).Value;
                        var Effective_Date = ((Microsoft.Office.Interop.Excel.Range)wks.Cells[i, 15]).Value;
                        var END_Date = ((Microsoft.Office.Interop.Excel.Range)wks.Cells[i, 17]).Value;
                        int ENdolen = Convert.ToString(Endo_Effective_Date).Length;
                        if (ENdolen > 11)
                        {
                            Endo_Effective_Date = Endo_Effective_Date.ToString("dd/MM/yyyy");
                        }
                        int Efflen = Convert.ToString(Effective_Date).Length;
                        if (Efflen > 11)
                        {
                            Effective_Date = Effective_Date.ToString("dd/MM/yyyy");
                        }
                        int ENDlen = Convert.ToString(END_Date).Length;
                        if (ENDlen > 11)
                        {
                            END_Date = END_Date.ToString("dd/MM/yyyy");
                        }
                        var Premium_Amt = Convert.ToString(((Microsoft.Office.Interop.Excel.Range)wks.Cells[i, 20]).Value);
                        Premium_Amt = Premium_Amt.Replace(",", "").Replace(".", "").Replace("(", "").Replace(")", "").TrimStart();
                        //Terrorism = Convert.ToString(((Microsoft.Office.Interop.Excel.Range)wks.Cells[i, 23]).Value).Replace(",", "").Replace("(", "").Replace(")", "").TrimStart();
                        Revenue_Amt = Convert.ToString(((Microsoft.Office.Interop.Excel.Range)wks.Cells[i, 21]).Value);
                        Revenue_Amt = Revenue_Amt.Replace(",", "").Replace("(", "").Replace(")", "").Replace(".", "").TrimStart();
                        var Revenue_Pcnt = Convert.ToString(((Microsoft.Office.Interop.Excel.Range)wks.Cells[i, 18]).Text);
                        Revenue_Pcnt = Revenue_Pcnt.Replace("\n", "").Replace("%", "").Replace(",", "").TrimStart();
                        if (PolicyNo != null && PolicyNo != "" && PolicyNo != " ")
                        {
                            PolicyNo = Regex.Replace(PolicyNo, "[^a-zA-Z0-9_.]+", "", RegexOptions.Compiled);
                        }
                        Policy_Type = Convert.ToString(((Microsoft.Office.Interop.Excel.Range)wks.Cells[i, 3]).Value);
                        if (InsuredType == "Organizational")
                        {
                            InsuredType = "Corporate";
                        }
                        else
                        {
                            InsuredType = "Retail";
                        }
                        string endno = Convert.ToString(((Microsoft.Office.Interop.Excel.Range)wks.Cells[i, 7]).Value).Replace("\n", "").TrimStart();
                        if (endno != null && endno != "" && endno != " " && endno != ":")
                        {
                            Policy_Endorsement = "Endorsement";
                            PolicyNo = PolicyNo + " " + Regex.Replace(endno, "[^a-zA-Z0-9_.]+", "", RegexOptions.Compiled);
                        }
                        else
                        {
                            Policy_Endorsement = "Policy";
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
                        sql.ExecuteSQLNonQuery(strconn,"SP_NewIndiaPDFTransaction",
                                   new SqlParameter { ParameterName = "@Imode", Value = 1 },
                                   new SqlParameter { ParameterName = "@RDate", Value = RDate },
                                   new SqlParameter { ParameterName = "@Rmonth", Value = Rmonth },
                                   new SqlParameter { ParameterName = "@ClientName", Value = InsuredName },
                                   new SqlParameter { ParameterName = "@Itype", Value = InsuredType },
                                   new SqlParameter { ParameterName = "@Policy_No", Value = PolicyNo },
                                   new SqlParameter { ParameterName = "@Revenue_Pct", Value = Revenue_Pcnt },
                                   new SqlParameter { ParameterName = "@Policy_Type", Value = Policy_Type },
                                   new SqlParameter { ParameterName = "@Premium_Amt", Value = Premium_Amt },
                                   new SqlParameter { ParameterName = "@Endo_Effective_Date", Value = Endo_Effective_Date },
                                   new SqlParameter { ParameterName = "@Effective_Date", Value = Effective_Date },
                                   new SqlParameter { ParameterName = "@END_Date", Value = END_Date },
                                   new SqlParameter { ParameterName = "@TranID", Value = TranID },
                                   new SqlParameter { ParameterName = "@Revenue_Amt", Value = Revenue_Amt },
                                   new SqlParameter { ParameterName = "@Terrorism", Value = Terrorism },
                                   new SqlParameter { ParameterName = "@Insurance", Value = Insurance },
                                   new SqlParameter { ParameterName = "@Salesby", Value = Salesby },
                                   new SqlParameter { ParameterName = "@Serviceby", Value = Serviceby },
                                   new SqlParameter { ParameterName = "@location", Value = location },
                                   new SqlParameter { ParameterName = "@Support", Value = Support },
                                   new SqlParameter { ParameterName = "@Policy_Endorsement", Value = Policy_Endorsement },
                                   new SqlParameter { ParameterName = "@RFormat", Value = "F2" },
                                   new SqlParameter { ParameterName = "@InvNo", Value = "NIP2" },
                                   new SqlParameter { ParameterName = "@ReportId", Value = "NIP2" },
                                   new SqlParameter { ParameterName = "@DocName", Value = "New India Assurance Company Ltd." }
                                   );
                    }
                }
            }
            return FormatResult;
        }
        public static string InsertTransaction3(Microsoft.Office.Interop.Excel.Workbook WB, string TranID, string RDate, string Insurance, string Salesby, string Serviceby, string location, string Support, string Rmonth, string strconn)
        {
            //New India Assurance Company Limited.
            SQLProcs sql = new SQLProcs();
            Microsoft.Office.Interop.Excel.Worksheet wks = (Microsoft.Office.Interop.Excel.Worksheet)WB.Worksheets[2];
            Microsoft.Office.Interop.Excel.Range lastCell = wks.Cells.SpecialCells(Microsoft.Office.Interop.Excel.XlCellType.xlCellTypeLastCell, Type.Missing);
            int lastrow = lastCell.Row; string Policy_Type = ""; var PType = "";

            string CheckFormat = ((Microsoft.Office.Interop.Excel.Range)wks.Cells[1, 1]).Value;
            string FormatResult = "Not OK";
            if (CheckFormat != null)
            {
                CheckFormat = CheckFormat.Replace("\n", "").TrimStart();
                if (CheckFormat.Contains("Department"))
                {
                    FormatResult = "OK";
                }
                else
                {
                    FormatResult = "Not OK";
                }
            }
            if (FormatResult == "OK")
            {
                for (int i = 2; i <= lastrow; i++)
                {
                    var Terrorism = ""; var Policy_Endorsement = ""; var InsuredType = ""; var Revenue_Pcnt = ""; var Revenue_Amt = ""; var PolicyNo = "";
                    string InsuredName = ((Microsoft.Office.Interop.Excel.Range)wks.Cells[i, 6]).Value;
                    if (InsuredName != null && InsuredName != "" && InsuredName != " " && InsuredName != "Insured Name")
                    {
                        InsuredName = InsuredName.Replace("\n", " ").TrimStart();
                        InsuredType = ((Microsoft.Office.Interop.Excel.Range)wks.Cells[i, 7]).Value;
                        var PNo = Convert.ToString(((Microsoft.Office.Interop.Excel.Range)wks.Cells[i, 2]).Value);
                        if (PNo != null && PNo != "" && PNo != " ")
                        {
                            PType = PNo.Replace("\n", "").TrimStart();
                            PType = Regex.Replace(PType, "[^a-zA-Z0-9_.]+", "", RegexOptions.Compiled);
                        }
                        var endno = Convert.ToString(((Microsoft.Office.Interop.Excel.Range)wks.Cells[i, 5]).Value);
                        if (endno != null && endno != "" && endno != " " && endno != ":")
                        {
                            Policy_Endorsement = "Endorsement";
                            PolicyNo = PType + " " + Regex.Replace(endno, "[^a-zA-Z0-9_.]+", "", RegexOptions.Compiled);
                        }
                        else
                        {
                            Policy_Endorsement = "Policy";
                            PolicyNo = PType;
                        }
                        if (InsuredType == "Organizational")
                        {
                            InsuredType = "Corporate";
                        }
                        else
                        {
                            InsuredType = "Retail";
                        }
                        var Endo_Effective_Date = ((Microsoft.Office.Interop.Excel.Range)wks.Cells[i, 8]).Value;
                        var Effective_Date = ((Microsoft.Office.Interop.Excel.Range)wks.Cells[i, 8]).Value;
                        var END_Date = ((Microsoft.Office.Interop.Excel.Range)wks.Cells[i, 9]).Value;
                        int ENdolen = Convert.ToString(Endo_Effective_Date).Length;
                        if (ENdolen > 11)
                        {
                            Endo_Effective_Date = Endo_Effective_Date.ToString("dd/MM/yyyy");
                        }
                        int Efflen = Convert.ToString(Effective_Date).Length;
                        if (Efflen > 11)
                        {
                            Effective_Date = Effective_Date.ToString("dd/MM/yyyy");
                        }
                        int ENDlen = Convert.ToString(END_Date).Length;
                        if (ENDlen > 11)
                        {
                            END_Date = END_Date.ToString("dd/MM/yyyy");
                        }
                        var Premium_Amt = Convert.ToString(((Microsoft.Office.Interop.Excel.Range)wks.Cells[i, 10]).Value).Replace(",", "").Replace("(", "").Replace(")", "").Replace(".", "").TrimStart();
                        //Terrorism = Convert.ToString(((Microsoft.Office.Interop.Excel.Range)wks.Cells[i, 23]).Value).Replace(",", "").Replace("(", "").Replace(")", "").TrimStart();
                        Revenue_Amt = Convert.ToString(((Microsoft.Office.Interop.Excel.Range)wks.Cells[i, 11]).Value).Replace(",", "").Replace("(", "").Replace(")", "").Replace(".", "").TrimStart();
                        //var Revenue_Pcnt = Convert.ToString(((Microsoft.Office.Interop.Excel.Range)wks.Cells[i, 21]).Text);
                        //Revenue_Pcnt = Revenue_Pcnt.Replace("\n", "").Replace("%", "").Replace(",", "").TrimStart();

                        string dept = Convert.ToString(((Microsoft.Office.Interop.Excel.Range)wks.Cells[i, 1]).Value);
                        if (dept != null && dept != "" && dept != " ")
                        {
                            Policy_Type = Regex.Match(dept, @"\d+").Value;
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
                        sql.ExecuteSQLNonQuery(strconn,"SP_NewIndiaPDFTransaction",
                                   new SqlParameter { ParameterName = "@Imode", Value = 1 },
                                   new SqlParameter { ParameterName = "@RDate", Value = RDate },
                                   new SqlParameter { ParameterName = "@Rmonth", Value = Rmonth },
                                   new SqlParameter { ParameterName = "@ClientName", Value = InsuredName },
                                   new SqlParameter { ParameterName = "@Itype", Value = InsuredType },
                                   new SqlParameter { ParameterName = "@Policy_No", Value = PolicyNo },
                                   new SqlParameter { ParameterName = "@Revenue_Pct", Value = Revenue_Pcnt },
                                   new SqlParameter { ParameterName = "@Policy_Type", Value = Policy_Type },
                                   new SqlParameter { ParameterName = "@Premium_Amt", Value = Premium_Amt },
                                   new SqlParameter { ParameterName = "@Endo_Effective_Date", Value = Endo_Effective_Date },
                                   new SqlParameter { ParameterName = "@Effective_Date", Value = Effective_Date },
                                   new SqlParameter { ParameterName = "@END_Date", Value = END_Date },
                                   new SqlParameter { ParameterName = "@TranID", Value = TranID },
                                   new SqlParameter { ParameterName = "@Revenue_Amt", Value = Revenue_Amt },
                                   new SqlParameter { ParameterName = "@Terrorism", Value = Terrorism },
                                   new SqlParameter { ParameterName = "@Insurance", Value = Insurance },
                                   new SqlParameter { ParameterName = "@Salesby", Value = Salesby },
                                   new SqlParameter { ParameterName = "@Serviceby", Value = Serviceby },
                                   new SqlParameter { ParameterName = "@location", Value = location },
                                   new SqlParameter { ParameterName = "@Support", Value = Support },
                                   new SqlParameter { ParameterName = "@Policy_Endorsement", Value = Policy_Endorsement },
                                   new SqlParameter { ParameterName = "@RFormat", Value = "F3" },
                                   new SqlParameter { ParameterName = "@InvNo", Value = "NIP3" },
                                   new SqlParameter { ParameterName = "@ReportId", Value = "NIP3" },
                                   new SqlParameter { ParameterName = "@DocName", Value = "New India Assurance Company Ltd." }
                                   );
                    }
                }
            }
            return FormatResult;
        }

        //public static string InsertTransaction1(Microsoft.Office.Interop.Excel.Workbook WB, string TranID, string RDate, string Insurance, string Salesby, string Serviceby, string location, string Support, string Rmonth, string strconn)
        //{
        //    SQLProcs sql = new SQLProcs();
        //    Microsoft.Office.Interop.Excel.Worksheet wks = (Microsoft.Office.Interop.Excel.Worksheet)WB.Worksheets[2];
        //    Microsoft.Office.Interop.Excel.Range lastCell = wks.Cells.SpecialCells(Microsoft.Office.Interop.Excel.XlCellType.xlCellTypeLastCell, Type.Missing);
        //    int lastrow = lastCell.Row; string PolicyNo = ""; int J = 0;
        //    string Checknull = ((Microsoft.Office.Interop.Excel.Range)wks.Cells[1, 13]).Value;
        //    string CheckFormat = ((Microsoft.Office.Interop.Excel.Range)wks.Cells[1, 3]).Value;
        //    string FormatResult = "Not OK";
        //    if (Checknull != null && Checknull.Contains("Insured Name"))
        //    {
        //        J = 1;
        //    }
        //    if (CheckFormat != null)
        //    {
        //        CheckFormat = CheckFormat.Replace("\n", "").TrimStart();
        //        if (CheckFormat.Contains("Office Code") || CheckFormat.Contains("OfficeCode"))
        //        {
        //            FormatResult = "OK";
        //        }
        //        else
        //        {
        //            FormatResult = "Not OK";
        //        }
        //    }
        //    if (FormatResult == "OK")
        //    {
        //        for (int i = 2; i <= lastrow; i++)
        //        {
        //            var Terrorism = ""; var Policy_Endorsement = ""; var InsuredType = ""; var Policy_Type = ""; var New_Renewal = ""; var Revenue_Amt = "";

        //            string InsuredName = ((Microsoft.Office.Interop.Excel.Range)wks.Cells[i, 14 - J]).Value;
        //            InsuredType = ((Microsoft.Office.Interop.Excel.Range)wks.Cells[i, 16 - J]).Value;
        //            if (InsuredType != null && InsuredType != "" && InsuredType != " ")
        //            {
        //                InsuredName = InsuredName.Replace("\n", "").TrimStart();
        //                string Pno = Convert.ToString(((Microsoft.Office.Interop.Excel.Range)wks.Cells[i, 4]).Value);
        //                //var Endorsementno = Convert.ToString(((Microsoft.Office.Interop.Excel.Range)wks.Cells[i, 15]).Value);
        //                if (Pno != null && Pno != "" && Pno != " ")
        //                {
        //                    PolicyNo = Pno.Replace("\n", "").TrimStart();
        //                }
        //                var Endo_Effective_Date = ((Microsoft.Office.Interop.Excel.Range)wks.Cells[i, 18 - J]).Value;
        //                var Effective_Date = ((Microsoft.Office.Interop.Excel.Range)wks.Cells[i, 18 - J]).Value;
        //                var END_Date = ((Microsoft.Office.Interop.Excel.Range)wks.Cells[i, 19 - J]).Value;
        //                int ENdolen = Convert.ToString(Endo_Effective_Date).Length;
        //                if (ENdolen > 11)
        //                {
        //                    Endo_Effective_Date = Endo_Effective_Date.ToString("dd/MM/yyyy");
        //                }
        //                int Efflen = Convert.ToString(Effective_Date).Length;
        //                if (Efflen > 11)
        //                {
        //                    Effective_Date = Effective_Date.ToString("dd/MM/yyyy");
        //                }
        //                int ENDlen = Convert.ToString(END_Date).Length;
        //                if (ENDlen > 11)
        //                {
        //                    END_Date = END_Date.ToString("dd/MM/yyyy");
        //                }
        //                var Premium_Amt = Convert.ToString(((Microsoft.Office.Interop.Excel.Range)wks.Cells[i, 21 - J]).Value).Replace(",", "").Replace("(", "").Replace(")", "").TrimStart();
        //                //Terrorism = Convert.ToString(((Microsoft.Office.Interop.Excel.Range)wks.Cells[i, 23]).Value).Replace(",", "").Replace("(", "").Replace(")", "").TrimStart();
        //                Revenue_Amt = Convert.ToString(((Microsoft.Office.Interop.Excel.Range)wks.Cells[i, 22 - J]).Value).Replace(",", "").Replace("(", "").Replace(")", "").Replace("-", "").TrimStart();
        //                var Revenue_Pcnt = Convert.ToString(((Microsoft.Office.Interop.Excel.Range)wks.Cells[i, 20 - J]).Text);
        //                Revenue_Pcnt = Revenue_Pcnt.Replace("\n", "").Replace("%", "").Replace(",", "").TrimStart();
        //                PolicyNo = Regex.Replace(PolicyNo, "[^a-zA-Z0-9_.]+", "", RegexOptions.Compiled);
        //                string endno = Convert.ToString(((Microsoft.Office.Interop.Excel.Range)wks.Cells[i, 8]).Value).Replace("\n", "").TrimStart();
        //                if (endno != null && endno != "" && endno != " " && endno != ":")
        //                {
        //                    Policy_Endorsement = "Endorsement";
        //                }
        //                else
        //                {
        //                    Policy_Endorsement = "Policy";
        //                }
        //                Policy_Type = Convert.ToString(((Microsoft.Office.Interop.Excel.Range)wks.Cells[i, 12]).Value);
        //                //if (InsuredName.ToUpper().Contains("LIMITED") || InsuredName.ToUpper().Contains("LTD"))
        //                if (InsuredType == "Organizational")
        //                {
        //                    InsuredType = "Corporate";
        //                }
        //                else
        //                {
        //                    InsuredType = "Retail";
        //                }

        //                if (Premium_Amt == "" || Premium_Amt == " " || Premium_Amt == null)
        //                {
        //                    Premium_Amt = 0;
        //                }
        //                if (Revenue_Amt == "" || Revenue_Amt == " " || Revenue_Amt == null)
        //                {
        //                    Revenue_Amt = "0";
        //                }
        //                if (Revenue_Pcnt == "" || Revenue_Pcnt == " " || Revenue_Pcnt == null)
        //                {
        //                    Revenue_Pcnt = "0";
        //                }
        //                if (Terrorism == "" || Terrorism == " " || Terrorism == null)
        //                {
        //                    Terrorism = "0";
        //                }
        //                sql.ExecuteSQLNonQuery(strconn, "SP_NewIndiaPDFTransaction",
        //                           new SqlParameter { ParameterName = "@Imode", Value = 1 },
        //                           new SqlParameter { ParameterName = "@RDate", Value = RDate },
        //                           new SqlParameter { ParameterName = "@Rmonth", Value = Rmonth },
        //                           new SqlParameter { ParameterName = "@ClientName", Value = InsuredName },
        //                           new SqlParameter { ParameterName = "@Itype", Value = InsuredType },
        //                           new SqlParameter { ParameterName = "@Policy_No", Value = PolicyNo },
        //                           new SqlParameter { ParameterName = "@Revenue_Pct", Value = Revenue_Pcnt },
        //                           new SqlParameter { ParameterName = "@Policy_Type", Value = Policy_Type },
        //                           new SqlParameter { ParameterName = "@Premium_Amt", Value = Premium_Amt },
        //                           new SqlParameter { ParameterName = "@Endo_Effective_Date", Value = Endo_Effective_Date },
        //                           new SqlParameter { ParameterName = "@Effective_Date", Value = Effective_Date },
        //                           new SqlParameter { ParameterName = "@END_Date", Value = END_Date },
        //                           new SqlParameter { ParameterName = "@TranID", Value = TranID },
        //                           new SqlParameter { ParameterName = "@Revenue_Amt", Value = Revenue_Amt },
        //                           new SqlParameter { ParameterName = "@Terrorism", Value = Terrorism },
        //                           new SqlParameter { ParameterName = "@Insurance", Value = Insurance },
        //                           new SqlParameter { ParameterName = "@Salesby", Value = Salesby },
        //                           new SqlParameter { ParameterName = "@Serviceby", Value = Serviceby },
        //                           new SqlParameter { ParameterName = "@location", Value = location },
        //                           new SqlParameter { ParameterName = "@Support", Value = Support },
        //                           new SqlParameter { ParameterName = "@Policy_Endorsement", Value = Policy_Endorsement },
        //                           new SqlParameter { ParameterName = "@RFormat", Value = "F1" },
        //                           new SqlParameter { ParameterName = "@InvNo", Value = "NIP1" },
        //                           new SqlParameter { ParameterName = "@ReportId", Value = "NIP1" },
        //                           new SqlParameter { ParameterName = "@DocName", Value = "New India Assurance Company Ltd." }
        //                           );
        //            }
        //        }
        //    }
        //    return FormatResult;
        //}

        //public static string InsertTransaction2(Microsoft.Office.Interop.Excel.Workbook WB, string TranID, string RDate, string Insurance, string Salesby, string Serviceby, string location, string Support, string Rmonth, string strconn)
        //{
        //    //New India Assurance Company Limited.
        //    SQLProcs sql = new SQLProcs();
        //    Microsoft.Office.Interop.Excel.Worksheet wks = (Microsoft.Office.Interop.Excel.Worksheet)WB.Worksheets[2];
        //    Microsoft.Office.Interop.Excel.Range lastCell = wks.Cells.SpecialCells(Microsoft.Office.Interop.Excel.XlCellType.xlCellTypeLastCell, Type.Missing);
        //    int lastrow = lastCell.Row;
        //    string CheckFormat = ((Microsoft.Office.Interop.Excel.Range)wks.Cells[1, 3]).Value;
        //    string FormatResult = "Not OK";
        //    if (CheckFormat != null)
        //    {
        //        CheckFormat = CheckFormat.Replace("\n", "").TrimStart();
        //        if (CheckFormat.Contains("Dept"))
        //        {
        //            FormatResult = "OK";
        //        }
        //        else
        //        {
        //            FormatResult = "Not OK";
        //        }
        //    }
        //    if (FormatResult == "OK")
        //    {
        //        for (int i = 2; i <= lastrow; i++)
        //        {
        //            var Terrorism = ""; var Policy_Endorsement = ""; var InsuredType = ""; var Policy_Type = ""; var New_Renewal = ""; var Revenue_Amt = "";
        //            string InsuredName = ((Microsoft.Office.Interop.Excel.Range)wks.Cells[i, 9]).Value;
        //            if (InsuredName != null && InsuredName != "" && InsuredName != " ")
        //            {
        //                InsuredName = InsuredName.Replace("\n", "").TrimStart();
        //                InsuredType = ((Microsoft.Office.Interop.Excel.Range)wks.Cells[i, 12]).Value;
        //                string PolicyNo = Convert.ToString(((Microsoft.Office.Interop.Excel.Range)wks.Cells[i, 4]).Value).Replace("\n", "").TrimStart();
        //                //var Endorsementno = Convert.ToString(((Microsoft.Office.Interop.Excel.Range)wks.Cells[i, 15]).Value);

        //                var Endo_Effective_Date = ((Microsoft.Office.Interop.Excel.Range)wks.Cells[i, 15]).Value;
        //                var Effective_Date = ((Microsoft.Office.Interop.Excel.Range)wks.Cells[i, 15]).Value;
        //                var END_Date = ((Microsoft.Office.Interop.Excel.Range)wks.Cells[i, 17]).Value;
        //                int ENdolen = Convert.ToString(Endo_Effective_Date).Length;
        //                if (ENdolen > 11)
        //                {
        //                    Endo_Effective_Date = Endo_Effective_Date.ToString("dd/MM/yyyy");
        //                }
        //                int Efflen = Convert.ToString(Effective_Date).Length;
        //                if (Efflen > 11)
        //                {
        //                    Effective_Date = Effective_Date.ToString("dd/MM/yyyy");
        //                }
        //                int ENDlen = Convert.ToString(END_Date).Length;
        //                if (ENDlen > 11)
        //                {
        //                    END_Date = END_Date.ToString("dd/MM/yyyy");
        //                }
        //                var Premium_Amt = Convert.ToString(((Microsoft.Office.Interop.Excel.Range)wks.Cells[i, 20]).Value).Replace(",", "").Replace("(", "").Replace(")", "").TrimStart();
        //                //Terrorism = Convert.ToString(((Microsoft.Office.Interop.Excel.Range)wks.Cells[i, 23]).Value).Replace(",", "").Replace("(", "").Replace(")", "").TrimStart();
        //                Revenue_Amt = Convert.ToString(((Microsoft.Office.Interop.Excel.Range)wks.Cells[i, 21]).Value).Replace(",", "").Replace("(", "").Replace(")", "").Replace("-", "").TrimStart();
        //                var Revenue_Pcnt = Convert.ToString(((Microsoft.Office.Interop.Excel.Range)wks.Cells[i, 18]).Text);
        //                Revenue_Pcnt = Revenue_Pcnt.Replace("\n", "").Replace("%", "").Replace(",", "").TrimStart();
        //                if (PolicyNo != null && PolicyNo != "" && PolicyNo != " ")
        //                {
        //                    PolicyNo = Regex.Replace(PolicyNo, "[^a-zA-Z0-9_.]+", "", RegexOptions.Compiled);
        //                }
        //                Policy_Type = Convert.ToString(((Microsoft.Office.Interop.Excel.Range)wks.Cells[i, 3]).Value);
        //                if (InsuredType == "Organizational")
        //                {
        //                    InsuredType = "Corporate";
        //                }
        //                else
        //                {
        //                    InsuredType = "Retail";
        //                }
        //                string endno = Convert.ToString(((Microsoft.Office.Interop.Excel.Range)wks.Cells[i, 7]).Value).Replace("\n", "").TrimStart();
        //                if (endno != null && endno != "" && endno != " " && endno != ":")
        //                {
        //                    Policy_Endorsement = "Endorsement";
        //                    PolicyNo = PolicyNo + " " + Regex.Replace(endno, "[^a-zA-Z0-9_.]+", "", RegexOptions.Compiled);
        //                }
        //                else
        //                {
        //                    Policy_Endorsement = "Policy";
        //                }
        //                if (Premium_Amt == "" || Premium_Amt == " " || Premium_Amt == null)
        //                {
        //                    Premium_Amt = 0;
        //                }
        //                if (Revenue_Amt == "" || Revenue_Amt == " " || Revenue_Amt == null)
        //                {
        //                    Revenue_Amt = "0";
        //                }
        //                if (Revenue_Pcnt == "" || Revenue_Pcnt == " " || Revenue_Pcnt == null)
        //                {
        //                    Revenue_Pcnt = "0";
        //                }
        //                if (Terrorism == "" || Terrorism == " " || Terrorism == null)
        //                {
        //                    Terrorism = "0";
        //                }
        //                sql.ExecuteSQLNonQuery(strconn, "SP_NewIndiaPDFTransaction",
        //                           new SqlParameter { ParameterName = "@Imode", Value = 1 },
        //                           new SqlParameter { ParameterName = "@RDate", Value = RDate },
        //                           new SqlParameter { ParameterName = "@Rmonth", Value = Rmonth },
        //                           new SqlParameter { ParameterName = "@ClientName", Value = InsuredName },
        //                           new SqlParameter { ParameterName = "@Itype", Value = InsuredType },
        //                           new SqlParameter { ParameterName = "@Policy_No", Value = PolicyNo },
        //                           new SqlParameter { ParameterName = "@Revenue_Pct", Value = Revenue_Pcnt },
        //                           new SqlParameter { ParameterName = "@Policy_Type", Value = Policy_Type },
        //                           new SqlParameter { ParameterName = "@Premium_Amt", Value = Premium_Amt },
        //                           new SqlParameter { ParameterName = "@Endo_Effective_Date", Value = Endo_Effective_Date },
        //                           new SqlParameter { ParameterName = "@Effective_Date", Value = Effective_Date },
        //                           new SqlParameter { ParameterName = "@END_Date", Value = END_Date },
        //                           new SqlParameter { ParameterName = "@TranID", Value = TranID },
        //                           new SqlParameter { ParameterName = "@Revenue_Amt", Value = Revenue_Amt },
        //                           new SqlParameter { ParameterName = "@Terrorism", Value = Terrorism },
        //                           new SqlParameter { ParameterName = "@Insurance", Value = Insurance },
        //                           new SqlParameter { ParameterName = "@Salesby", Value = Salesby },
        //                           new SqlParameter { ParameterName = "@Serviceby", Value = Serviceby },
        //                           new SqlParameter { ParameterName = "@location", Value = location },
        //                           new SqlParameter { ParameterName = "@Support", Value = Support },
        //                           new SqlParameter { ParameterName = "@Policy_Endorsement", Value = Policy_Endorsement },
        //                           new SqlParameter { ParameterName = "@RFormat", Value = "F2" },
        //                           new SqlParameter { ParameterName = "@InvNo", Value = "NIP2" },
        //                           new SqlParameter { ParameterName = "@ReportId", Value = "NIP2" },
        //                           new SqlParameter { ParameterName = "@DocName", Value = "New India Assurance Company Ltd." }
        //                           );
        //            }
        //        }
        //    }
        //    return FormatResult;
        //}
        //public static string InsertTransaction3(Microsoft.Office.Interop.Excel.Workbook WB, string TranID, string RDate, string Insurance, string Salesby, string Serviceby, string location, string Support, string Rmonth, string strconn)
        //{
        //    //New India Assurance Company Limited.
        //    SQLProcs sql = new SQLProcs();
        //    Microsoft.Office.Interop.Excel.Worksheet wks = (Microsoft.Office.Interop.Excel.Worksheet)WB.Worksheets[2];
        //    Microsoft.Office.Interop.Excel.Range lastCell = wks.Cells.SpecialCells(Microsoft.Office.Interop.Excel.XlCellType.xlCellTypeLastCell, Type.Missing);
        //    int lastrow = lastCell.Row; string Policy_Type = ""; var PType = "";
        //    string CheckFormat = ((Microsoft.Office.Interop.Excel.Range)wks.Cells[1, 1]).Value;
        //    string FormatResult = "Not OK";
        //    if (CheckFormat != null)
        //    {
        //        CheckFormat = CheckFormat.Replace("\n", "").TrimStart();
        //        if (CheckFormat.Contains("Department"))
        //        {
        //            FormatResult = "OK";
        //        }
        //        else
        //        {
        //            FormatResult = "Not OK";
        //        }
        //    }
        //    if (FormatResult == "OK")
        //    {
        //        for (int i = 2; i <= lastrow; i++)
        //        {
        //            var Terrorism = ""; var Policy_Endorsement = ""; var InsuredType = ""; var Revenue_Pcnt = ""; var Revenue_Amt = ""; var PolicyNo = "";
        //            string InsuredName = ((Microsoft.Office.Interop.Excel.Range)wks.Cells[i, 6]).Value;
        //            if (InsuredName != null && InsuredName != "" && InsuredName != " ")
        //            {
        //                InsuredName = InsuredName.Replace("\n", "").TrimStart();
        //                InsuredType = ((Microsoft.Office.Interop.Excel.Range)wks.Cells[i, 7]).Value;
        //                var PNo = Convert.ToString(((Microsoft.Office.Interop.Excel.Range)wks.Cells[i, 2]).Value);
        //                if (PNo != null && PNo != "" && PNo != " ")
        //                {
        //                    PType = PNo.Replace("\n", "").TrimStart();
        //                    PType = Regex.Replace(PType, "[^a-zA-Z0-9_.]+", "", RegexOptions.Compiled);
        //                }
        //                var endno = Convert.ToString(((Microsoft.Office.Interop.Excel.Range)wks.Cells[i, 5]).Value);
        //                if (endno != null && endno != "" && endno != " " && endno != ":")
        //                {
        //                    Policy_Endorsement = "Endorsement";
        //                    PolicyNo = PType + " " + Regex.Replace(endno, "[^a-zA-Z0-9_.]+", "", RegexOptions.Compiled);
        //                }
        //                else
        //                {
        //                    Policy_Endorsement = "Policy";
        //                    PolicyNo = PType;
        //                }
        //                if (InsuredType == "Organizational")
        //                {
        //                    InsuredType = "Corporate";
        //                }
        //                else
        //                {
        //                    InsuredType = "Retail";
        //                }
        //                var Endo_Effective_Date = ((Microsoft.Office.Interop.Excel.Range)wks.Cells[i, 8]).Value;
        //                var Effective_Date = ((Microsoft.Office.Interop.Excel.Range)wks.Cells[i, 8]).Value;
        //                var END_Date = ((Microsoft.Office.Interop.Excel.Range)wks.Cells[i, 9]).Value;
        //                int ENdolen = Convert.ToString(Endo_Effective_Date).Length;
        //                if (ENdolen > 11)
        //                {
        //                    Endo_Effective_Date = Endo_Effective_Date.ToString("dd/MM/yyyy");
        //                }
        //                int Efflen = Convert.ToString(Effective_Date).Length;
        //                if (Efflen > 11)
        //                {
        //                    Effective_Date = Effective_Date.ToString("dd/MM/yyyy");
        //                }
        //                int ENDlen = Convert.ToString(END_Date).Length;
        //                if (ENDlen > 11)
        //                {
        //                    END_Date = END_Date.ToString("dd/MM/yyyy");
        //                }
        //                var Premium_Amt = Convert.ToString(((Microsoft.Office.Interop.Excel.Range)wks.Cells[i, 10]).Value).Replace(",", "").Replace("(", "").Replace(")", "").TrimStart();
        //                //Terrorism = Convert.ToString(((Microsoft.Office.Interop.Excel.Range)wks.Cells[i, 23]).Value).Replace(",", "").Replace("(", "").Replace(")", "").TrimStart();
        //                Revenue_Amt = Convert.ToString(((Microsoft.Office.Interop.Excel.Range)wks.Cells[i, 11]).Value).Replace(",", "").Replace("(", "").Replace(")", "").Replace("-", "").TrimStart();
        //                //var Revenue_Pcnt = Convert.ToString(((Microsoft.Office.Interop.Excel.Range)wks.Cells[i, 21]).Text);
        //                //Revenue_Pcnt = Revenue_Pcnt.Replace("\n", "").Replace("%", "").Replace(",", "").TrimStart();

        //                string dept = Convert.ToString(((Microsoft.Office.Interop.Excel.Range)wks.Cells[i, 1]).Value);
        //                if (dept != null && dept != "" && dept != " ")
        //                {
        //                    Policy_Type = Regex.Match(dept, @"\d+").Value;
        //                }

        //                if (Premium_Amt == "" || Premium_Amt == " " || Premium_Amt == null)
        //                {
        //                    Premium_Amt = 0;
        //                }
        //                if (Revenue_Amt == "" || Revenue_Amt == " " || Revenue_Amt == null)
        //                {
        //                    Revenue_Amt = "0";
        //                }
        //                if (Revenue_Pcnt == "" || Revenue_Pcnt == " " || Revenue_Pcnt == null)
        //                {
        //                    Revenue_Pcnt = "0";
        //                }
        //                if (Terrorism == "" || Terrorism == " " || Terrorism == null)
        //                {
        //                    Terrorism = "0";
        //                }
        //                sql.ExecuteSQLNonQuery(strconn, "SP_NewIndiaPDFTransaction",
        //                           new SqlParameter { ParameterName = "@Imode", Value = 1 },
        //                           new SqlParameter { ParameterName = "@RDate", Value = RDate },
        //                           new SqlParameter { ParameterName = "@Rmonth", Value = Rmonth },
        //                           new SqlParameter { ParameterName = "@ClientName", Value = InsuredName },
        //                           new SqlParameter { ParameterName = "@Itype", Value = InsuredType },
        //                           new SqlParameter { ParameterName = "@Policy_No", Value = PolicyNo },
        //                           new SqlParameter { ParameterName = "@Revenue_Pct", Value = Revenue_Pcnt },
        //                           new SqlParameter { ParameterName = "@Policy_Type", Value = Policy_Type },
        //                           new SqlParameter { ParameterName = "@Premium_Amt", Value = Premium_Amt },
        //                           new SqlParameter { ParameterName = "@Endo_Effective_Date", Value = Endo_Effective_Date },
        //                           new SqlParameter { ParameterName = "@Effective_Date", Value = Effective_Date },
        //                           new SqlParameter { ParameterName = "@END_Date", Value = END_Date },
        //                           new SqlParameter { ParameterName = "@TranID", Value = TranID },
        //                           new SqlParameter { ParameterName = "@Revenue_Amt", Value = Revenue_Amt },
        //                           new SqlParameter { ParameterName = "@Terrorism", Value = Terrorism },
        //                           new SqlParameter { ParameterName = "@Insurance", Value = Insurance },
        //                           new SqlParameter { ParameterName = "@Salesby", Value = Salesby },
        //                           new SqlParameter { ParameterName = "@Serviceby", Value = Serviceby },
        //                           new SqlParameter { ParameterName = "@location", Value = location },
        //                           new SqlParameter { ParameterName = "@Support", Value = Support },
        //                           new SqlParameter { ParameterName = "@Policy_Endorsement", Value = Policy_Endorsement },
        //                           new SqlParameter { ParameterName = "@RFormat", Value = "F3" },
        //                           new SqlParameter { ParameterName = "@InvNo", Value = "NIP3" },
        //                           new SqlParameter { ParameterName = "@ReportId", Value = "NIP3" },
        //                           new SqlParameter { ParameterName = "@DocName", Value = "New India Assurance Company Ltd." }
        //                           );
        //            }
        //        }
        //    }
        //    return FormatResult;
        //}
    }
}
