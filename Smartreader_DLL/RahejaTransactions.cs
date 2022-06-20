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
namespace Smartreader_DLL
{
    public class RahejaTransactions
    {
        public static void InsertTransaction(Microsoft.Office.Interop.Excel.Workbook WB, string TranID, string RDate, string Insurance, string Salesby, string Serviceby, string location, string Support, string Rmonth, string strconn)
        {
            //Raheja QBE General Insurance Co.Ltd.
            SQLProcs sql = new SQLProcs();
            Microsoft.Office.Interop.Excel.Worksheet wks = (Microsoft.Office.Interop.Excel.Worksheet)WB.Worksheets[1];
            Microsoft.Office.Interop.Excel.Range lastCell = wks.Cells.SpecialCells(Microsoft.Office.Interop.Excel.XlCellType.xlCellTypeLastCell, Type.Missing);
            int lastrow = lastCell.Row; 
            for (int i = 2; i <= lastrow; i++)
            {
                var Terrorism = ""; var Policy_Endorsement = ""; var InsuredType = ""; var Policy_Type = ""; var New_Renewal = ""; var Premium_Amt = "";
                string InsuredName = "";
                string In1 = ((Microsoft.Office.Interop.Excel.Range)wks.Cells[i, 4]).Value;
                string In2 = ((Microsoft.Office.Interop.Excel.Range)wks.Cells[i, 5]).Value;
                if ((In1 != null && In1 != "" && In1 != " ") || (In2 != null && In2 != "" && In2 != " "))
                {
                    if (In2 != null && In1 != null)
                    {
                        InsuredName = In1.Replace("\n", "").TrimStart() + " " + In2.TrimStart();
                    }
                    else if (In2 != null && In1 == null)
                    {
                        InsuredName = In2.Replace("\n", "").TrimStart();
                    }
                    else
                    {
                        InsuredName = In1.Replace("\n", "").TrimStart();
                    }
                    var BRcode = Convert.ToString(((Microsoft.Office.Interop.Excel.Range)wks.Cells[i, 1]).Value);
                    string EndosNo = Convert.ToString(((Microsoft.Office.Interop.Excel.Range)wks.Cells[i, 7]).Value);
                    bool result = EndosNo.Any(x => !char.IsLetter(x));
                    string PolicyNo = Convert.ToString(((Microsoft.Office.Interop.Excel.Range)wks.Cells[i, 6]).Value);
                    if (PolicyNo != null && PolicyNo != "" && PolicyNo != " ")
                    {
                        if (result == true)
                        {
                            if (Convert.ToDecimal(EndosNo) > 1)
                            {
                                if (EndosNo.Length == 1)
                                {
                                    EndosNo = "0000" + Convert.ToString(Convert.ToDecimal(EndosNo) - 1);
                                }
                                else if (EndosNo.Length == 2)
                                {
                                    EndosNo = "000" + Convert.ToString(Convert.ToDecimal(EndosNo) - 1);
                                }
                                else if (EndosNo.Length == 3)
                                {
                                    EndosNo = "00" + Convert.ToString(Convert.ToDecimal(EndosNo) - 1);
                                }
                                else if (EndosNo.Length == 3)
                                {
                                    EndosNo = "0" + Convert.ToString(Convert.ToDecimal(EndosNo) - 1);
                                }
                                PolicyNo = "0" + BRcode + PolicyNo + EndosNo;
                            }
                            else
                            {
                                PolicyNo = "0" + BRcode + PolicyNo + "00000";
                            }
                        }
                        else
                        {
                            PolicyNo = "0" + BRcode + PolicyNo + EndosNo;
                        }
                    }

                    var Client_N_E = ((Microsoft.Office.Interop.Excel.Range)wks.Cells[i, 3]).Value.Replace("\n", "").TrimStart();
                    //Policy_Endorsement = ((Microsoft.Office.Interop.Excel.Range)wks.Cells[i, 32]).Value.Replace("\n", "").TrimStart();
                    if (Client_N_E == "New Business" || Client_N_E == "New")
                    {
                        Client_N_E = "New Client";
                        New_Renewal = "New Policy";
                    }
                    else
                    {
                        Client_N_E = "Existing Client";
                        New_Renewal = "Renewal Policy";
                    }
                    var Endo_Effective_Date = ((Microsoft.Office.Interop.Excel.Range)wks.Cells[i, 27]).Value;
                    var Effective_Date = ((Microsoft.Office.Interop.Excel.Range)wks.Cells[i, 13]).Value;
                    var END_Date = ((Microsoft.Office.Interop.Excel.Range)wks.Cells[i, 14]).Value;
                    var offlocation = Convert.ToString(((Microsoft.Office.Interop.Excel.Range)wks.Cells[i, 2]).Value);
                    //END_Date = END_Date.Substring(0, END_Date.LastIndexOf(" "));
                    if (Endo_Effective_Date != null)
                    {
                        int ENdolen = Convert.ToString(Endo_Effective_Date).Length;
                        if (ENdolen > 11)
                        {
                            Endo_Effective_Date = Endo_Effective_Date.ToString("dd/MM/yyyy");
                        }
                        else if (ENdolen <= 8)
                        {
                            //DateTime date = Convert.ToDateTime(Endo_Effective_Date, CultureInfo.InvariantCulture);
                            //var d1 = date.ToString("yyyy-MM-dd");
                            var year = Convert.ToString(Endo_Effective_Date).Substring(0, 4);
                            string month = Convert.ToString(Endo_Effective_Date).Substring(4);
                            month = month.Substring(0, 2);
                            string date = Convert.ToString(Endo_Effective_Date).Substring(6);
                            Endo_Effective_Date = Convert.ToString(date + "/" + month + "/" + year);
                        }
                    }
                    int Efflen = Convert.ToString(Effective_Date).Length;
                    if (Efflen == 16)
                    {
                        Effective_Date = Effective_Date.Substring(0, Effective_Date.LastIndexOf(" "));
                    }
                    else if (Efflen > 16)
                    {
                        Effective_Date = Effective_Date.ToString("dd/MM/yyyy");
                    }
                    int ENDlen = Convert.ToString(END_Date).Length;
                    if (ENDlen == 16)
                    {
                        END_Date = END_Date.Substring(0, END_Date.LastIndexOf(" "));
                    }
                    else if (ENDlen > 16)
                    {
                        END_Date = END_Date.ToString("dd/MM/yyyy");
                    }

                    var PA = Convert.ToString(((Microsoft.Office.Interop.Excel.Range)wks.Cells[i, 17]).Value);
                    var PA1 = Convert.ToString(((Microsoft.Office.Interop.Excel.Range)wks.Cells[i, 28]).Value);
                    var Revenue_Amt = Convert.ToString(((Microsoft.Office.Interop.Excel.Range)wks.Cells[i, 18]).Value);
                    //var Revenue_Pcnt = Convert.ToString(((Microsoft.Office.Interop.Excel.Range)wks.Cells[i, 19]).Text);
                    //Revenue_Pcnt = Revenue_Pcnt.Replace("\n", "").Replace("%", "").Replace(",", "").TrimStart();
                    Terrorism = Convert.ToString(((Microsoft.Office.Interop.Excel.Range)wks.Cells[i, 29]).Value);
                    var Reward_Pcnt = Convert.ToString(((Microsoft.Office.Interop.Excel.Range)wks.Cells[i, 50]).Value);
                    var Reward_Amt = Convert.ToString(((Microsoft.Office.Interop.Excel.Range)wks.Cells[i, 51]).Value);
                    Policy_Type = Convert.ToString(((Microsoft.Office.Interop.Excel.Range)wks.Cells[i, 10]).Value);

                    if (InsuredName.ToUpper().Contains("LIMITED") || InsuredName.ToUpper().Contains("LTD") || InsuredName.ToUpper().Contains("INSURANCE"))
                    {
                        InsuredType = "Corporate";
                    }
                    else
                    {
                        InsuredType = "Retail";
                    }
                    if (PA == "" || PA == " " || PA == null)
                    {
                        PA = "0";
                    }
                    else
                    {
                        PA = PA.Replace(",", "").TrimStart();
                    }
                    if (PA1 == "" || PA1 == " " || PA1 == null)
                    {
                        PA1 = "0";
                    }
                    else
                    {
                        PA1 = PA1.Replace(",", "").TrimStart();
                    }
                    if (Policy_Type.ToUpper().Contains("MVA"))
                    {
                        Premium_Amt = PA1;
                    }
                    else
                    {
                        Premium_Amt = PA;
                    }
                    if (Premium_Amt == "" || Premium_Amt == " " || Premium_Amt == null)
                    {
                        Premium_Amt = "0";
                    }
                    if (Revenue_Amt == "" || Revenue_Amt == " " || Revenue_Amt == null)
                    {
                        Revenue_Amt = 0;
                    }
                    if (Reward_Pcnt == "" || Reward_Pcnt == " " || Reward_Pcnt == null)
                    {
                        Reward_Pcnt = "0";
                    }
                    if (Reward_Amt == "" || Reward_Amt == " " || Reward_Amt == null)
                    {
                        Reward_Amt = "0";
                    }
                    else
                    {
                        Reward_Amt = Reward_Amt.Replace(",", "").TrimStart();
                    }
                    if (Terrorism == "" || Terrorism == " " || Terrorism == null)
                    {
                        Terrorism = "0";
                    }
                    sql.ExecuteSQLNonQuery(strconn, "SP_RahejaTransactions",
                               new SqlParameter { ParameterName = "@Imode", Value = 1 },
                               new SqlParameter { ParameterName = "@RDate", Value = RDate },
                               new SqlParameter { ParameterName = "@Rmonth", Value = Rmonth },
                               new SqlParameter { ParameterName = "@ClientName", Value = InsuredName },
                               new SqlParameter { ParameterName = "@Itype", Value = InsuredType },
                               new SqlParameter { ParameterName = "@Policy_No", Value = PolicyNo },
                               new SqlParameter { ParameterName = "@Reward_Pcnt", Value = Reward_Pcnt },
                               new SqlParameter { ParameterName = "@Reward_Amt", Value = Reward_Amt },
                               new SqlParameter { ParameterName = "@Client_N_E", Value = Client_N_E },
                               new SqlParameter { ParameterName = "@New_Renewal", Value = New_Renewal },
                               new SqlParameter { ParameterName = "@Endo_Effective_Date", Value = Endo_Effective_Date },
                               new SqlParameter { ParameterName = "@Effective_Date", Value = Effective_Date },
                               new SqlParameter { ParameterName = "@END_Date", Value = END_Date },
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
                               new SqlParameter { ParameterName = "@InvNo", Value = "RGIX" },
                               new SqlParameter { ParameterName = "@ReportId", Value = "RGIX" },
                               new SqlParameter { ParameterName = "@DocName", Value = "Reliance General Insurance Co. Ltd." }
                               );
                }
            }
        }
    }
}
