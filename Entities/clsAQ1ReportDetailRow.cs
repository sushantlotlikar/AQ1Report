using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AQ1Report.Entities
{
    class clsAQ1ReportDetailRow
    {
        public clsAQ1ReportHeaderRow HeaderRow = null;

        public UInt32? ExcelRowNo = null;

        public string A_row1_Folio = "";
        public string B_row1_CCN = "";
        public string C_row1_CCNStat = "";
        public string D_row1_CCNLink = "";
        public string E_row1_CCN10 = "";
        public string F_row1_MtrStat = "";
        public string G_row1_GAP = "";
        public string H_row1_CurrentDate = "";
        public string H_row2_CurrentRdg = "";
        public string I_row1_PreviousDate = "";
        public string I_row2_PreviousRdg = "";
        public string J_row1_CutRemvDt = "";
        public string J_row2_CutRemv = "";
        public string K_row1_RstrRplcDt = "";
        public string K_row2_RstrRplc = "";
        public string L_row1_CsmpByMeter = "";
        public string M_row1_Days = "";
        public string N_row1_CsmpBilled = "";
        public string O_row1_Water = "";
        public string P_row1_Sewarage = "";
        public string Q_row1_Rent = "";
        public string R_row1_BillAmt = "";
        public string S_row1_Additional = "";
        public string T_row1_Flat = "";
        public string U_row1_Flag = "";
        public string V_row1_GroupCode = "";
        public string W_row1_RateCharge = "";
        public string X_row1_CutDate = "";
        public string X_row2_UNKNOWN_FIELD_NAME = "";
        public string Y_row1_Reason = "";

        public clsAQ1ReportDetailRow(clsAQ1ReportHeaderRow headerRow, UInt32 excelRowNo)
        {
            HeaderRow = headerRow;
            ExcelRowNo = excelRowNo;
        }

        public void WriteToConsole()
        {
            Console.Write("WorkSheet No: " + this.HeaderRow.PageNo.ToString() + " | Row No: " + this.ExcelRowNo.ToString() + " | ");

            Console.WriteLine("Folio: " + this.A_row1_Folio);

            Console.Write("Folio: " + this.A_row1_Folio + ", ");
            Console.Write("CCN: " + this.B_row1_CCN + ", ");
            Console.Write("CCN Stat: " + this.C_row1_CCNStat + ", ");
            Console.Write("CCN Link: " + this.D_row1_CCNLink + ", ");
            Console.Write("CCN 10: " + this.E_row1_CCN10 + ", ");
            Console.Write("Mtr: " + this.F_row1_MtrStat + ", ");
            Console.Write("GAP: " + this.G_row1_GAP + ", ");
            Console.Write("Current Date: " + this.H_row1_CurrentDate + ", ");
            Console.Write("Current Rdg: " + this.H_row2_CurrentRdg + ", ");
            Console.Write("Previous Date: " + this.I_row1_PreviousDate + ", ");
            Console.Write("Previous Rdg: " + this.I_row2_PreviousRdg + ", ");
            Console.Write("Cut Remv Dt: " + this.J_row1_CutRemvDt + ", ");
            Console.Write("Cut Remv: " + this.J_row2_CutRemv + ", ");
            Console.Write("Rstr Rplc Dt: " + this.K_row1_RstrRplcDt + ", ");
            Console.Write("Rstr Rplc: " + this.K_row2_RstrRplc + ", ");
            Console.Write("Csmp By Meter: " + this.L_row1_CsmpByMeter + ", ");
            Console.Write("Days: " + this.M_row1_Days + ", ");
            Console.Write("Csmp Billed: " + this.N_row1_CsmpBilled + ", ");
            Console.Write("Water: " + this.O_row1_Water + ", ");
            Console.Write("Sewarage: " + this.P_row1_Sewarage + ", ");
            Console.Write("Rent: " + this.Q_row1_Rent + ", ");
            Console.Write("Bill Amt: " + this.R_row1_BillAmt + ", ");
            Console.Write("Additional: " + this.S_row1_Additional + ", ");
            Console.Write("Flat: " + this.T_row1_Flat + ", ");
            Console.Write("Flag: " + this.U_row1_Flag + ", ");
            Console.Write("Group Code: " + this.V_row1_GroupCode + ", ");
            Console.Write("Rate Charge: " + this.W_row1_RateCharge + ", ");
            Console.Write("Cut Date: " + this.X_row1_CutDate + ", ");
            Console.Write("Cut Date-Row2: " + this.X_row2_UNKNOWN_FIELD_NAME + ", ");
            Console.WriteLine("Reason: " + this.Y_row1_Reason);
            Console.WriteLine("");
        }
    }
}
