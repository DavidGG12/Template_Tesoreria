using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Template_Tesoreria.Models
{
    public class TblTesoreria_Model
    {
        public string Bank_Account_Number { get; set; }
        public string Transaction_Code { get; set; }
        public string Bank_Account_Currency { get; set; }
        public string Booking_Date { get; set; }
        public string Value_Date { get; set; }
        public string Credit { get; set; }
        public string Debit { get; set; }
        public string Check_Number { get; set; }
        public string Addenda_Text { get; set; }
        public string Account_Servicer_Reference { get; set; }
        public string Customer_Reference { get; set; }
        public string Clearing_System_Reference { get; set; }
        public string Contract_Identifier { get; set; }
        public string Instruction_Identifier { get; set; }
        public string End_To_End_Identifier { get; set; }
        public string Servicer_Status { get; set; }
        public string Commision_Waiver_Indicator_Flag { get; set; }
        public string Reversal_Indicator_Flag { get; set; }
        public string Structured_Payment_Reference { get; set; }
        public string Reconciliation_Reference { get; set; }
        public string Message_Identifier { get; set; }
        public string Payment_Information_Identifier { get; set; }
        public string Open_Balance { get; set; }
        public string Close_Balance { get; set; }
        
        public bool parseDate(List<TblTesoreria_Model> data)
        {
            try
            {
                foreach(var row in data)
                {
                    string[] formats = {
                        "ddMMyyyy", "dMMyyyy", "yyyyMMdd", "yyyy-MM-dd",
                        "yyyy/MM/dd", "dd/MM/yyyy", "MM/dd/yyyy", "ddMMyy",
                        "yyyy-MM-dd HH:mm:ss", "yyyy/MM/dd HH:mm:ss"
                    };
                    DateTime dateParse;

                    if (row.Booking_Date.Length < 8)
                        row.Booking_Date = $"{new string('0', 1)}{row.Booking_Date}";
                    if(row.Value_Date.Length < 8)
                        row.Value_Date = $"{new string('0', 1)}{row.Value_Date}";

                    if (DateTime.TryParseExact(
                        row.Booking_Date,
                        formats,
                        System.Globalization.CultureInfo.InvariantCulture,
                        System.Globalization.DateTimeStyles.None,
                        out dateParse
                    ))
                    {
                        row.Booking_Date = dateParse.ToString("dd/MM/yyyy");
                    }

                    if (DateTime.TryParseExact(
                        row.Value_Date,
                        formats,
                        System.Globalization.CultureInfo.InvariantCulture,
                        System.Globalization.DateTimeStyles.None,
                        out dateParse
                    ))
                    {
                        row.Value_Date = dateParse.ToString("dd/MM/yyyy");
                    }
                }

                return true;
            }
            catch(Exception ex)
            {
                return false;
            }
        }
    }

}
