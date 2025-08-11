using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using Template_Tesoreria.Helpers.MangementLog;
using Template_Tesoreria.Models;
using Excel = Microsoft.Office.Interop.Excel;


namespace Template_Tesoreria.Helpers.Files
{
    public class ManagementExcel
    {
        private string _path;
        private string bank;
        private int _rowHeader;
        private int _rowBalances;
        private FileInfo _file;
        private Log _log;
        private ExecutionTimer _et;
        private List<BankPrefix_Model> _preBank;
        public ManagementExcel(string pathExcel, string bank) 
        {
            this._rowHeader = 5;
            this._rowBalances = 5;
            this.bank = bank;
            this._path = pathExcel;
            this._file = new FileInfo(this._path);
            this._log = new Log();
            this._et = new ExecutionTimer();
            this._preBank = new List<BankPrefix_Model>()
            {
                new BankPrefix_Model(){ NombreBanco = "Inbursa",       Prefijo = "INB"  },
                new BankPrefix_Model(){ NombreBanco = "HSBC",          Prefijo = "HSBC" },
                new BankPrefix_Model(){ NombreBanco = "Bancomer",      Prefijo = "BBVA" },
                new BankPrefix_Model(){ NombreBanco = "Scotiabank",    Prefijo = "SCOT" },
                new BankPrefix_Model(){ NombreBanco = "Citibanamex",   Prefijo = "CITI" },
                new BankPrefix_Model(){ NombreBanco = "Santander",     Prefijo = "SANT" },
                new BankPrefix_Model(){ NombreBanco = "Banorte",       Prefijo = "BAN"  }
            };
            ExcelPackage.License.SetNonCommercialOrganization("Grupo Sanborns");
        }

        public void setFilePath(string pathExcel)
        {
            this._path = pathExcel;
            this._file = new FileInfo(this._path);
        }

        public string cleanSheets(string sheet)
        {
            this._et.startExecution();
            this._log.writeLog($"(INFO) LIMPIEZA DE LA HOJA {sheet}");
            try
            {
                using(var package = new ExcelPackage(this._file))
                {
                    var sheetToClean = package.Workbook.Worksheets[sheet];
                    sheetToClean.DeleteRow(5, 15);
                    package.Save();
                    this._log.writeLog($"(SUCCESS) LIMPIEZA TERMINADA, TODO CORRECTO. ||| TIEMPO DE EJECUCIÓN: {this._et.endExecution()}");
                    return "ELIMINADO";
                }
            }
            catch (Exception ex)
            {
                this._log.writeLog($"(ERROR) HUBO UN LIGERO ERROR AL QUERER LIMPIAR LA HOJA {sheet} ||| TIEMPO DE EJECUCIÓN: {this._et.endExecution()}\n\t\tERROR: {ex.Message}");
                return ex.Message;
            }
        }

        public bool fillHeaderSheet(List<TblTesoreria_Model> data)
        {
            try
            {
                this._log.writeLog("(INFO) COMENZANDO CON LA INSERCIÓN DE INFORMACIÓN DENTRO DE LA HOJA DEL HEADER, DENTRO DEL TEMPLATE");
                this._et.startExecution();

                using (var package = new ExcelPackage(this._file))
                {
                    var sheet = package.Workbook.Worksheets["Statement Headers"];

                    var lstDistinct = data
                        .Select(x => x.Bank_Account_Number)
                        .Distinct();

                    int i = 5;

                    foreach(var accounts in lstDistinct)
                    {
                        var minDate = data
                            .Where(x => x.Bank_Account_Number == accounts)
                            .Min(x => DateTime.Parse(x.Booking_Date));

                        var maxDate = data
                            .Where(x => x.Bank_Account_Number == accounts)
                            .Max(x => DateTime.Parse(x.Booking_Date));

                        var hola = accounts.Substring(0, accounts.Length - 6);

                        var stmntNumber = string.Concat(
                                this._preBank.Where(x => x.NombreBanco == this.bank).Select(x => x.Prefijo).FirstOrDefault(), "-",
                                Int64.Parse(accounts.Substring(0, accounts.Length - 6)), "-",
                                minDate.ToString("MMddyyyy")
                            );

                        sheet.Cells[$"A{i}"].Value = stmntNumber;
                        sheet.Cells[$"B{i}"].Value = accounts;
                        sheet.Cells[$"C{i}"].Value = "N";
                        sheet.Cells[$"D{i}"].Value = minDate.ToString("MM/dd/yyyy");
                        sheet.Cells[$"E{i}"].Value = data.Where(x => x.Bank_Account_Number == accounts).Select(x => x.Bank_Account_Currency).FirstOrDefault();
                        sheet.Cells[$"F{i}"].Value = minDate.ToString("MM/dd/yyyy");
                        sheet.Cells[$"G{i}"].Value = maxDate.ToString("MM/dd/yyyy");

                        i++;
                    }

                    sheet.Cells[sheet.Dimension.Address].AutoFitColumns();
                    sheet.Row(1).CustomHeight = false;
                    package.Save();

                    this._log.writeLog($"(SUCCESS) SE LLENÓ LA HOJA DEL HEADER, CORRECTAMENTE ||| TIEMPO DE EJECUCIÓN: {this._et.endExecution()}");
                    return true;
                }
            }
            catch(Exception ex)
            {
                this._log.writeLog($"(ERROR) HUBO UN PEQUEÑO ERROR AL QUERER LLENAR EL HEADER DEL TEMPLATE. ||| TIEMPO DE EJECUCIÓN: {this._et.endExecution()} NOS ARROJÓ: {ex.Message}");
                return false;
            }
        }

        public bool fillBalanceSheet(List<TblTesoreria_Model> data)
        {
            try
            {
                this._log.writeLog("(INFO) COMENZANDO CON LA INSERCIÓN DE INFORMACIÓN DENTRO DE LA HOJA DEL BALANCES, DENTRO DEL TEMPLATE");
                this._et.startExecution();

                using (var package = new ExcelPackage(this._file))
                {
                    var sheet = package.Workbook.Worksheets["Statement Balances"];

                    var lstDistinct = data
                        .Select(x => x.Bank_Account_Number)
                        .Distinct();

                    int i = 5;

                    foreach (var accounts in lstDistinct)
                    {
                        var lstAccount = data
                            .Where(x => x.Bank_Account_Number == accounts)
                            .Select(x => new
                            {
                                x.Open_Balance,
                                x.Close_Balance,
                                x.Bank_Account_Currency
                            })
                            .FirstOrDefault();

                        var minDate = data
                            .Where(x => x.Bank_Account_Number == accounts)
                            .Min(x => DateTime.Parse(x.Booking_Date));

                        var maxDate = data
                            .Where(x => x.Bank_Account_Number == accounts)
                            .Max(x => DateTime.Parse(x.Booking_Date));

                        var stmntNumber = string.Concat(
                                this._preBank.Where(x => x.NombreBanco == this.bank).Select(x => x.Prefijo).FirstOrDefault().FirstOrDefault(), "-",
                                Int64.Parse(accounts.Substring(0, accounts.Length - 6)), "-",
                                minDate.ToString("MMddyyyy")
                            );

                        sheet.Cells[$"A{i}:A{i + 1}"].Value = stmntNumber;
                        sheet.Cells[$"B{i}:B{i + 1}"].Value = accounts;
                        sheet.Cells[$"C{i}"].Value          = "OPBD";
                        sheet.Cells[$"C{i + 1}"].Value      = "CLBD";
                        sheet.Cells[$"D{i}"].Value          = lstAccount.Open_Balance;
                        sheet.Cells[$"D{i + 1}"].Value      = lstAccount.Close_Balance;
                        sheet.Cells[$"E{i}:E{i + 1}"].Value = lstAccount.Bank_Account_Currency;
                        sheet.Cells[$"F{i}:F{i + 1}"].Value = "CRDT";
                        sheet.Cells[$"G{i}"].Value          = minDate.ToString("MM/dd/yyyy");
                        sheet.Cells[$"G{i + 1}"].Value      = maxDate.ToString("MM/dd/yyyy");

                        i = i + 2;
                    }

                    sheet.Cells[sheet.Dimension.Address].AutoFitColumns();
                    sheet.Row(1).CustomHeight = false;
                    package.Save();

                    this._log.writeLog($"(SUCCESS) SE LLENÓ LA HOJA DEL BALANCES, CORRECTAMENTE ||| TIEMPO DE EJECUCIÓN {this._et.endExecution()}");
                    return true;
                }
            }
            catch(Exception ex)
            {
                this._log.writeLog($"(ERROR) HUBO UN PEQUEÑO ERROR AL QUERER LLENAR EL BALANCES DEL TEMPLATE. NOS ARROJÓ: {ex.Message} ||| TIEMPO DE EJECUCIÓN {this._et.endExecution()}");
                return false;
            }
        }

        public bool fillLinesSheet(List<TblTesoreria_Model> data)
        {
            try
            {
                this._log.writeLog("(INFO) COMENZANDO CON LA INSERCIÓN DE INFORMACIÓN DENTRO DE LA HOJA DE LINES, DENTRO DEL TEMPLATE");
                this._et.startExecution();

                using (var package = new ExcelPackage(this._file))
                {
                    var sheet = package.Workbook.Worksheets["Statement Lines"];
                    int i = 5;
                    int j = 1;

                    DateTime maxDate;
                    DateTime minDate;

                    foreach (var rows in data)
                    {
                        maxDate = data
                            .Where(x => !string.IsNullOrEmpty(x.Booking_Date))
                            .Max(x => DateTime.Parse(x.Booking_Date));

                        minDate = data
                            .Where(x => !string.IsNullOrEmpty(x.Booking_Date))
                            .Min(x => DateTime.Parse(x.Booking_Date));

                        if (rows.Booking_Date != null)
                        {
                            maxDate = data
                                .Where(x => x.Bank_Account_Number == rows.Bank_Account_Number)
                                .Max(x => DateTime.Parse(x.Booking_Date));

                            minDate = data
                                .Where(x => x.Bank_Account_Number == rows.Bank_Account_Number)
                                .Min(x => DateTime.Parse(x.Booking_Date));
                        }

                        var stmntNumber = string.Concat(
                                this._preBank.Where(x => x.NombreBanco == this.bank).Select(x => x.Prefijo).FirstOrDefault(), "-",
                                Int64.Parse(rows.Bank_Account_Number.Substring(0, rows.Bank_Account_Number.Length - 6)), "-",
                                minDate.ToString("MMddyyyy")
                            );

                        if (sheet.Cells[$"B{i - 1}"].Text != rows.Bank_Account_Number) j = 1;

                        if (!string.Equals(rows.Credit, "SIN MOVIMIENTOS", StringComparison.CurrentCultureIgnoreCase) || !string.Equals(rows.Debit, "SIN MOVIMIENTOS", StringComparison.CurrentCultureIgnoreCase))
                        {
                            DateTime bookingDate = DateTime.Parse(rows.Booking_Date);
                            DateTime valueDate = DateTime.Parse(rows.Value_Date);

                            sheet.Cells[$"A{i}"].Value = stmntNumber;
                            sheet.Cells[$"B{i}"].Value = rows.Bank_Account_Number;
                            sheet.Cells[$"C{i}"].Value = j;
                            sheet.Cells[$"D{i}"].Value = rows.Transaction_Code ?? "0";
                            sheet.Cells[$"E{i}"].Value = "MSC";
                            sheet.Cells[$"F{i}"].Value = rows.Debit != "0.0" ? rows.Debit : rows.Credit;
                            sheet.Cells[$"G{i}"].Value = rows.Bank_Account_Currency;
                            sheet.Cells[$"H{i}"].Value = bookingDate.ToString("MM/dd/yyyy");
                            sheet.Cells[$"I{i}"].Value = valueDate.ToString("MM/dd/yyyy");
                            sheet.Cells[$"J{i}"].Value = rows.Debit != "0.0" ? "DBIT" : "CRDT";
                            sheet.Cells[$"L{i}"].Value = rows.Check_Number ?? "";
                            sheet.Cells[$"R{i}"].Value = rows.Addenda_Text ?? "";
                            sheet.Cells[$"S{i}"].Value = rows.Account_Servicer_Reference ?? "";
                            sheet.Cells[$"T{i}"].Value = rows.Customer_Reference ?? "";
                            sheet.Cells[$"U{i}"].Value = rows.Clearing_System_Reference ?? "";
                            sheet.Cells[$"V{i}"].Value = rows.Contract_Identifier ?? "";
                            sheet.Cells[$"W{i}"].Value = rows.Instruction_Identifier ?? "";
                            sheet.Cells[$"X{i}"].Value = rows.End_To_End_Identifier ?? "";
                            sheet.Cells[$"Y{i}"].Value = rows.Servicer_Status ?? "";
                            sheet.Cells[$"Z{i}"].Value = rows.Commision_Waiver_Indicator_Flag ?? "";
                            sheet.Cells[$"AA{i}"].Value = rows.Reversal_Indicator_Flag ?? "";
                            sheet.Cells[$"BM{i}"].Value = rows.Structured_Payment_Reference ?? "";
                            sheet.Cells[$"BN{i}"].Value = rows.Reconciliation_Reference ?? "";
                            sheet.Cells[$"BO{i}"].Value = rows.Message_Identifier ?? "";
                            sheet.Cells[$"BP{i}"].Value = rows.Payment_Information_Identifier ?? "";

                            i++;
                            j++;
                        }
                    }

                    sheet.Cells[sheet.Dimension.Address].AutoFitColumns();
                    sheet.Row(1).CustomHeight = false;
                    package.Save();

                    this._log.writeLog($"(SUCCESS) SE LLENÓ LA HOJA DE LINES, CORRECTAMENTE ||| TIEMPO DE EJECUCIÓN {this._et.endExecution()}");
                    return true;
                }
            }
            catch (Exception ex)
            {
                this._log.writeLog($"(ERROR) HUBO UN PEQUEÑO ERROR AL QUERER LLENAR LA HOJA DE LINES DEL TEMPLATE. NOS ARROJÓ: {ex.Message} ||| TIEMPO DE EJECUCIÓN: {this._et.endExecution()}");
                return false;
            }
        }

        public string getTemplate(List<TblTesoreria_Model> data)
        {
            try
            {
                using(var package = new ExcelPackage(this._file))
                {
                    var sheet = package.Workbook.Worksheets["Statement Lines"];
                    var sheetHeader = package.Workbook.Worksheets["Statement Headers"];
                    var sheetBalances = package.Workbook.Worksheets["Statement Balances"];

                    DateTime maxDate;
                    DateTime minDate;

                    var stmntNumber = "";

                    var i = 5;
                    var j = 1;
                    
                    this._log.writeLog($"(INFO) COMIENZO CON CICLO PARA LA INSERCIÓN DE DATOS.\n\t\tSE INSERTARAN {data.Count} REGISTROS");

                    foreach (var rows in data)
                    {
                        var accounts = rows.Bank_Account_Number;
                        accounts = accounts.Substring(accounts.Length - 6);

                        maxDate = data
                            .Where(x => !string.IsNullOrEmpty(x.Booking_Date))
                            .Max(x => DateTime.Parse(x.Booking_Date));

                        minDate = data
                            .Where(x => !string.IsNullOrEmpty(x.Booking_Date))
                            .Min(x => DateTime.Parse(x.Booking_Date));

                        if(rows.Booking_Date != null)
                        {
                            maxDate = data
                                .Where(x => x.Bank_Account_Number == rows.Bank_Account_Number)
                                .Max(x => DateTime.Parse(x.Booking_Date));

                            minDate = data
                                .Where(x => x.Bank_Account_Number == rows.Bank_Account_Number)
                                .Min(x => DateTime.Parse(x.Booking_Date));
                        }

                        if (sheet.Cells[$"B{i - 1}"].Text == rows.Bank_Account_Number) j++;
                        else
                        {
                            stmntNumber = string.Concat(
                                this._preBank.Find(x => x.NombreBanco.Contains(bank)).Prefijo, "-",
                                int.Parse(accounts), "-",
                                minDate.ToString("MMddyyyy")
                            );

                            sheetHeader.Cells[$"A{_rowHeader}"].Value = stmntNumber;
                            sheetHeader.Cells[$"B{_rowHeader}"].Value = rows.Bank_Account_Number;
                            sheetHeader.Cells[$"C{_rowHeader}"].Value = "N";
                            sheetHeader.Cells[$"D{_rowHeader}"].Value = minDate.ToString("MM/dd/yyyy");
                            sheetHeader.Cells[$"E{_rowHeader}"].Value = rows.Bank_Account_Currency;
                            sheetHeader.Cells[$"F{_rowHeader}"].Value = minDate.ToString("MM/dd/yyyy");
                            sheetHeader.Cells[$"G{_rowHeader}"].Value = maxDate.ToString("MM/dd/yyyy");

                            sheetBalances.Cells[$"A{_rowBalances}:A{_rowBalances + 1}"].Value   = stmntNumber;
                            sheetBalances.Cells[$"B{_rowBalances}:B{_rowBalances + 1}"].Value   = rows.Bank_Account_Number;
                            sheetBalances.Cells[$"C{_rowBalances}"].Value                       = "OPBD";
                            sheetBalances.Cells[$"C{_rowBalances + 1}"].Value                   = "CLBD";
                            sheetBalances.Cells[$"D{_rowBalances}"].Value                       = rows.Open_Balance;
                            sheetBalances.Cells[$"D{_rowBalances + 1}"].Value                   = rows.Close_Balance;
                            sheetBalances.Cells[$"E{_rowBalances}:E{_rowBalances + 1}"].Value   = rows.Bank_Account_Currency;
                            sheetBalances.Cells[$"F{_rowBalances}:F{_rowBalances + 1}"].Value   = "CRDT";
                            sheetBalances.Cells[$"G{_rowBalances}"].Value                       = minDate.ToString("MM/dd/yyyy");
                            sheetBalances.Cells[$"G{_rowBalances + 1}"].Value                   = maxDate.ToString("MM/dd/yyyy");

                            this._rowBalances = this._rowBalances + 2;
                            this._rowHeader++;
                            j = 1;
                        }

                        if (!string.Equals(rows.Credit, "SIN MOVIMIENTOS", StringComparison.CurrentCultureIgnoreCase) || !string.Equals(rows.Debit, "SIN MOVIMIENTOS", StringComparison.CurrentCultureIgnoreCase))
                        {
                            DateTime bookingDate = DateTime.Parse(rows.Booking_Date);
                            DateTime valueDate = DateTime.Parse(rows.Value_Date);

                            sheet.Cells[$"A{i}"].Value  = stmntNumber;
                            sheet.Cells[$"B{i}"].Value  = rows.Bank_Account_Number;
                            sheet.Cells[$"C{i}"].Value  = j;
                            sheet.Cells[$"D{i}"].Value  = rows.Transaction_Code ?? "0";
                            sheet.Cells[$"E{i}"].Value  = "MSC";
                            sheet.Cells[$"F{i}"].Value  = rows.Debit != "0.0" ? rows.Debit : rows.Credit;
                            sheet.Cells[$"G{i}"].Value  = rows.Bank_Account_Currency;
                            sheet.Cells[$"H{i}"].Value  = bookingDate.ToString("MM/dd/yyyy");
                            sheet.Cells[$"I{i}"].Value  = valueDate.ToString("MM/dd/yyyy");
                            sheet.Cells[$"J{i}"].Value  = rows.Debit != "0.0" ? "DBIT" : "CRDT";
                            sheet.Cells[$"L{i}"].Value  = rows.Check_Number ?? "";
                            sheet.Cells[$"R{i}"].Value  = rows.Addenda_Text ?? "";
                            sheet.Cells[$"S{i}"].Value  = rows.Account_Servicer_Reference ?? "";
                            sheet.Cells[$"T{i}"].Value  = rows.Customer_Reference ?? "";
                            sheet.Cells[$"U{i}"].Value  = rows.Clearing_System_Reference ?? "";
                            sheet.Cells[$"V{i}"].Value  = rows.Contract_Identifier ?? "";
                            sheet.Cells[$"W{i}"].Value  = rows.Instruction_Identifier ?? "";
                            sheet.Cells[$"X{i}"].Value  = rows.End_To_End_Identifier ?? "";
                            sheet.Cells[$"Y{i}"].Value  = rows.Servicer_Status ?? "";
                            sheet.Cells[$"Z{i}"].Value  = rows.Commision_Waiver_Indicator_Flag ?? "";
                            sheet.Cells[$"AA{i}"].Value  = rows.Reversal_Indicator_Flag ?? "";
                            sheet.Cells[$"BM{i}"].Value  = rows.Structured_Payment_Reference ?? "";
                            sheet.Cells[$"BN{i}"].Value  = rows.Reconciliation_Reference ?? "";
                            sheet.Cells[$"BO{i}"].Value  = rows.Message_Identifier ?? "";
                            sheet.Cells[$"BP{i}"].Value  = rows.Payment_Information_Identifier ?? "";

                            i++;
                        }

                    }
                    sheet.Cells[sheet.Dimension.Address].AutoFitColumns();
                    sheet.Row(1).CustomHeight = false;
                    package.Save();
                    this._log.writeLog($"(SUCCESS) SE INSERTARON LOS REGISTROS CORRECTAMENTE");
                    return "CORRECTO";
                }
            }
            catch(Exception ex)
            {
                this._log.writeLog($"(ERROR) HUBO UN LIGERO ERROR AL INSERTAR LOS DATOS\n\t\tERROR: {ex.Message}");
                return $"Hubo un pequeño error: {ex.Message}";
            }
        }

        public void closeDocument()
        {
            Excel.Application excelApp = null;

            var index = this._path.LastIndexOf(@"\\");
            var file = "";

            if (index != -1)
                file = this._path.Substring(index + 1);

            try
            {
                excelApp = (Excel.Application)Marshal.GetActiveObject("Excel.Application");
                foreach(Excel.Workbook wb in excelApp.Workbooks)
                {
                    if(wb.FullName.EndsWith(file))
                    {
                        wb.Close(true);
                        break;
                    }
                }

                if (excelApp.Workbooks.Count == 0)
                    excelApp.Quit();
            }
            catch(Exception ex)
            {
                this._log.writeLog($"HUBO UN PEQUEÑO ERROR AL QUERER CERRAR EL DOCUMENTO DE EXCEL\n\t\tERROR: {ex.Message}");
            }
        }
    }
}
