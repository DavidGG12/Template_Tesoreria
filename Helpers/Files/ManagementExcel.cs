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
        private string _bank;
        private FileInfo _file;
        private Log _log;
        private ExecutionTimer _et;
        private List<BankPrefix_Model> _preBank;

        public ManagementExcel(string pathExcel, string bank) 
        {
            this._bank = bank;
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
                                this._preBank.Where(x => x.NombreBanco == this._bank).Select(x => x.Prefijo).FirstOrDefault(), "-",
                                Int64.Parse(accounts), "-",
                                minDate.ToString("MMddyyyy")
                            );

                        sheet.Cells[i, 1].Value = stmntNumber;
                        sheet.Cells[i, 2].Value = accounts;
                        sheet.Cells[i, 3].Value = "N";
                        sheet.Cells[i, 4].Value = minDate.ToString("MM/dd/yyyy");
                        sheet.Cells[i, 5].Value = data.Where(x => x.Bank_Account_Number == accounts).Select(x => x.Bank_Account_Currency).FirstOrDefault();
                        sheet.Cells[i, 6].Value = minDate.ToString("MM/dd/yyyy");
                        sheet.Cells[i, 7].Value = maxDate.ToString("MM/dd/yyyy");

                        i++;
                    }

                    sheet.Cells[1, 1, i, 20].AutoFitColumns();
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
                                this._preBank.Where(x => x.NombreBanco == this._bank).Select(x => x.Prefijo).FirstOrDefault(), "-",
                                Int64.Parse(accounts), "-",
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

                    sheet.Cells[1, 1, i, 20].AutoFitColumns();
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

                // Precalcular fechas por cuenta
                var fechasPorCuenta = data
                    .Where(x => !string.IsNullOrEmpty(x.Booking_Date))
                    .GroupBy(x => x.Bank_Account_Number)
                    .ToDictionary(
                        g => g.Key,
                        g => new
                        {
                            Min = g.Min(x => DateTime.Parse(x.Booking_Date)),
                            Max = g.Max(x => DateTime.Parse(x.Booking_Date))
                        }
                    );

                using (var package = new ExcelPackage(this._file))
                {
                    var sheet = package.Workbook.Worksheets["Statement Lines"];
                    int i = 5; // Empieza en fila 5
                    int j = 1;

                    foreach (var rows in data)
                    {
                        if (!fechasPorCuenta.TryGetValue(rows.Bank_Account_Number, out var fechas))
                            continue;

                        var minDate = fechas.Min;
                        var stmntNumber = string.Concat(
                            this._preBank.Where(x => x.NombreBanco == this._bank).Select(x => x.Prefijo).FirstOrDefault(), "-",
                            Int64.Parse(rows.Bank_Account_Number), "-",
                            minDate.ToString("MMddyyyy")
                        );

                        if (sheet.Cells[i - 1, 2].Text != rows.Bank_Account_Number) j = 1;

                        if (!string.Equals(rows.Credit, "SIN MOVIMIENTOS", StringComparison.CurrentCultureIgnoreCase) ||
                            !string.Equals(rows.Debit, "SIN MOVIMIENTOS", StringComparison.CurrentCultureIgnoreCase))
                        {
                            var bookingDate = DateTime.Parse(rows.Booking_Date);
                            var valueDate = DateTime.Parse(rows.Value_Date);

                            sheet.Cells[i, 1].Value = stmntNumber;
                            sheet.Cells[i, 2].Value = rows.Bank_Account_Number;
                            sheet.Cells[i, 3].Value = j;
                            sheet.Cells[i, 4].Value = rows.Transaction_Code ?? "0";
                            sheet.Cells[i, 5].Value = "MSC";
                            sheet.Cells[i, 6].Value = rows.Debit != "0.0" ? rows.Debit : rows.Credit;
                            sheet.Cells[i, 7].Value = rows.Bank_Account_Currency;
                            sheet.Cells[i, 8].Value = bookingDate.ToString("MM/dd/yyyy");
                            sheet.Cells[i, 9].Value = valueDate.ToString("MM/dd/yyyy");
                            sheet.Cells[i, 10].Value = rows.Debit != "0.0" ? "DBIT" : "CRDT";
                            sheet.Cells[i, 12].Value = rows.Check_Number ?? "";
                            sheet.Cells[i, 18].Value = rows.Addenda_Text ?? "";
                            sheet.Cells[i, 19].Value = rows.Account_Servicer_Reference ?? "";
                            sheet.Cells[i, 20].Value = rows.Customer_Reference ?? "";
                            sheet.Cells[i, 21].Value = rows.Clearing_System_Reference ?? "";
                            sheet.Cells[i, 22].Value = rows.Contract_Identifier ?? "";
                            sheet.Cells[i, 23].Value = rows.Instruction_Identifier ?? "";
                            sheet.Cells[i, 24].Value = rows.End_To_End_Identifier ?? "";
                            sheet.Cells[i, 25].Value = rows.Servicer_Status ?? "";
                            sheet.Cells[i, 26].Value = rows.Commision_Waiver_Indicator_Flag ?? "";
                            sheet.Cells[i, 27].Value = rows.Reversal_Indicator_Flag ?? "";
                            sheet.Cells[i, 65].Value = rows.Structured_Payment_Reference ?? "";
                            sheet.Cells[i, 66].Value = rows.Reconciliation_Reference ?? "";
                            sheet.Cells[i, 67].Value = rows.Message_Identifier ?? "";
                            sheet.Cells[i, 68].Value = rows.Payment_Information_Identifier ?? "";

                            i++;
                            j++;
                        }
                    }

                    sheet.Cells[1, 1, i, 20].AutoFitColumns();
                    sheet.Row(1).CustomHeight = false;
                    package.Save();

                    this._log.writeLog($"(SUCCESS) SE LLENÓ LA HOJA DE LINES, CORRECTAMENTE ||| TIEMPO DE EJECUCIÓN {this._et.endExecution()}");
                    return true;
                }
            }
            catch (Exception ex)
            {
                this._log.writeLog($"(ERROR) HUBO UN ERROR AL LLENAR LA HOJA DE LINES: {ex.Message} ||| TIEMPO: {this._et.endExecution()}");
                return false;
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
