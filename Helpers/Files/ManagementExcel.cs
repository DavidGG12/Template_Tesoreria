//using Microsoft.Office.Interop.ExcKel;
using Microsoft.SqlServer.Server;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.IO;
using System.Linq;
using System.Runtime.ExceptionServices;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using Template_Tesoreria.Models;
using Excel = Microsoft.Office.Interop.Excel;
//using Spire.Xls;


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
        private List<BankPrefix_Model> _preBank;
        public ManagementExcel(string pathExcel, string bank) 
        {
            this._rowHeader = 5;
            this._rowBalances = 5;
            this.bank = bank;
            this._path = pathExcel;
            this._file = new FileInfo(this._path);
            this._log = new Log();
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

        public string cleanSheets(string sheet)
        {
            this._log.writeLog($"(INFO) LIMPIEZA DE LA HOJA {sheet}");
            try
            {
                using(var package = new ExcelPackage(this._file))
                {
                    var sheetToClean = package.Workbook.Worksheets[sheet];
                    sheetToClean.DeleteRow(5, 15);
                    package.Save();
                    this._log.writeLog($"(SUCCESS) LIMPIEZA TERMINADA, TODO CORRECTO");
                    return "ELIMINADO";
                }
            }
            catch (Exception ex)
            {
                this._log.writeLog($"(ERROR) HUBO UN LIGERO ERROR AL QUERER LIMPIAR LA HOJA {sheet}\n\t\tERROR: {ex.Message}");
                return ex.Message;
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
                            .Max(x => DateTime.Parse(x.Booking_Date));

                        if (rows.Booking_Date != null || rows.Value_Date != null)
                        {
                            maxDate = data
                                .Where(a => a.Bank_Account_Number == rows.Bank_Account_Number)
                                .Max(x => DateTime.Parse(x.Booking_Date));

                            minDate = data
                                .Where(a => a.Bank_Account_Number == rows.Bank_Account_Number)
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
                    //sheet.Cells[sheet.Dimension.Address].AutoFitColumns;
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

        //private void fillHeader(TblTesoreria_Model data)
        //{
        //    try
        //    {
        //        this._log.writeLog($"COMIENZO DE LA INSERCIÓN DEL HEADER DE LA CUENTA {data.Cuenta}");
        //        var sheetsList = new List<string>()
        //        {
        //            { "Statement Headers" }
        //        };

        //        using (var package = new ExcelPackage(this._file))
        //        {
        //            foreach(var nmSheet in sheetsList)
        //            {
        //                var sheet = package.Workbook.Worksheets[nmSheet];
        //                this._log.writeLog($"SE TRABAJARÁ CON LA PESTAÑA {nmSheet}");

        //                var accounts = data.Cuenta.Replace("-PESOS", "") ?? "";
        //                accounts = accounts.Substring(accounts.Length - 6);

        //                var stmntNumber = string.Concat(
        //                    this._preBank.Find(x => x.NombreBanco.Contains(bank)).Prefijo, "-",
        //                    int.Parse(accounts), "-",
        //                    data.Fecha.Replace("/", "")
        //                );

                        
        //            }
        //            package.Save();
        //        }
        //    }
        //    catch(Exception ex)
        //    {
        //        this._log.writeLog($"Hubo un ligero error al querer llenar el Header.\n\tError: {ex.Message}");
        //        throw ex;
        //    }
        //}

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
