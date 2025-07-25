//using Microsoft.Office.Interop.ExcKel;
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
        private List<BancoPrefijoModel> _preBank;
        public ManagementExcel(string pathExcel, string bank) 
        {
            this._rowHeader = 5;
            this._rowBalances = 5;
            this.bank = bank;
            this._path = pathExcel;
            this._file = new FileInfo(this._path);
            this._log = new Log();
            this._preBank = new List<BancoPrefijoModel>()
            {
                new BancoPrefijoModel(){ NombreBanco = "Inbursa",       Prefijo = "INB"  },
                new BancoPrefijoModel(){ NombreBanco = "HSBC",          Prefijo = "HSBC" },
                new BancoPrefijoModel(){ NombreBanco = "Bancomer",      Prefijo = "BBVA" },
                new BancoPrefijoModel(){ NombreBanco = "Scotiabank",    Prefijo = "SCOT" },
                new BancoPrefijoModel(){ NombreBanco = "Citibanamex",   Prefijo = "CITI" },
                new BancoPrefijoModel(){ NombreBanco = "Santander",     Prefijo = "SANT" },
                new BancoPrefijoModel(){ NombreBanco = "Banorte",       Prefijo = "BAN"  }
            };
            ExcelPackage.License.SetNonCommercialOrganization("Grupo Sanborns");
        }

        public string cleanSheets(string sheet)
        {
            this._log.writeLog($"LIMPIEZA DE LA HOJA {sheet}");
            try
            {
                using(var package = new ExcelPackage(this._file))
                {
                    var sheetToClean = package.Workbook.Worksheets[sheet];
                    sheetToClean.DeleteRow(5, 15);
                    package.Save();
                    this._log.writeLog($"LIMPIEZA TERMINADA, TODO CORRECTO");
                    return "ELIMINADO";
                }
            }
            catch (Exception ex)
            {
                this._log.writeLog($"HUBO UN LIGERO ERROR AL QUERER LIMPIAR LA HOJA {sheet}\n\t\tERROR: {ex.Message}");
                return ex.Message;
            }
        }

        public string getTemplate(List<Tbl_Tesoreria_Ext_Bancario> data)
        {
            try
            {
                using(var package = new ExcelPackage(this._file))
                {
                    var sheet = package.Workbook.Worksheets["Statement Lines"];
                    var sheetHeader = package.Workbook.Worksheets["Statement Headers"];
                    var sheetBalances = package.Workbook.Worksheets["Statement Balances"];
                    var dateDoc = data.Find(x => x.Fecha != null && x.Fecha.Any(f => f != null)).Fecha;

                    string[] formats = { "MM/dd/yyyy", "dd/MM/yyyy", "yyyy-MM-dd" };
                    DateTime dateParse;

                    bool tryParse = DateTime.TryParseExact(
                        dateDoc,
                        formats,
                        System.Globalization.CultureInfo.InvariantCulture,
                        System.Globalization.DateTimeStyles.None,
                        out dateParse
                    );

                    var formDate = dateParse.ToString("MM/dd/yyyy");
                    var i = 5;
                    var j = 1;
                    
                    this._log.writeLog($"COMIENZO CON CICLO PARA LA INSERCIÓN DE DATOS.\n\t\tSE INSERTARAN {data.Count} REGISTROS");

                    foreach (var rows in data)
                    {
                        var accounts = rows.Cuenta.Replace("-PESOS", "") ?? "";
                        accounts = accounts.Substring(accounts.Length - 6);

                        var stmntNumber = string.Concat(
                            this._preBank.Find(x => x.NombreBanco.Contains(bank)).Prefijo, "-",
                            int.Parse(accounts), "-",
                            formDate.Replace("/", "")
                        );

                        if (sheet.Cells[$"B{i - 1}"].Text == rows.Cuenta.Replace("-PESOS", "")) j++;
                        else
                        {
                            sheetHeader.Cells[$"A{_rowHeader}"].Value = stmntNumber;
                            sheetHeader.Cells[$"B{_rowHeader}"].Value = rows.Cuenta.Replace("-PESOS", "") ?? "";
                            sheetHeader.Cells[$"C{_rowHeader}"].Value = "N";
                            sheetHeader.Cells[$"D{_rowHeader}"].Value = formDate;
                            sheetHeader.Cells[$"E{_rowHeader}"].Value = rows.Moneda;
                            sheetHeader.Cells[$"F{_rowHeader}"].Value = formDate;
                            sheetHeader.Cells[$"G{_rowHeader}"].Value = formDate;

                            sheetBalances.Cells[$"A{_rowBalances}:A{_rowBalances + 1}"].Value   = stmntNumber;
                            sheetBalances.Cells[$"B{_rowBalances}:B{_rowBalances + 1}"].Value   = rows.Cuenta.Replace("-PESOS", "") ?? "";
                            sheetBalances.Cells[$"C{_rowBalances}"].Value                       = "OPBD";
                            sheetBalances.Cells[$"C{_rowBalances + 1}"].Value                   = "CLBD";
                            sheetBalances.Cells[$"D{_rowBalances}"].Value                       = rows.Saldo_Inicial;
                            sheetBalances.Cells[$"D{_rowBalances + 1}"].Value                   = rows.Saldo_Final;
                            sheetBalances.Cells[$"E{_rowBalances}:E{_rowBalances + 1}"].Value   = rows.Moneda;
                            sheetBalances.Cells[$"F{_rowBalances}:F{_rowBalances + 1}"].Value   = "CRDT";
                            sheetBalances.Cells[$"G{_rowBalances}:G{_rowBalances + 1}"].Value   = formDate;

                            this._rowBalances = this._rowBalances + 2;
                            this._rowHeader++;
                            j = 1;
                        }

                        if (!string.Equals(rows.Referencia, "SIN MOVIMIENTOS", StringComparison.CurrentCultureIgnoreCase))
                        {
                            sheet.Cells[$"A{i}"].Value  = stmntNumber;
                            sheet.Cells[$"B{i}"].Value  = rows.Cuenta.Replace("-PESOS", "") ?? "";
                            sheet.Cells[$"C{i}"].Value  = j;
                            sheet.Cells[$"D{i}"].Value  = rows.Concepto ?? "";
                            sheet.Cells[$"E{i}"].Value  = "MSC";
                            sheet.Cells[$"F{i}"].Value  = rows.Cargo != "0.0" ? rows.Cargo : rows.Abono;
                            sheet.Cells[$"G{i}"].Value  = rows.Moneda;
                            sheet.Cells[$"H{i}"].Value  = formDate;
                            sheet.Cells[$"J{i}"].Value  = rows.Cargo != "0.0" ? "DBIT" : "CRDT";

                            if (string.IsNullOrEmpty(rows.Informacion_Env))
                            {
                                sheet.Cells[$"L{i}"].Value = rows.Referencia ?? "";
                                sheet.Cells[$"S{i}"].Value  = rows.RFC_Ordenante ?? "";
                            }
                            else
                            {
                                var addText = "";       //Columna R dentro del template: "Addenda Text"
                                var accServRef = "";    //Columna S dentro del template: "Account Servicer Reference"

                                foreach(var letter in rows.Informacion_Env)
                                {
                                    if (string.Equals(letter, ' '))
                                        break;
                                    addText += letter;
                                }

                                accServRef = rows.Informacion_Env.Replace(addText, "").Trim();

                                sheet.Cells[$"R{i}"].Value = addText;
                                sheet.Cells[$"S{i}"].Value = accServRef;
                            }

                            sheet.Cells[$"T{i}"].Value  = rows.Ordenante ?? "";
                            sheet.Cells[$"W{i}"].Value  = rows.Movimiento ?? "";
                            sheet.Cells[$"X{i}"].Value  = rows.Referencia_Leyenda ?? "";
                            sheet.Cells[$"BN{i}"].Value = rows.Referencia_Ext ?? "";
                            sheet.Cells[$"BP{i}"].Value = rows.Referencia_Numerica ?? "";
                            sheet.Cells[$"BO{i}"].Value = rows.Referencia_Leyenda ?? "";
                            i++;
                        }

                    }
                    sheet.Cells[sheet.Dimension.Address].AutoFitColumns();
                    sheet.Row(1).CustomHeight = false;
                    //sheet.Cells[sheet.Dimension.Address].AutoFitColumns;
                    package.Save();
                    this._log.writeLog($"SE INSERTARON LOS REGISTROS CORRECTAMENTE");
                    return "CORRECTO";
                }
            }
            catch(Exception ex)
            {
                this._log.writeLog($"HUBO UN LIGERO ERROR AL INSERTAR LOS DATOS\n\t\tERROR: {ex.Message}");
                return $"Hubo un pequeño error: {ex.Message}";
            }
        }

        private void fillHeader(Tbl_Tesoreria_Ext_Bancario data)
        {
            try
            {
                this._log.writeLog($"COMIENZO DE LA INSERCIÓN DEL HEADER DE LA CUENTA {data.Cuenta}");
                var sheetsList = new List<string>()
                {
                    { "Statement Headers" }
                };

                using (var package = new ExcelPackage(this._file))
                {
                    foreach(var nmSheet in sheetsList)
                    {
                        var sheet = package.Workbook.Worksheets[nmSheet];
                        this._log.writeLog($"SE TRABAJARÁ CON LA PESTAÑA {nmSheet}");

                        var accounts = data.Cuenta.Replace("-PESOS", "") ?? "";
                        accounts = accounts.Substring(accounts.Length - 6);

                        var stmntNumber = string.Concat(
                            this._preBank.Find(x => x.NombreBanco.Contains(bank)).Prefijo, "-",
                            int.Parse(accounts), "-",
                            data.Fecha.Replace("/", "")
                        );

                        
                    }
                    package.Save();
                }
            }
            catch(Exception ex)
            {
                this._log.writeLog($"Hubo un ligero error al querer llenar el Header.\n\tError: {ex.Message}");
                throw ex;
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
