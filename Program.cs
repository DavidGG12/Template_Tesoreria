using System;
using System.Collections.Generic;
using System.Threading.Tasks;
using HTML = HtmlAgilityPack;
using System.IO;
using System.Net;
using Template_Tesoreria.Helpers.DataAccess;
using Template_Tesoreria.Models;
using Template_Tesoreria.Helpers.Files;
using System.Diagnostics;
using System.Threading;
using System.Net.Sockets;
using Template_Tesoreria.Helpers.GUI;
using Template_Tesoreria.Helpers.Network;

namespace Template_Tesoreria
{
    internal class Program
    {
        public static string getIP(Log log)
        {
            try
            {
                log.writeLog("SE OBTIENE LA IP DEL USUARIO.");
                foreach (var ipv4 in Dns.GetHostEntry(Dns.GetHostName()).AddressList)
                    if (ipv4.AddressFamily == AddressFamily.InterNetwork)
                    {
                        log.writeLog("OBTENCIÓN DE IP CORRECTA");
                        return ipv4.ToString();
                    }
                return null;
            }
            catch(Exception ex)
            {
                log.writeLog($"(ERROR) HUBO UN ERROR AL QUERER OBTENER LA IP, NOS ARROJA: {ex.Message}");
                return null;
            }
        }

        public static string downloadTemplate(string nmBank, Log log)
        {
            try
            {
                WebClient client1 = new WebClient();
                var urlFile = "";
                var pathDirectory = "";
                var pathDestiny = "";

                string htmlCode = client1.DownloadString("https://docs.oracle.com/en/cloud/saas/financials/25b/oefbf/cashmanagementbankstatementdataimport-3168.html#cashmanagementbankstatementdataimport-3168");
                string[] lines = htmlCode.Split('\n');

                HTML.HtmlDocument htmlDocument = new HTML.HtmlDocument();
                htmlDocument.LoadHtml(lines[58].ToString().Trim());

                var linkNodes = htmlDocument.DocumentNode.SelectNodes("//a[@href]");

                if (linkNodes != null)
                    foreach (var linkNode in linkNodes)
                        urlFile = linkNode.GetAttributeValue("href", string.Empty);

                log.writeLog($"SE OBTUVO LA INFORMACIÓN PARA PODER DESCARGAR CORRECTAMENTE EL TEMPLATE");

                pathDirectory = Environment.GetFolderPath(Environment.SpecialFolder.UserProfile) + @"\\Documents\\Templates";

                //Si no existe la Carpeta la creamos
                if (!Directory.Exists(pathDirectory)) Directory.CreateDirectory(pathDirectory);


                //Definimos la ruta donde guardaremos el archivo
                //http://www.oracle.com/webfolder/technetwork/docs/fbdi-25b/fbdi/xlsm/CashManagementBankStatementImportTemplate.xlsm                
                pathDestiny = Environment.GetFolderPath(Environment.SpecialFolder.UserProfile) + @"\\Documents\\Templates\\CashManagementBankStatementImportTemplate_" + nmBank + ".xlsm";
                log.writeLog($"EL TEMPLATE SE INSERTARÁ EN LA SIGUIENTE RUTA: {pathDestiny}");

                WebClient myWebClient = new WebClient();
                myWebClient.DownloadFile(urlFile, pathDestiny);

                log.writeLog($"SE DESCARGA EL TEMPLATE");
                log.writeLog($"EMPIEZA LA INSERCIÓN DE LOS DATOS EN EL TEMPLATE");

                return "TEMPLATE DESCARGADO";
            }
            catch(Exception ex)
            {
                log.writeLog($"(ERROR): AL DESCARGAR EL TEMPLATE SE GENERÓ UN ERROR: {ex.Message}");
                return null;
            }
        }

        static void Main(string[] args)
        {
            var dtService = new DataService();
            var gui = new GUI_Main();
            var cnn = new ConnectionDb();
            var cts = new CancellationTokenSource();
            var log = new Log();
            var options = new List<MenuOption_Model>()
            {
                new MenuOption_Model() { ID = "1", Option = "1. - INBURSA", Value = "Inbursa" },
                new MenuOption_Model() { ID = "2", Option = "2. - HSBC", Value = "HSBC" },
                new MenuOption_Model() { ID = "3", Option = "3. - BANCOMER", Value = "Bancomer" },
                new MenuOption_Model() { ID = "4", Option = "4. - SCOTIABANK", Value = "Scotiabank" },
                new MenuOption_Model() { ID = "5", Option = "5. - CITIBANAMEX", Value = "Citibanamex" },
                new MenuOption_Model() { ID = "6", Option = "6. - SANTANDER", Value = "Santander" },
                new MenuOption_Model() { ID = "7", Option = "7. - BANORTE", Value = "Banorte" }
            };
            var ip = "";
            var nmBank = "";
            var pathDestiny = "";
            var id = 1;

            ConsoleKey key;

            while(true)
            {
                try
                {
                    log.writeLog("COMENZANDO PROCESO");

                    nmBank = gui.viewMenu("Extracto bancario ", "", options);

                    var shrdDirectory = new SharedDirectory("10.128.10.19");
                    var nmFile = gui.viewMenu(" ", "", shrdDirectory.getFiles());

                    gui.viewMainMessage("********COMENZANDO PROCESO********");

                    gui.viewInfoMessage("*Descargando el template desde el sitio de Oracle*");

                    var rsltDownload = "";
                    Task.Run(() =>
                    {
                        rsltDownload = downloadTemplate(nmBank, log);
                        cts.Cancel();
                    });
                    gui.Spinner("Descargando...", cts.Token);
                    cts = new CancellationTokenSource();

                    if(!string.Equals(rsltDownload, "TEMPLATE DESCARGADO", StringComparison.CurrentCultureIgnoreCase))
                    {
                        gui.viewErrorMessage("(ERROR) Algo ocurrió al querer descargar el template.");
                        break;
                    }

                    gui.viewInfoMessage("*Obteniendo IP de la PC con la que se va a trabajar*");
                    Task.Run(() =>
                    {
                        ip = getIP(log);
                        cts.Cancel();
                    });
                    gui.Spinner("Obteniendo IP...", cts.Token);
                    cts = new CancellationTokenSource();


                    var valueFile = new ValueFile_Model();
                    switch (nmBank)
                    {
                        case "Inbursa":
                            //valueFile.FileName = $"INBM{DateTime.Now.ToString("ddMMyy")}.xlsx";
                            valueFile.FileName = $"INBM280725.xlsx";
                            valueFile.SPName = $"pa_Tesoreria_CargaExcel_Inbursa";
                            break;

                        case "HSBC":
                            //valueFile.FileName = $"HSBC{DateTime.Now.ToString("ddMMyy")}.xlsx";
                            valueFile.FileName = $"HSBC280725.xlsx";
                            valueFile.SPName = $"pa_Tesoreria_CargaExcel_HSBC";
                            break;

                        case "Bancomer":
                            //valueFile.FileName = $"BBVA{DateTime.Now.ToString("ddMMyy")}.xlsx";
                            valueFile.FileName = $"BBVA140725.xlsx";
                            valueFile.SPName = $"pa_Tesoreria_CargaExcel_BBVA";
                            break;

                        case "Scotiabank":
                            //valueFile.FileName = $"SCOT{DateTime.Now.ToString("ddMMyy")}.xlsx";
                            valueFile.FileName = $"SCOT100625.xlsx";
                            valueFile.SPName = $"pa_Tesoreria_CargaExcel_Scotiabank";
                            break;

                        case "Citibanamex":
                            //valueFile.FileName = $"CITI{DateTime.Now.ToString("ddMMyy")}.xlsx";
                            valueFile.FileName = $"CITI140725.xlsx";
                            valueFile.SPName = $"pa_Tesoreria_CargaExcel_Citi";
                            break;

                        case "Santander":
                            //valueFile.FileName = $"SANT{DateTime.Now.ToString("ddMMyy")}.xlsx";
                            valueFile.FileName = $"SANT100725.xlsx";
                            valueFile.SPName = $"pa_Tesoreria_CargaExcel_Santander";
                            break;

                        case "Banorte":
                            //valueFile.FileName = $"BANO{DateTime.Now.ToString("ddMMyy")}.xlsx";
                            valueFile.FileName = $"BANO100725.xlsx";
                            valueFile.SPName = $"pa_Tesoreria_CargaExcel_Banorte";
                            break;
                    }


                    //Empezamos con la recolección de datos y el llenado de la información
                    var data = new List<TblTesoreria_Model>();
                    var parameters = new Dictionary<string, object>()
                    {
                        { "@Ip", "10.128.10.19" },
                        { "@Excelname", valueFile.FileName }
                    };

                    gui.viewInfoMessage("*Extrayendo información para el llenado de template*");
                    Task.Run(() =>
                        {
                            data = dtService.GetDataList<TblTesoreria_Model>(cnn.DbTesoreria1019(), valueFile.SPName, parameters);
                            cts.Cancel();
                        }
                    );
                    gui.Spinner("Obteniendo...", cts.Token);
                    cts = new CancellationTokenSource();


                    gui.viewInfoMessage("*Limpiando template para su llenado*");
                    log.writeLog($"LIMPIAMOS EL TEMPLATE PARA PODER INSERTAR LOS DATOS");
                    pathDestiny = Environment.GetFolderPath(Environment.SpecialFolder.UserProfile) + @"\\Documents\\Templates\\CashManagementBankStatementImportTemplate_" + nmBank + ".xlsm";
                    var mngmntExcel = new ManagementExcel(pathDestiny, nmBank);
                    var errorList = new List<SheetError_Model>()
                    {
                        new SheetError_Model() { Sheet = "Statement Headers", Message = mngmntExcel.cleanSheets("Statement Headers") },
                        new SheetError_Model() { Sheet = "Statement Balances", Message = mngmntExcel.cleanSheets("Statement Balances") },
                        new SheetError_Model() { Sheet = "Statement Balance Availability", Message = mngmntExcel.cleanSheets("Statement Balance Availability") },
                        new SheetError_Model() { Sheet = "Statement Lines", Message = mngmntExcel.cleanSheets("Statement Lines") },
                        new SheetError_Model() { Sheet = "Statement Line Avilability", Message = mngmntExcel.cleanSheets("Statement Line Availability") },
                        new SheetError_Model() { Sheet = "Statement Statement Line Charges", Message = mngmntExcel.cleanSheets("Statement Line Charges") }
                    };

                    var error = errorList.Find(x => !x.Message.Contains("ELIMINADO"));
                    if(error != null)
                    {
                        gui.viewErrorMessage($"(ERROR) Hubo un ligero error al querer limpiar los datos de la hoja {error.Sheet}. Nos arroja: {error.Message}");
                        log.writeLog($"**********************************************************************");
                        break;
                    }

                    log.writeLog($"TERMINO DE LIMPIEZA, SE PROSIGUE CON LA INSERCIÓN DE DATOS");

                    var fillData = "";
                    gui.viewInfoMessage($"*Llenando template con los datos recuperados. Siendo un total de {data.Count} registros*");
                    Task.Run(() =>
                    {
                        fillData = mngmntExcel.getTemplate(data);
                        cts.Cancel();
                    });
                    gui.Spinner("Llenando...", cts.Token);
                    cts = new CancellationTokenSource();

                    if (!string.Equals(fillData, "CORRECTO", StringComparison.CurrentCultureIgnoreCase))
                    {
                        gui.viewErrorMessage("(ERROR) Hubo un ligero error al querer llenar el template.");
                        break;
                    }

                    Console.Write("\n¿Desea llenar otro template? [S/N]:");
                    var again = Console.ReadLine().Trim();
                    
                    Process.Start(pathDestiny);

                    if (string.Equals(again, "n", StringComparison.OrdinalIgnoreCase))
                        break;

                    log.writeLog($"ABRIENDO ARCHIVO\n\t\t**PROCESO TERMINADO**");
                    log.writeLog($"**********************************************************************");
                }
                catch (Exception ex)
                {
                    gui.viewErrorMessage($"(ERROR) Algo ocurrió durante el proceso de ejecución.");
                    log.writeLog($"ALGO OCURRIÓ DURANTE EL PROCESO PRINCIPAL {ex.Message}");
                    log.writeLog($"**********************************************************************");
                    break;
                }

                gui.viewMainMessage("********FIN PROCESO********");
            }
        }
    }
}
