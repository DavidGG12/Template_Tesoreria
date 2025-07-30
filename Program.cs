using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using HTML = HtmlAgilityPack;
using Microsoft.Office.Interop.Excel;
using System.Data;
using System.IO;
using System.Net;
using Template_Tesoreria.Helpers.DataAccess;
using Template_Tesoreria.Models;
using Template_Tesoreria.Helpers.Files;
using Template_Tesoreria.Helpers.ProcessExe;
using System.Diagnostics;
using System.Threading;
using System.Net.Sockets;

namespace Template_Tesoreria
{
    internal class Program
    {
        static void Spinner(string message, CancellationToken token)
        {
            char[] secuence = { '|', '/', '-', '\\' };
            int pos = 0;

            Console.ForegroundColor = ConsoleColor.Green;

            while (!token.IsCancellationRequested)
            {
                Console.Write($"\r{message} {secuence[pos]}");
                pos = (pos + 1) % secuence.Length;
                Thread.Sleep(100);
            }

            Console.ResetColor();

            Console.Write($"\r{new string(' ', Console.WindowWidth)}");
            Console.Write("\rTerminado\n");
        }

        static void wrtFooter(string txt)
        {
            int rowFooter = Console.WindowHeight - 1;
            Console.SetCursorPosition(0, rowFooter);
            Console.BackgroundColor = ConsoleColor.Gray;
            Console.ForegroundColor= ConsoleColor.Black;
            Console.WriteLine(txt.PadRight(Console.WindowWidth));
            Console.ResetColor();
        }

        public static string menuFiles(string ip)
        {
            var path = @"\\10.128.10.19\FormatoBancos";
            var credentials = new NetworkCredential("svrsafin", "Lp5vr5a71n", "sanborns");
            ConsoleKey key;
            
            string[] files;

            using (var mngFile = new ManagementFilesSD(path, credentials))
            {
                files = Directory.GetFiles(path);
            }

            var selection = 0;

            do
            {
                Console.Clear();

                for(int i = 0; i < files.Length; i++)
                {
                    if(i == selection)
                    {
                        Console.BackgroundColor = ConsoleColor.Gray;
                        Console.ForegroundColor = ConsoleColor.Black;
                    }

                    Console.WriteLine(files[i]);
                    Console.ResetColor();
                }

                key = Console.ReadKey(true).Key;

                if (key == ConsoleKey.UpArrow)
                    selection = (selection == 0) ? files.Length - 1 : selection - 1;
                else if (key == ConsoleKey.DownArrow)
                    selection = (selection == files.Length - 1) ? 0 : selection + 1;
            }
            while (key != ConsoleKey.Enter);

            return files[selection];
        }

        static void Main(string[] args)
        {
            var dtService = new DataService();
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
            string opc = "", opc2 = "", nombreBanco = "", rutaCarpeta = "", urlArchivoDescaga = "", pathDestino = "";
            var id = 1;

            ConsoleKey key;

            try
            {
                log.writeLog("COMENZANDO PROCESO");

                
                #region MENU
                do
                {
                    Console.Clear();
                    log.writeLog("IMPRESIÓN DEL MENÚ");

                    Console.Title = "Template Tesoreria";
                    Console.ForegroundColor = ConsoleColor.Cyan;

                    Console.WriteLine("╔════════════════════════════════════════════════════╗");
                    Console.WriteLine("║                 TEMPLATE  TESORERIA                ║");
                    Console.WriteLine("║                                                    ║");
                    Console.WriteLine("║  Por favor selecciona el banco de la siguiente     ║");
                    Console.WriteLine("║  lista para continuar:                             ║");
                    Console.WriteLine("╚════════════════════════════════════════════════════╝\n");

                    Console.ResetColor();

                    Console.WriteLine("Selecciona la compañía que deseas generar:\n");

                    foreach (var option in options)
                    {
                        if(id.ToString() == option.ID)
                        {
                            Console.BackgroundColor = ConsoleColor.Gray;
                            Console.ForegroundColor = ConsoleColor.Black;
                            opc = option.Option;
                        }
                        Console.WriteLine(option.Option);
                        Console.ResetColor();
                    }

                    //wrtFooter("[UpArrow]/[DownArrow] Navegar  |  [Enter] Seleccionar");

                    key = Console.ReadKey(true).Key;

                    var chsOpt = options.Find(x => x.Option.Contains(opc));

                    if(key == ConsoleKey.UpArrow || key == ConsoleKey.W)
                        id = (id == 1) ? 1 : int.Parse(chsOpt.ID) - 1;
                    else if(key == ConsoleKey.DownArrow || key == ConsoleKey.S)
                        id = (id == options.Count) ? options.Count : int.Parse(chsOpt.ID) + 1; 

                    if (key == ConsoleKey.Enter)
                    {
                        nombreBanco = chsOpt.Value;
                        Console.Write($"\n¿Está seguro de querer trabajar con {nombreBanco}? [S/N]: ");
                        opc2 = Console.ReadLine().Trim();
                        if (opc2.Equals("s", StringComparison.OrdinalIgnoreCase))
                        {
                            log.writeLog($"SE CONFIRMA EL USO DEL BANCO {nombreBanco}");
                            Console.Clear();
                            break;
                        }
                        else
                            key = ConsoleKey.UpArrow;
                        Console.Clear();
                    }

                } while (key != ConsoleKey.Enter);
                #endregion

                #region Proceso
                Console.Write("\nComenzando proceso.\n\n");

                #region Descarga Template
                WebClient client1 = new WebClient();
                string htmlCode = client1.DownloadString("https://docs.oracle.com/en/cloud/saas/financials/25b/oefbf/cashmanagementbankstatementdataimport-3168.html#cashmanagementbankstatementdataimport-3168");
                string[] lines = htmlCode.Split('\n');

                HTML.HtmlDocument htmlDocument = new HTML.HtmlDocument();
                htmlDocument.LoadHtml(lines[58].ToString().Trim());

                var linkNodes = htmlDocument.DocumentNode.SelectNodes("//a[@href]");

                if (linkNodes != null)
                    foreach (var linkNode in linkNodes)
                        urlArchivoDescaga = linkNode.GetAttributeValue("href", string.Empty);

                log.writeLog($"SE OBTUVO LA INFORMACIÓN PARA PODER DESCARGAR CORRECTAMENTE EL TEMPLATE");

                rutaCarpeta = Environment.GetFolderPath(Environment.SpecialFolder.UserProfile) + @"\\Documents\\Templates";

                //Si no existe la Carpeta la creamos
                if (!Directory.Exists(rutaCarpeta)) Directory.CreateDirectory(rutaCarpeta);


                //Definimos la ruta donde guardaremos el archivo
                //http://www.oracle.com/webfolder/technetwork/docs/fbdi-25b/fbdi/xlsm/CashManagementBankStatementImportTemplate.xlsm                
                pathDestino = Environment.GetFolderPath(Environment.SpecialFolder.UserProfile) + @"\\Documents\\Templates\\CashManagementBankStatementImportTemplate_" + nombreBanco + ".xlsm";
                var mngmntExcel = new ManagementExcel(pathDestino, nombreBanco);

                mngmntExcel.closeDocument();

                log.writeLog($"EL TEMPLATE SE INSERTARÁ EN LA SIGUIENTE RUTA: {pathDestino}");
                
                WebClient myWebClient = new WebClient();
                myWebClient.DownloadFile(urlArchivoDescaga, pathDestino);

                Console.Write("\nTemplate Descargado.\n\n");
                Console.Write("\nSe insertan los datos.\n\n");

                log.writeLog($"SE DESCARGA EL TEMPLATE");
                log.writeLog($"EMPIEZA LA INSERCIÓN DE LOS DATOS EN EL TEMPLATE");
                #endregion

                #region Obtención de IP
                //Obtenemos la ip del usuario
                Console.Write("\nSe obtiene la IP del usuario.\n\n");
                var ip = "";

                foreach(var ipv4 in Dns.GetHostEntry(Dns.GetHostName()).AddressList)
                {
                    if(ipv4.AddressFamily == AddressFamily.InterNetwork)
                    {
                        ip = ipv4.ToString();
                        break;
                    }
                }

                Console.Write($"\nSe trabajará con la IP: {ip}\n\n");
                #endregion


                //var file = menuFiles(ip);
                //var file = menuFiles("10.128.10.19");
                //var file = "INBURSA EJEMPLO EXTRACTO BANCARIO 100725 M.N.xlsx";

                var valueFile = new ValueFile_Model();

                switch (nombreBanco)
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

                #region Inserción de Datos en Template
                //Empezamos con la recolección de datos y el llenado de la información
                var data = new List<TblTesoreria_Model>();
                var parameters = new Dictionary<string, object>()
                {
                    { "@Ip", "10.128.10.19" },
                    { "@Excelname", valueFile.FileName }
                };

                Console.Write($"\nObteniendo los datos que se insertaran en el template.\n\n");

                Task.Run(() =>
                    {
                        data = dtService.GetDataList<TblTesoreria_Model>(cnn.DbTesoreria1019(), valueFile.SPName, parameters);
                        cts.Cancel();
                    }
                );

                Spinner("Procesando...", cts.Token);


                //Limpiamos el template para trabajar con él
                log.writeLog($"LIMPIAMOS EL TEMPLATE PARA PODER INSERTAR LOS DATOS");
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
                    Console.WriteLine($"Hubo un ligero error al querer limpiar los datos de la hoja {error.Sheet}.\nError: {error.Message}");
                    log.writeLog($"**********************************************************************");
                    return;
                }

                log.writeLog($"TERMINO DE LIMPIEZA, SE PROSIGUE CON LA INSERCIÓN DE DATOS");

                //Insertamos los datos que se encuentran en la base de datos
                var fillData = mngmntExcel.getTemplate(data);

                Console.Write("Template de Oracle llenado con éxito.\n\n");
                #endregion

                Console.Write("\nPresiona cualquier tecla para salir...");
                Console.ReadKey();

                Process.Start(pathDestino);
                log.writeLog($"ABRIENDO ARCHIVO\n\t\t**PROCESO TERMINADO**");
                log.writeLog($"**********************************************************************");


                //Proceso para Leer Formato de Banco
                //UploadFile("");
                #endregion
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
                log.writeLog($"ALGO OCURRIÓ DURANTE EL PROCESO PRINCIPAL {ex.Message}");
                log.writeLog($"**********************************************************************");
            }
        }

        public static void UploadFile(string rutaFormato)
        {
            try
            {
                FtpWebRequest request = (FtpWebRequest)WebRequest.Create(System.Configuration.ConfigurationManager.AppSettings.Get("FtpRuta"));
                request.Method = WebRequestMethods.Ftp.UploadFile;
                request.Credentials = new NetworkCredential(System.Configuration.ConfigurationManager.AppSettings.Get("FtpUser"), System.Configuration.ConfigurationManager.AppSettings.Get("FtpPass"));

                byte[] fileContents = File.ReadAllBytes(rutaFormato);
                request.ContentLength = fileContents.Length;

                using (Stream requestStream = request.GetRequestStream())
                {
                    requestStream.Write(fileContents, 0, fileContents.Length);
                }

                using (FtpWebResponse response = (FtpWebResponse)request.GetResponse())
                {
                    Console.WriteLine($"Upload File Complete, status {response.StatusDescription}");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }

        }
    }
}
