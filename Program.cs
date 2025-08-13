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
using System.Linq;

namespace Template_Tesoreria
{
    internal class Program
    {
        public static string getIP(Log log)
        {
            try
            {
                log.writeLog("(INFO) SE OBTIENE LA IP DEL USUARIO.");
                foreach (var ipv4 in Dns.GetHostEntry(Dns.GetHostName()).AddressList)
                    if (ipv4.AddressFamily == AddressFamily.InterNetwork)
                    {
                        log.writeLog("(SUCCESS) OBTENCIÓN DE IP CORRECTA");
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

        public static string errorInSomeProcess(string mssgFirstTry, string mssgMoreTries, int tryings, Log log)
        {
            do
            {
                log.writeLog("(INFO) HUBO ERROR EN ALGÚN PROCESO, SE LE PREGUNTARÁ AL USUARIO SI DESEA CONTINUAR");
                log.writeLog($"(INFO) EL USUARIO LLEVA {tryings}");

                var messageTry = tryings == 1 ? $"\n{mssgFirstTry}" : $"\n{mssgMoreTries}";

                Console.Write($"{messageTry}");
                var tryAgain = Console.ReadLine().Trim();

                if (tryings >= 2 && string.Equals(tryAgain, "s", StringComparison.CurrentCultureIgnoreCase))
                {
                    log.writeLog("(INFO) SE REDIRECCIONARÁ AL MENÚ PRINCIPAL");
                    return "PRINCIPIO";
                }
                else if (tryings >= 2 && string.Equals(tryAgain, "n", StringComparison.CurrentCultureIgnoreCase))
                {
                    Console.Write($"¿Deseas volver a intentarlo? [S/N]: ");
                    tryAgain = Console.ReadLine().Trim();

                    if (string.Equals(tryAgain, "n", StringComparison.CurrentCultureIgnoreCase))
                    {
                        log.writeLog("(INFO) EL USUARIO DECIDIÓ TERMINAR CON EL PROCESO");
                        return "NO";
                    }

                    log.writeLog("(INFO) SE REDIRECCIONARÁ A LA SELECCIÓN DE ARCHIVO");
                    return "ESCOGER";
                }

                if (tryings == 1 && string.Equals(tryAgain, "s", StringComparison.CurrentCultureIgnoreCase))
                    return "ESCOGER";
                else if (tryings == 1 && string.Equals(tryAgain, "n", StringComparison.CurrentCultureIgnoreCase))
                    return "NO";
            } while (true);
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
                new MenuOption_Model() { ID = "7", Option = "7. - BANORTE", Value = "Banorte" },
                new MenuOption_Model() { ID = "8", Option = "     SALIR", Value = "Salir" },
            };
            var ip = "";
            var nmBank = "";
            var pathDestiny = "";
            var id = 1;
            var tryings = 1;
            var exception = "";

            ConsoleKey key;

            while(true)
            {
                try
                {
                COMIENZO_PROCESO:
                    Console.Clear();

                    ip = getIP(log);
                    //var shrdDirectory = new SharedDirectory(ip);
                    var shrdDirectory = new SharedDirectory("10.128.10.19");
                    var filesMenu = shrdDirectory.getFiles();

                    log.writeLog("**COMENZANDO PROCESO**");

                    nmBank = gui.viewMenu("Extracto bancario ", "Por favor selecciona el banco de la siguiente lista para continuar:", options);

                    if(string.Equals(nmBank, "Salir", StringComparison.CurrentCultureIgnoreCase))
                    {
                        log.writeLog("(INFO) SE DESEÓ SALIR DEL APLICATIVO");
                        log.writeLog("**PROCESO TERMINADO**");
                        log.writeLog($"**********************************************************************");
                        return;
                    }

                ESCOGER_ARCHIVO:
                    if(filesMenu == null || filesMenu.Count <= 0)
                    {
                        ConsoleKey keyClose;
                        
                        gui.viewErrorMessage($"No se encontró ningún archivo excel dentro de la carpeta.");
                        gui.viewErrorMessage($@"Intente subir un archivo dentro de su carpeta \\{ip}\FormatosBancos e intente de nuevo.");
                        log.writeLog($"(WARNING) NO SE ENCONTRÓ NINGÚN ARCHIVO EXCEL DENTRO DE LA CARPETA COMPARTIDA.");
                        gui.viewInfoMessage("Presiona cualquier tecla para cerrar el aplicativo...");

                        keyClose = Console.ReadKey(true).Key;
                        return;
                    }

                    var nmFile = gui.viewMenu("Extracto bancario ", $"Por favor, selecciona el archivo con el que desea llenar el template. Se escogió el banco {nmBank.ToUpper()}:", filesMenu);

                    if(string.Equals(nmFile, "Regresar", StringComparison.CurrentCultureIgnoreCase))
                    {
                        log.writeLog("(INFO) SE DESEÓ REGRESAR AL MENÚ PRINCIPAL PARA ESCOGER OTRO BANCO");
                        goto COMIENZO_PROCESO;
                    }

                    gui.viewMainMessage("********COMENZANDO PROCESO********");

                    //Empezamos con la recolección de datos y el llenado de la información
                    var data = new List<TblTesoreria_Model>();
                    var spName = $"pa_Tesoreria_CargaExcel_{nmBank}";
                    var parameters = new Dictionary<string, object>()
                    {
                        { "@Ip", "10.128.10.19" },
                        { "@Excelname", nmFile }
                    };

                    gui.viewInfoMessage("*Extrayendo información para el llenado de template*");
                    Task.Run(() =>
                        {
                            data = dtService.GetDataList<TblTesoreria_Model>(cnn.DbTesoreria1019(), spName, parameters);
                            cts.Cancel();
                        }
                    );
                    gui.Spinner("Obteniendo...", cts.Token);
                    cts = new CancellationTokenSource();

                    var _parse = new TblTesoreria_Model();
                    _parse.parseDate(data);

                    if (data == null || data.Count == 0)
                    {
                        exception = errorInSomeProcess($"No se encontró ningún dato en el archivo {nmFile}. ¿Quieres escoger de nuevo el archivo? [S/N]: ", $"No se volvió a encontrar ningún dato en el archivo {nmFile}. ¿Quieres ir al menú principal? [S/N]: ", tryings, log);
                        tryings++;

                        switch(exception)
                        {
                            case "PRINCIPIO":
                                goto COMIENZO_PROCESO;

                            case "ESCOGER":
                                goto ESCOGER_ARCHIVO;

                            case "NO":
                                log.writeLog("**PROCESO TERMINADO**");
                                log.writeLog($"**********************************************************************");
                                return;
                        }
                    }

                    tryings = 1;

                DESCARGA:
                    gui.viewInfoMessage("*Descargando el template desde el sitio de Oracle*");

                    var rsltDownload = false;
                    var dwnld = new PortalOracle(nmBank);
                    Task.Run(() =>
                    {
                        rsltDownload = dwnld.downloadTemplate();
                        cts.Cancel();
                    });
                    gui.Spinner("Descargando...", cts.Token);
                    cts = new CancellationTokenSource();

                    if (!rsltDownload)
                    {
                        exception = errorInSomeProcess($"No se pudo descargar el template. ¿Quiere intentarlo de nuevo? [S/N]: ", $"No se pudo volver a descargar el template. ¿Quieres intentarlo de nuevo? [S/N]: ", tryings, log);

                        switch (exception)
                        {
                            case "PRINCIPIO":
                                Console.Clear();
                                goto DESCARGA;

                            case "ESCOGER":
                                Console.Clear();
                                goto DESCARGA;

                            case "NO":
                                log.writeLog("**PROCESO TERMINADO**");
                                log.writeLog($"**********************************************************************");
                                return;
                        }
                    }

                    gui.viewInfoMessage("*Limpiando template para su llenado*");
                    log.writeLog($"(INFO) LIMPIAMOS EL TEMPLATE PARA PODER INSERTAR LOS DATOS");
                    pathDestiny = Environment.GetFolderPath(Environment.SpecialFolder.UserProfile) + @"\\Downloads\\Templates\\CashManagementBankStatementImportTemplate_" + nmBank + ".xlsm";
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

                    log.writeLog($"(SUCCESS) TERMINO DE LIMPIEZA, SE PROSIGUE CON LA INSERCIÓN DE DATOS");

                    //var fillData = "";
                    var fillHeader = false;
                    var fillBalances = false;
                    var fillLines = false;

                    //Vamos a separar los dólares en otra lista
                    var lstUSD = data.ToList();
                    data.RemoveAll(x => x.Bank_Account_Currency != null && x.Bank_Account_Currency.Trim().Equals("USD", StringComparison.OrdinalIgnoreCase));
                    lstUSD.RemoveAll(x => x.Bank_Account_Currency != null && x.Bank_Account_Currency.Trim().Equals("MXN", StringComparison.OrdinalIgnoreCase));

                    gui.viewInfoMessage($"*Llenando template con los datos recuperados. Siendo un total de {data.Count} registros*");
                    Task.Run(() =>
                    {
                        /*
                         * Hay un escenario donde el extracto sólo tendrá USD, en ese caso
                         * vamos a pasarle la lista de dólares en lugar de la principal,
                         * esto para no hacer el proceso de nuevo.
                         */
                        fillHeader      = mngmntExcel.fillHeaderSheet(data ?? lstUSD);
                        fillBalances    = mngmntExcel.fillBalanceSheet(data ?? lstUSD);
                        fillLines       = mngmntExcel.fillLinesSheet(data ?? lstUSD);

                        cts.Cancel();
                    });
                    gui.Spinner("Llenando...", cts.Token);
                    cts = new CancellationTokenSource();

                    if(!fillHeader || !fillBalances || !fillLines)
                    {
                        gui.viewErrorMessage("(ERROR) Hubo un ligero error al querer llenar el template.");
                        break;
                    }

                    log.writeLog($"(INFO) SE COMPROBARÁ SI HAY DÓLARES DENTRO DE NUESTRA INFORMACIÓN");
                    
                    if((lstUSD != null && data != null) || (lstUSD.Count > 0 && data.Count > 0))
                    {
                        log.writeLog($"(INFO) HAY INFORMACIÓN DE DÓLARES DENTRO DE NUESTRA INFORMACIÓN");
                        log.writeLog($"(INFO) SE REINICIARÁ EL PROCESO PARA CARGAR LA INFORMACIÓN DE DÓLARES EN OTRO TEMPLATE");
                        dwnld.setNmBank($"{nmBank}_USD");
                        goto DESCARGA;
                    }

                    Console.Write("\n¿Desea llenar otro template? [S/N]: ");
                    var again = Console.ReadLine().Trim();
                    
                    Process.Start(pathDestiny);

                    if (string.Equals(again, "n", StringComparison.OrdinalIgnoreCase))
                        break;

                    log.writeLog($"(INFO) ABRIENDO ARCHIVO\n\t\t**PROCESO TERMINADO**");
                    log.writeLog($"**********************************************************************");
                }
                catch (Exception ex)
                {
                    gui.viewErrorMessage($"(ERROR) Algo ocurrió durante el proceso de ejecución.");
                    log.writeLog($"(ERROR) ALGO OCURRIÓ DURANTE EL PROCESO PRINCIPAL {ex.Message}");
                    log.writeLog($"**********************************************************************");
                    break;
                }

                gui.viewMainMessage("********FIN PROCESO********");
            }
        }
    }
}
