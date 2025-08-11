using System;
using System.IO;
using System.Net;
using HTML = HtmlAgilityPack;
using Template_Tesoreria.Helpers.MangementLog;

namespace Template_Tesoreria.Helpers.Files
{
    public class PortalOracle
    {
        private Log _log;
        private ExecutionTimer _timer;
        private string _nmBank;

        public PortalOracle(string nmBank) 
        {
            this._log = new Log();
            this._timer = new ExecutionTimer();
            this._nmBank = nmBank;
        }

        public void setNmBank(string nmBank)
        {
            this._nmBank = nmBank;
        }

        public bool downloadTemplate()
        {
            this._log.writeLog($"(INFO) COMENZANDO CON LA DESCARGA DEL TEMPLATE");
            this._timer.startExecution();
            
            try
            {
                var client1 = new WebClient();
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

                this._log.writeLog($"(INFO) SE OBTUVO LA INFORMACIÓN PARA PODER DESCARGAR CORRECTAMENTE EL TEMPLATE");

                pathDirectory = Environment.GetFolderPath(Environment.SpecialFolder.UserProfile) + @"\\Downloads\\Templates";

                //Si no existe la Carpeta la creamos
                if (!Directory.Exists(pathDirectory)) Directory.CreateDirectory(pathDirectory);

                //Definimos la ruta donde guardaremos el archivo
                //http://www.oracle.com/webfolder/technetwork/docs/fbdi-25b/fbdi/xlsm/CashManagementBankStatementImportTemplate.xlsm                
                pathDestiny = Environment.GetFolderPath(Environment.SpecialFolder.UserProfile) + @"\\Downloads\\Templates\\CashManagementBankStatementImportTemplate_" + this._nmBank + ".xlsm";
                this._log.writeLog($"(INFO) EL TEMPLATE SE INSERTARÁ EN LA SIGUIENTE RUTA: {pathDestiny}");

                WebClient myWebClient = new WebClient();
                myWebClient.DownloadFile(urlFile, pathDestiny);

                this._log.writeLog($"(SUCCESS) SE DESCARGA EL TEMPLATE ||| TIEMPO DE EJECUCIÓN: {this._timer.endExecution()}");
                return true;
            }
            catch (Exception ex)
            {
                this._log.writeLog($"(ERROR) HUBO UN LIGERO ERROR AL QUERER DESCARGAR EL TEMPLATE DE ORACLE. NOS ARROJÓ: {ex.Message} ||| TIEMPO DE EJECUCIÓN: {this._timer.endExecution()}");
                return false;
            }
        }
    }
}
