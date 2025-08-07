using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using Template_Tesoreria.Helpers.DataAccess;
using Template_Tesoreria.Helpers.Files;
using Template_Tesoreria.Models;

namespace Template_Tesoreria.Helpers.Network
{
    public class SharedDirectory
    {
        private Log _log;
        private Crypto _crypto;

        private string _ip;
        private string _svrUser;
        private string _svrPassword;

        public SharedDirectory(string ip)
        {
            this._log = new Log();

            this._crypto = new Crypto();

            this._ip = ip;
            this._svrUser = this._crypto.Decrypt(System.Configuration.ConfigurationManager.AppSettings["SvrUser"]);
            this._svrPassword = this._crypto.Decrypt(System.Configuration.ConfigurationManager.AppSettings["SvrPwd"]);
        }

        public void setIP(string ip)
        {
            this._ip = ip;
        }

        [DllImport("mpr.dll")]
        private static extern int WNetAddConnection2(ref NETRESOURCE netResource, string password, string username, int flags);

        [DllImport("mpr.dll")]
        private static extern int WNetCancelConnection2(string name, int flags, bool force);

        [StructLayout(LayoutKind.Sequential)]
        private struct NETRESOURCE
        {
            public int dwScope;
            public int dwType;
            public int dwDisplayType;
            public int dwUsage;
            public string lpLocalName;
            public string lpRemoteName;
            public string lpComment;
            public string lpProvider;
        }

        public List<MenuOption_Model> getFiles()
        {
            var listFiles = new List<MenuOption_Model>();

            try
            {

                var networkPath = $@"\\{this._ip}\FormatosBancos";
                var erExcel = @".xlsx|.xls";
                
                this._log.writeLog("(INFO) EMPEZAMOS LA CONEXIÓN CON LA CARPETA COMPARTIDA.");
                
                NETRESOURCE nr = new NETRESOURCE
                {
                    dwType = 1, // Disk
                    lpRemoteName = networkPath
                };

                var result = WNetAddConnection2(ref nr, this._svrPassword, this._svrUser, 0);

                if(result == 0)
                {
                    this._log.writeLog("(SUCCESS) CONEXIÓN EXITOSA");

                    var id = 1;
                    var files = Directory.GetFiles(networkPath, "*.xls*").Where(f =>
                    {
                        var nombre = Path.GetFileName(f);
                        var atributos = File.GetAttributes(f);

                        // Ignorar ocultos, de sistema y temporales de Office
                        return !nombre.StartsWith("~$") &&
                               !nombre.StartsWith("\\") &&
                               (atributos & (FileAttributes.Hidden | FileAttributes.System)) == 0;
                    });

                    this._log.writeLog($"(INFO) ARCHIVOS ENCONTRADOS: {files.Count()}");
                    
                    foreach(var file in files)
                    {
                        var lstIndex = file.LastIndexOf(@"\");
                        var nameFile = file.Substring(lstIndex, (file.Length - lstIndex));

                        listFiles.Add(new MenuOption_Model() { ID = id.ToString(), Option = nameFile, Value = nameFile.Replace(@"\", "") });

                        id++;
                    }

                    this._log.writeLog($"(SUCCESS) SE REGRESA EL LISTADO DE ARCHIVOS");

                    listFiles.Add(new MenuOption_Model() { ID = id.ToString(), Option = "<-- REGRESAR A LA ELECCIÓN DE BANCO", Value = "Regresar" });

                    WNetCancelConnection2(networkPath, 0, true);
                    return listFiles;
                }
                else
                {
                    this._log.writeLog($"(ERROR) FALLO AL ESTABLECER CON LA CONEXIÓN DE LA CARPETA COMPARTIDA. CÓDIGO DE ERROR: {result}");
                    return null;
                }
            }
            catch(Exception ex)
            {
                this._log.writeLog($"(ERROR) FALLO AL ESTABLECER CON LA CONEXIÓN DE LA CARPETA COMPARTIDA. EXEPCIÓN: {ex.Message}");
                return null;
            }
        }
    }
}
