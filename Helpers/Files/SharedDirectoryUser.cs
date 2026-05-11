using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Management;
using System.Text;
using System.Threading.Tasks;
using Template_Tesoreria.Helpers.DataAccess;
using Template_Tesoreria.Models;

namespace Template_Tesoreria.Helpers.Files
{
    public class SharedDirectoryUser
    {
        private Log _log;
        private Crypto _crypto;
        private string _sharedPath;
        private const string FOLDER_NAME = "FormatosBancos";

        public SharedDirectoryUser()
        {
            this._log = new Log();
            this._crypto = new Crypto();
        }

        private void GetPath()
        {
            try
            {
                this._log.writeLog("(INFO) BUSCANDO LA EXISTENCIA DE LA CARPETA COMPARTIDA.");

                var searcher = new ManagementObjectSearcher("SELECT * FROM Win32_Share");
                
                foreach (ManagementObject share in searcher.Get())
                    if (string.Equals(share["Name"]?.ToString(), FOLDER_NAME, StringComparison.OrdinalIgnoreCase))
                    {
                        _sharedPath = share["Path"]?.ToString();
                        return;
                    }

                this._log.writeLog("(INFO) NO SE HA PODIDO ENCONTRAR LA CARPETA.");
            }
            catch (Exception ex)
            {
                this._log.writeLog($"(ERROR) FALLO AL OBTENER LA RUTA DE LA CARPETA COMPARTIDA. ERROR: {ex.Message}");
                throw;

            }
        }

        public List<MenuOption_Model> getFiles()
        {
            var listFiles = new List<MenuOption_Model>();

            try
            {
                GetPath();

                if (string.IsNullOrEmpty(this._sharedPath))
                    throw new Exception("No se ha encontrado la carpeta.");

                var id = 1;
                var files = new[] { "*.xls*", "*.txt" }
                    .SelectMany(pattern => Directory.GetFiles(_sharedPath, pattern))
                    .Where(file =>
                    {
                        var nameFile = Path.GetFileName(file);
                        var attributes = File.GetAttributes(file);

                        return !nameFile.StartsWith("~$") && !nameFile.StartsWith("\\") && (attributes & (FileAttributes.Hidden | FileAttributes.System)) == 0;
                    });

                this._log.writeLog($"(INFO) ARCHIVOS ENCONTRADOS: {files.Count()}");

                foreach(var file in files)
                {
                    var lstIndex = file.LastIndexOf(@"\");
                    var lstNameFile = file.Substring(lstIndex, (file.Length - lstIndex));

                    listFiles.Add(new MenuOption_Model() 
                    { 
                        ID = id.ToString(), 
                        Option = lstNameFile, 
                        Value = lstNameFile.Replace(@"\", "") 
                    });

                    id++;
                }

                listFiles.Add(new MenuOption_Model()
                {
                    ID = id.ToString(),
                    Option = "<-- REGRESAR A LA ELECCIÓN DE BANCO",
                    Value = "Regresar"
                });

                return listFiles;
            }
            catch(Exception ex)
            {
                this._log.writeLog($"(ERROR) FALLO AL RECUPERAR LA LISTA DE ARCHIVOS DENTRO DE LA CARPETA COMPARTIDA. ERROR: {ex.Message}");
                throw;
            }
        }
    }
}
