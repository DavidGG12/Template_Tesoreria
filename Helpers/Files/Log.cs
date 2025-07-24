using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Template_Tesoreria.Helpers.Files
{
    public class Log
    {
        private string _pathLogSave;
        public Log() 
        {
            var execDate = DateTime.Now.ToString("yyyyMMdd");
            var dirLog = $@"{Environment.GetFolderPath(Environment.SpecialFolder.UserProfile)}\\Documents\\Templates_Tesoreria";

            if(!Directory.Exists(dirLog)) Directory.CreateDirectory(dirLog);

            this._pathLogSave = $@"{dirLog}\\TemplatesTesoreria_Log_{execDate}.txt";

            if(!File.Exists(this._pathLogSave))
            {
                try
                {
                    using (StreamWriter wrt = File.CreateText(this._pathLogSave))
                    {
                        wrt.WriteLine($"**************************** LOG TESORERIA ****************************");
                    }
                }
                catch(Exception ex)
                {
                    Console.WriteLine($"Hubo un error al crear el log de la aplicación.\nError: {ex.ToString()}");
                }
            }
        }

        public void writeLog(string message)
        {
            var date = DateTime.Now.ToString("dd/MM/yyyy HH:mm:ss");

            try
            {
                File.AppendAllText(this._pathLogSave, $"{date} | {message}\n");
            }
            catch (Exception ex)
            {
                var error = ex.Message;
            }
        }
    }
}
