using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using Template_Tesoreria.Models;

namespace Template_Tesoreria.Helpers.Network
{
    public class SharedDirectory
    {
        private string _ip;
        private string _svrUser;
        private string _svrPassword;

        public SharedDirectory(string ip)
        {
            this._ip = ip;
            this._svrUser = Environment.GetEnvironmentVariable("USER_REMOTE");
            this._svrPassword = Environment.GetEnvironmentVariable("PWD_REMOTE");
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
                
                NETRESOURCE nr = new NETRESOURCE
                {
                    dwType = 1, // Disk
                    lpRemoteName = networkPath
                };

                var result = WNetAddConnection2(ref nr, this._svrPassword, this._svrUser, 0);

                if(result == 0)
                {
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

                    foreach(var file in files)
                    {
                        var lstIndex = file.LastIndexOf(@"\");
                        var nameFile = file.Substring(lstIndex, (file.Length - lstIndex));

                        listFiles.Add(new MenuOption_Model() { ID = id.ToString(), Option = nameFile, Value = nameFile.Replace(@"\", "") });

                        id++;
                    }

                    WNetCancelConnection2(networkPath, 0, true);
                    return listFiles;
                }
                else
                {
                    return null;
                }
            }
            catch(Exception ex)
            {
                return null;
            }
        }
    }
}
