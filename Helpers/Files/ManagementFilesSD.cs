using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using System.Text;
using System.Threading.Tasks;
using System.Runtime.InteropServices;
using System.Security.AccessControl;

namespace Template_Tesoreria.Helpers.Files
{
    public class ManagementFilesSD : IDisposable
    {
        string _networkName;

        public ManagementFilesSD(string networkName, NetworkCredential credentials)
        {
            _networkName = networkName;

            var netResource = new NetResource()
            {
                Scope = 2, 
                ResourceType = 1,
                DisplayType = 3,
                Usage = 1,
                RemoteName = networkName
            };

            var username = credentials.Domain != null ? $@"{credentials.Domain}\{credentials.UserName}" : credentials.UserName;
            var result = WNetAddConnection2(netResource, credentials.Password, username, 0);

            if(result == 1219)
            {
                WNetCancelConnection2(null, 0, true);
                result = WNetAddConnection2(netResource, credentials.Password, credentials.UserName, 0);
            }

            if (result != 0)
            {
                throw new InvalidOperationException("Error al conectar. Código: " + result);
            }
        }

        public void Dispose()
        {
            WNetCancelConnection2(_networkName, 0, true);
        }

        [DllImport("mpr.dll")]
        private static extern int WNetAddConnection2(NetResource netResource,
            string password, string username, int flags);

        [DllImport("mpr.dll")]
        private static extern int WNetCancelConnection2(string name, int flags, bool force);

        [StructLayout(LayoutKind.Sequential)]
        public class NetResource
        {
            public int Scope;
            public int ResourceType;
            public int DisplayType;
            public int Usage;
            public string LocalName;
            public string RemoteName;
            public string Comment;
            public string Provider;
        }
    }
}
