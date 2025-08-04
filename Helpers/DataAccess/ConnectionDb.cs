using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Template_Tesoreria.Helpers.DataAccess
{
    public class ConnectionDb
    {
        private Crypto _crypto;

        public ConnectionDb() 
        {
            this._crypto = new Crypto();
        }

        public string DbTesoreria1019()
        {
            var server = this._crypto.Decrypt(System.Configuration.ConfigurationManager.AppSettings["CnnServer"]);
            var bd = this._crypto.Decrypt(System.Configuration.ConfigurationManager.AppSettings["CnnBD"]);
            var user = this._crypto.Decrypt(System.Configuration.ConfigurationManager.AppSettings["CnnUser"]);
            var pass = this._crypto.Decrypt(System.Configuration.ConfigurationManager.AppSettings["CnnPwd"]);
            var cnn = string.Format("Data Source={0};Initial Catalog={1};Persist Security Info=True;User ID={2};Password={3}", server, bd, user, pass);
            return cnn;
        }
    }
}
