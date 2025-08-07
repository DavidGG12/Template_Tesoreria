using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Template_Tesoreria.Helpers.Files;

namespace Template_Tesoreria.Helpers.DataAccess
{
    public class ConnectionDb
    {
        private Crypto _crypto;
        private Log _log;

        public ConnectionDb() 
        {
            this._crypto = new Crypto();
            this._log = new Log();
        }

        public string DbTesoreria1019()
        {
            try
            {
                var error = "";

                this._log.writeLog("(INFO) INICIANDO CONSTRUCCIÓN DE LA CADENA DE CONEXIÓN PARA LA BASE DE DATOS");

                var server = this._crypto.Decrypt(System.Configuration.ConfigurationManager.AppSettings["CnnServer"]);
                var bd = this._crypto.Decrypt(System.Configuration.ConfigurationManager.AppSettings["CnnBD"]);
                var user = this._crypto.Decrypt(System.Configuration.ConfigurationManager.AppSettings["CnnUser"]);
                var pass = this._crypto.Decrypt(System.Configuration.ConfigurationManager.AppSettings["CnnPwd"]);

                if (string.IsNullOrEmpty(server))
                    error = error + "SERVIDOR VACÍO. ";
                if (string.IsNullOrEmpty(bd))
                    error = error + "NOMBRE DE BASE DE DATOS VACÍA. ";
                if (string.IsNullOrEmpty(user))
                    error = error + "USUARIO VACÍO. ";
                if (string.IsNullOrEmpty(pass))
                    error = error + "CONTRASEÑA VACÍA. ";

                if(string.IsNullOrEmpty(error.Trim()))
                {
                    var cnn = string.Format("Data Source={0};Initial Catalog={1};Persist Security Info=True;User ID={2};Password={3}", server, bd, user, pass);
                    this._log.writeLog("(SUCCESS) CADENA CONSTRUIDA CORRECTAMENTE, REGRESANDO CADENA");
                    return cnn;
                }

                this._log.writeLog($"(ERROR) {error}");
                return null;
            }
            catch(Exception ex)
            {
                this._log.writeLog($"(ERROR) HUBO UN ERROR AL QUERER CONSTRUIR LA CADENA DE CONEXIÓN. NOS ARROJA: {ex.Message}");
                return null;
            }
        }
    }
}
