using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Security.Cryptography;
using System.Text;
using System.Threading.Tasks;
using Template_Tesoreria.Helpers.Files;

namespace Template_Tesoreria.Helpers.DataAccess
{
    public class Crypto
    {
        private string _secretKey;
        private string _secretSalt;
        private Log _log;

        public Crypto()
        {
            this._secretKey = Environment.GetEnvironmentVariable("SECRET_KEY");
            this._secretSalt = Environment.GetEnvironmentVariable("SECRET_SALT");
            this._log = new Log();
        }

        public string Encrypt(string plainText)
        {
            var error = "";

            if (string.IsNullOrEmpty(this._secretKey))
                error = error + "NO SE HA PODIDO LEER LA VARIABLE DE ENTORNO: SECRET_KEY \n";

            if (string.IsNullOrEmpty(this._secretKey))
                error = error + "NO SE HA PODIDO LEER LA VARIABLE DE ENTORNO: SECRET_SALT \n";

            if(string.IsNullOrEmpty(error.Trim()))
            {
                try
                {
                    this._log.writeLog("(INFO) COMENZAMOS EL PROCESO DE ENCRIPTACIÓN");

                    using (Aes aes = Aes.Create())
                    {
                        byte[] saltBytes = Encoding.UTF8.GetBytes(this._secretSalt);
                        var key = new Rfc2898DeriveBytes(this._secretKey, saltBytes, 10000);
                        aes.Key = key.GetBytes(32);
                        aes.IV = key.GetBytes(16);

                        var encryptor = aes.CreateEncryptor(aes.Key, aes.IV);
                        using (var ms = new MemoryStream())
                        {
                            using (var cs = new CryptoStream(ms, encryptor, CryptoStreamMode.Write))
                            using (var sw = new StreamWriter(cs))
                                sw.Write(plainText);

                            this._log.writeLog($"(SUCCESS) SE ENCRIPTÓ EL TEXTO CORRECTAMENTE, REGRESANDO TEXTO ENCRIPTADO.");
                            return Convert.ToBase64String(ms.ToArray());
                        }
                    }
                }
                catch(Exception ex)
                {
                    this._log.writeLog($"(ERROR) HUBO UN ERROR AL QUERER ENCRIPTAR EL TEXTO. NOS ARROJA: {ex.Message}");
                    return null;
                }
            }

            this._log.writeLog($"(ERROR) {error}");
            return null;
        }

        public string Decrypt(string cipherText)
        {
            var error = "";

            if (string.IsNullOrEmpty(this._secretKey))
                error = error + "NO SE HA PODIDO LEER LA VARIABLE DE ENTORNO: SECRET_KEY \n";

            if (string.IsNullOrEmpty(this._secretKey))
                error = error + "NO SE HA PODIDO LEER LA VARIABLE DE ENTORNO: SECRET_SALT \n";

            if(string.IsNullOrEmpty(error.Trim()))
            {
                try
                {
                    this._log.writeLog("(INFO) COMENZAMOS EL PROCESO DE ENCRIPTACIÓN");

                    using (Aes aes = Aes.Create())
                    {
                        byte[] saltBytes = Encoding.UTF8.GetBytes(this._secretSalt);
                        var key = new Rfc2898DeriveBytes(this._secretKey, saltBytes, 10000);
                        aes.Key = key.GetBytes(32);
                        aes.IV = key.GetBytes(16);

                        byte[] buffer = Convert.FromBase64String(cipherText);

                        var decryptor = aes.CreateDecryptor(aes.Key, aes.IV);
                        using (var ms = new MemoryStream(buffer))
                        using (var cs = new CryptoStream(ms, decryptor, CryptoStreamMode.Read))
                        using (var sr = new StreamReader(cs))
                        {
                            this._log.writeLog($"(SUCCESS) SE DESENCRIPTÓ EL TEXTO CORRECTAMENTE, REGRESANDO TEXTO DESENCRIPTADO.");
                            return sr.ReadToEnd();
                        }
                    }
                }
                catch(Exception ex)
                {
                    this._log.writeLog($"(ERROR) HUBO UN ERROR AL QUERER DESENCRIPTAR EL TEXTO. NOS ARROJA: {ex.Message}");
                    return null;
                }
            }

            this._log.writeLog($"(ERROR) {error}");
            return null;
        }
    }
}
