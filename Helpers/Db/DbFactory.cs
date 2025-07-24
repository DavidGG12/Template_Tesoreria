using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Template_Tesoreria.Helpers.Db.Interface;
using Template_Tesoreria.Helpers.Files;

namespace Template_Tesoreria.Helpers.Db
{
    public class DbFactory : IDatabaseManager
    {
        private readonly string _connectionString;
        private Log log = new Log();

        public DbFactory(string connectionString)
        {
            log.writeLog($"SE CAPTURA LA CADENA DE CONEXIÓN A LA BD");
            _connectionString = connectionString;
        }

        public DataTable ExecuteStoredProcedure(string storedProcedureName, Dictionary<string, object> parameters)
        {
            log.writeLog($"COMENZANDO CON LA EJECUCIÓN DEL STORED PROCEDURE");

            DataTable resultTable = new DataTable();

            try
            {
                using (var con = new SqlConnection(_connectionString))
                {
                    using(var cmd = new SqlCommand(storedProcedureName, con))
                    {
                        cmd.CommandType = CommandType.StoredProcedure;

                        if(parameters != null)
                        {
                            foreach(var param in parameters)
                            {
                                if (param.Value is string)
                                    cmd.Parameters.AddWithValue(param.Key, param.Value.ToString().Replace("'", "") ?? "");
                                else
                                    cmd.Parameters.AddWithValue(param.Key, param.Value);
                            }
                        }

                        con.Open();

                        using (var adapter = new SqlDataAdapter(cmd))
                        {
                            log.writeLog($"EJECUCIÓN DEL STORED PROCEDURE {storedProcedureName} EXITOSA");
                            adapter.Fill(resultTable);
                            return resultTable;
                        }
                    }
                }
            }
            catch(SqlException sqlEx)
            {
                log.writeLog($"ALGO MALO OCURRIÓ AL QUERER EJECUTAR EL STORED PROCEDURE {storedProcedureName}\n\tERROR: {sqlEx.Message}");
                return default;
            }
            catch(Exception ex)
            {
                log.writeLog($"ALGO MALO OCURRIÓ\n\tERROR: {ex.Message}");
                return default;
            }
        }

        public class DatabaseManagerFactory
        {
            public static IDatabaseManager CreateDatabaseManager(string connectionString)
            {
                return new DbFactory(connectionString);
            }
        }
    }
}
