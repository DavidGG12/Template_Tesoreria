using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Template_Tesoreria.Helpers.Db;
using Template_Tesoreria.Helpers.Files;
using static Template_Tesoreria.Helpers.Db.DbFactory;

namespace Template_Tesoreria.Helpers.DataAccess
{
    public class DataService
    {
        private Log log = new Log();
        private T MapDataRowToModel<T>(DataRow row) where T : new()
        {
            var model = new T();
            var properties = typeof(T).GetProperties();

            foreach (var property in properties)
            {
                if (row.Table.Columns.Contains(property.Name) && !row.IsNull(property.Name))
                    property.SetValue(model, Convert.ChangeType(row[property.Name], property.PropertyType));
            }

            return model;
        }

        private List<T> MapDataList<T>(DataTable table) where T : new()
        {
            var modelList = new List<T>();
            var properties = typeof(T).GetProperties();

            foreach (DataRow row in table.Rows)
            {
                var model = new T();
                foreach (var property in properties)
                {
                    if (row.Table.Columns.Contains(property.Name) && !row.IsNull(property.Name))
                        property.SetValue(model, Convert.ChangeType(row[property.Name], property.PropertyType));
                }

                modelList.Add(model);
            }

            return modelList;
        } 

        public T GetData<T>(string conString, string storedName, Dictionary<string, object> parameters) where T : new()
        {
            var dbFactory = DatabaseManagerFactory.CreateDatabaseManager(conString);
            var result = dbFactory.ExecuteStoredProcedure(storedName, parameters);

            if (result.Rows.Count > 0)
                return MapDataRowToModel<T>(result.Rows[0]);

            return default;
        }

        public List<T> GetDataList<T>(string conString, string storedName, Dictionary<string, object> parameters) where T : new()
        {
            log.writeLog($"SE HARÁ LA CONEXIÓN CON LA BASE DE DATOS\n\t\tSE EJECUTARÁ EL STORED PROCEDURE: {storedName}");

            var dbFactory = DatabaseManagerFactory.CreateDatabaseManager(conString);
            var result = dbFactory.ExecuteStoredProcedure(storedName, parameters);

            if (result.Rows.Count > 0)
            {
                log.writeLog($"DEVOLVIENDO DATOS OBTENIDOS DEL STORED PROCEDURE");
                return MapDataList<T>(result);
            }

            log.writeLog($"NO SE OBTUVIERON DATOS DEL STORED PROCEDURE");
            return default;
        }
    }
}
