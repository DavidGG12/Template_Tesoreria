using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Template_Tesoreria.Helpers.Db.Interface
{
    public interface IDatabaseManager
    {
        DataTable ExecuteStoredProcedure(string storedProcedureName, Dictionary<string, object> parameters);
    }
}
