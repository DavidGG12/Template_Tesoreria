using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Template_Tesoreria.Models
{
    public class TblTesoreria_Model
    {
        public string Cuenta { get; set; }
        public string Fecha { get; set; }
        public string Referencia { get; set; }
        public string Informacion_Env { get; set; }
        public string Referencia_Ext { get; set; }
        public string Referencia_Leyenda { get; set; }
        public string Referencia_Numerica { get; set; }
        public string Concepto { get; set; }
        public string Movimiento { get; set; }
        public string Cargo { get; set; }
        public string Abono { get; set; }
        public string Ordenante { get; set; }
        public string RFC_Ordenante { get; set; }
        public string Saldo_Inicial { get; set; }
        public string Saldo_Final { get; set; }
        public string Moneda { get; set; }
    }
}
