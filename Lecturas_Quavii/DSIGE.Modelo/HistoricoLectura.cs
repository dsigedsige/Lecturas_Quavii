using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DSIGE.Modelo
{
    public class HistoricoLectura
    {

        public string Suministro { get; set; }
        public string suministro_lectura { get; set; }  
        public string TipoServicio { get; set; }
        public string Observacion { get; set; }
        public string fechaAsignacion_lectura { get; set; }
        
        public string Lectura { get; set; }
        public string Mes { get; set; }
        public string consumo { get; set; }
        public string Orden { get; set; }

        public string id_Lectura { get; set; }
        public string foto { get; set; }
    }
}
