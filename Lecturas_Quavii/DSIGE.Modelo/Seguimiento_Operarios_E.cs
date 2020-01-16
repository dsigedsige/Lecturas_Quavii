using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DSIGE.Modelo
{
   public  class Seguimiento_Operarios_E
    {
        public int id_operario { get; set; }
        public string operario { get; set; }
        public string latitud { get; set; }
        public string longitud { get; set; }

        public string fecha_gps { get; set; }
        public string lectura_movil { get; set; }
        public string fecha_lectura { get; set; }        

        public int id_lectura { get; set; }
        public string suministro { get; set; }
        public string cliente { get; set; }
        public string direccion { get; set; }
        public string distrito { get; set; }
        public int estado { get; set; }
        public string tiene_foto { get; set; }
    }
}
