using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DSIGE.Modelo
{
   public class Inspecciones_E
    {
        public int id_Inspeccion { get; set; }
        public int id_Operario_Lectura { get; set; }

        public string nro_Inspeccion { get; set; }
        public string fecha_inspeccion { get; set; }
        public string operario { get; set; }
 
    }
}
