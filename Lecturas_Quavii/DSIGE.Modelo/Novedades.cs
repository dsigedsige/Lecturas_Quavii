using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DSIGE.Modelo
{
    public class Novedades
    {

        public bool checkeado { get; set; }
        public int id_foto { get; set; }
        public int  id_Operario { get; set; }
        public string desc_operario { get; set; }

        public string cuentaContrato { get; set; }
        public string urlfoto { get; set; }
        public string fechaRegistro { get; set; }
        public string Latitud { get; set; }
        public string Longitud { get; set; }     
    }
}
