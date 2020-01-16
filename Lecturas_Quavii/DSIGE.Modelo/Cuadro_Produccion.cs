using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DSIGE.Modelo
{
  public   class Cuadro_Produccion
    {

        public int id_Operario { get; set; }

        public string usuario_Operario { get; set; }

        public string nombre_Operario { get; set; }

        public string ing_Operario { get; set; }

        public string cargo_Operario { get; set; }

        public string Abre_Operario { get; set; }

        public decimal Precio { get; set; }

        public decimal PrecioUnit { get; set; }

        public string fechaAsignacion_Lectura { get; set; }

        public int cantidad { get; set; }

    }
}
