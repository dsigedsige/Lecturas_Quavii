using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DSIGE.Modelo
{
 public   class Precios
    {

     public int id_Precio { get; set; }

     public string Precio { get; set; }

     public string Concepto { get; set; }

     public string anio { get; set; }

     public string id_Local { get; set; }

     public string nombre_local { get; set; }

     public string estado { get; set; }

     public int usuario_creacion { get; set; }

     public int usuario_edicion { get; set; }

    }
}
