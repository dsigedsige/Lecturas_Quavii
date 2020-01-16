using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DSIGE.Modelo;
using DSIGE.Dato;

namespace DSIGE.Negocio
{
  public  class NPrecios
    {
      public List<Precios> ListarPrecios()
      {
          return new DPrecios().ListarPrecios();
      }
      public bool RegistrarPrecios(Precios objPrecio)
      {
          return new DPrecios().RegistrarPrecios(objPrecio);
      }
      public bool ActualizarPrecios(Precios objPrecio)
      {
          return new DPrecios().ActualizarPrecios(objPrecio);
      }
    }
}
