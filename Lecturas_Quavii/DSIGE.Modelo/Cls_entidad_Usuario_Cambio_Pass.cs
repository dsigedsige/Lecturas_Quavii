using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DSIGE.Modelo
{
   public  class Cls_entidad_Usuario_Cambio_Pass
    {

       public int id_user { get; set; }
       public string contrasenia_Actual { get; set; }
       public string contrasenia_anterior { get; set; }
       public string contrasenia_nueva { get; set; }
       public string contrasenia_nueva_confirma { get; set; }
    }
}
