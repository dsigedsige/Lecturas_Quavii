using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DSIGE.Modelo
{
   public class Cls_Entidad_dias_trabajo
    {

       public int id_DiaTrabajo { get; set; }
       public int id_Operario { get; set; }
       public string desc_operario { get; set; }
       public string NombreDia { get; set; }
       public string  HoraEntrada { get; set; }
       public string HoraSalida { get; set; }
       public int estado { get; set; }
       public int usuario_creacion { get; set; }
       public int id_Local { get; set; }
       public string nombre_local { get; set; }
       public int usuario_edicion { get; set; }

       public string nombre_usu_crea { get; set; }       
       public string fecha_creacion { get; set; }       
       public string nombre_usu_modi { get; set; }
       public string fecha_edicion { get; set; }

    }
}
