using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DSIGE.Modelo
{
   public class Cls_Entidad_Evento_Movil_Operario
    {


           
   //    FECHA, id_Operario, OPERARIO, GPS, BATERIA, DATOS, MODO_AVION
 


            public int id_Local { get; set; }
            public string nombre_local { get; set; }
            public int id_Operario { get; set; }
            public string Nombre_Operario { get; set; }


 

            public string FECHA { get; set; }
            public string HORA { get; set; }
            public string GPS { get; set; }
            public string BATERIA { get; set; }
            public string DATOS { get; set; }
            public string MODO_AVION { get; set; }
            public string latitud_lectura { get; set; }
            public string longitud_lectura { get; set; }

            public string fecha_inicio_recorrido { get; set; }
            public string fecha_fin_recorrido{ get; set; }
            public int  cantidad_paradas { get; set; }





 


    }
}
