using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DSIGE.Modelo
{
   public  class Cls_Entidad_Parametro_Consumo
    {
        public int  id_parametro   { get; set; }
        public string   nombre     { get; set; }
        public double   valor      { get; set; }
        public int estado          { get; set; }


        public int usuario_creacion { get; set; }
        public string fecha_creacion { get; set; }
        public int usuario_edicion { get; set; }
        public string  fecha_edicion { get; set; }

        public string   nombre_usu_crea     { get; set; }
        public string   nombre_usu_modi     { get; set; }

         

    }
}
