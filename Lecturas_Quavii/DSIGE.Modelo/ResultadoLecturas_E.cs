using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DSIGE.Modelo
{
    public class ResultadoLecturas_E
    {

        public int id_Operario { get; set; }
        public string desc_operario { get; set; }
        public string Sector { get; set; }
        public int estado { get; set; }
        
        public string nombre_tiposervicio { get; set; }
        public int id_TipoServicio { get; set; }  

        public string  total_lectura { get; set; }
        public string  porc_lectura { get; set; }
        public string  total_ejecutado { get; set; }
        public string  porc_ejecutado { get; set; }
        public string  total_pendiente { get; set; }
        public string  porc_pendiente { get; set; }
        public string  total_fotos { get; set; }
        public string porc_fotos { get; set; }

        public string total_no_visitados { get; set; }
        public string total_verificaciones { get; set; }
        public string total_verificacion_ejecu { get; set; }
        public string total_verificaciones_no_ejecu { get; set; }
        public string total_cliente_no_visitados { get; set; }

        public string anio { get; set; }
        public string mes { get; set; }

        public string codigo { get; set; }
        public string observacion { get; set; }
        public string total { get; set; }
        public string conFoto { get; set; }
        public string sinFoto  { get; set; }
        public string porcentaje { get; set; }      


       public string apellidos_operario { get; set; }    
       public string Total { get; set; }    
       public string Ejecutada { get; set; }    
       public string Pendiente { get; set; }    
       public string Foto { get; set; }    
       public string Porcentaje { get; set; }    
       public string Minimo { get; set; }    
       public string Maximo { get; set; }    
       public string totalHora { get; set; }
       public string FueraHora { get; set; }    

       public int  id_Lectura { get; set; }   
       public string suministro_lectura { get; set; }   
       public string medidor_lectura { get; set; }
       public string fotourl { get; set; }
       public string fecha { get; set; }   

        public string hora { get; set; } 
        public string latitud { get; set; } 
        public string longitud { get; set; }
        public string Sector_lectura { get; set; } 

    }
}
