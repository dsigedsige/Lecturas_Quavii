using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DSIGE.Modelo
{
   public class ImportarArchivo
    {




       public bool checkeado { get; set; }    
        public int id_Archivo { get; set; }      
        public int id_Empresa   { get; set; }       
        public int  id_Local      { get; set; }
        public string nombre_local { get; set; }
        public string nombre_tiposervicio { get; set; }  
       
        public int id_TipoServicio  { get; set; }    
        public int id_TipoArchivo    { get; set; }         
        public string cuentaContrato_Archivo   { get; set; }          
        public string nombre_TipoArchivo   { get; set; }  
        public string fechaRegistro_Archivo { get; set; }         
        public string nombre_Archivo { get; set; }        
        public string url_Archivo    { get; set; }       
        public int estado   { get; set; }
        public string ruta_archivo { get; set; }   




    }
}
