using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Newtonsoft;
using Newtonsoft.Json;

namespace DSIGE.Modelo
{
    public class Permisos
    {

        [JsonProperty("usu_id")]
        public string usu_id { get; set; }

        [JsonProperty("id_Perfil")]
        public string id_Perfil { get; set; }

        [JsonProperty("mod_id")]
        public string modu_id { get; set; }

        [JsonProperty("mod_des")]
        public string mod_des { get; set; }

        [JsonProperty("det_id")]
        public string det_id { get; set; }

        [JsonProperty("det_des")]
        public string det_des { get; set; }

        [JsonProperty("id_Opcion")]
        public string id_Opcion { get; set; }

        [JsonProperty("permisos_opciones")]
        public string permisos_opciones { get; set; }

        [JsonProperty("Estado")]
        public string Estado { get; set; }

        [JsonProperty("usuario_creacion")]
        public int usuario_creacion { get; set; }

        [JsonProperty("fecha_creacion")]
        public string fecha_creacion { get; set; }

        [JsonProperty("usuario_edicion")]
        public int usuario_edicion { get; set; }

        [JsonProperty("fecha_edicion")]
        public string fecha_edicion { get; set; }

        [JsonProperty("id_RegistroAcceso")]
        public string id_RegistroAcceso { get; set; }
     
   
    }
  
    public class parametersPermisos
    {
        [JsonProperty("con")]
        public string per_con { get; set; }

        [JsonProperty("id")]
        public string per_id { get; set; }

    }
    
    public class BarClass
    { 
        public string Sub_Property1 { get; set; }
        public string Sub_Property2 { get; set; }
        public BarClass() { }
        public BarClass (string one, string two)
        {
            Sub_Property1 = one;
            Sub_Property2 = two;
        }
    }

    public class Person
    {
        public string Name { get; set; }

        public string Address { get; set; }

        public int Age { get; set; }

        public List<Permisos> Detalle;
    }
 
       
}
