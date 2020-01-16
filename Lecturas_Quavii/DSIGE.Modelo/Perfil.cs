using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Newtonsoft;
using Newtonsoft.Json;


namespace DSIGE.Modelo
{
   public class Perfil
    {

        [JsonProperty("_a")]
        public int id_Perfil { get; set; }

        [JsonProperty("_b")]
        public string nombre_perfil { get; set; }

        [JsonProperty("_c")]
        public string descripcion_perfil { get; set; }

        [JsonProperty("_d")]
        public string estado { get; set; }

        [JsonProperty("_e")]
        public int usuario_creacion { get; set; }

        [JsonProperty("_f")]
        public int usuario_edicion { get; set; }

        [JsonProperty("_g")]
        public string cheked { get; set; }

        


    }
}
