using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using Newtonsoft;
using Newtonsoft.Json;

namespace DSIGE.Modelo
{
 public   class Cargo_Personal
    {

        [JsonProperty("_a")]
        public int id_Cargo { get; set; }

        [JsonProperty("_b")]
        public string nombre_cargo { get; set; }

        [JsonProperty("_c")]
        public int estado { get; set; }

    }
}
