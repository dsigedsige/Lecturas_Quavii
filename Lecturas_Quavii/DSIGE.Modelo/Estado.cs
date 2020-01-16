using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Newtonsoft;
using Newtonsoft.Json;

namespace DSIGE.Modelo
{
    public class Estado
    {
        [JsonProperty("_a")]
        public int idEstado { get; set; }

        [JsonProperty("_b")]
        public string descEstado { get; set; }
    }
}
