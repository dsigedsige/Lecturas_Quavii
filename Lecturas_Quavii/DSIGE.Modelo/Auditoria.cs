using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using Newtonsoft;
using Newtonsoft.Json;

namespace DSIGE.Modelo
{
    public class Auditoria
    {
        [JsonProperty("crea_id", NullValueHandling = NullValueHandling.Ignore)]
        public int crea_id { get; set; }

        [JsonProperty("crea_nombre", NullValueHandling = NullValueHandling.Ignore)]
        public string crea_nombre { get; set; }

        [JsonProperty("crea_fecha", NullValueHandling = NullValueHandling.Ignore)]
        public string crea_fecha { get; set; }

        [JsonProperty("modifica_id", NullValueHandling = NullValueHandling.Ignore)]
        public int modifica_id { get; set; }

        [JsonProperty("modifica_nombre", NullValueHandling = NullValueHandling.Ignore)]
        public string modifica_nombre { get; set; }

        [JsonProperty("modifica_fecha", NullValueHandling = NullValueHandling.Ignore)]
        public string modifica_fecha { get; set; }

        [JsonProperty("clase", NullValueHandling = NullValueHandling.Ignore)]
        public string clase { get; set; }
    }
}
