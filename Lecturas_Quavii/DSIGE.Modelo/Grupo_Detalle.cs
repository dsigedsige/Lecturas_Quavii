using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using Newtonsoft;
using Newtonsoft.Json;

namespace DSIGE.Modelo
{
    public class Grupo_Detalle
    {
        [JsonProperty("_a", NullValueHandling = NullValueHandling.Ignore)]
        public int gde_id { get; set; }

        [JsonProperty("_b", NullValueHandling = NullValueHandling.Ignore)]
        public int gru_id { get; set; }

        [JsonProperty("_c", NullValueHandling = NullValueHandling.Ignore)]
        public string gde_descripcion { get; set; }

        [JsonProperty("_d", NullValueHandling = NullValueHandling.Ignore)]
        public string gde_abreviatura { get; set; }

        [JsonProperty("_e", NullValueHandling = NullValueHandling.Ignore)]
        public int gde_orden { get; set; }

        [JsonProperty("_f", NullValueHandling = NullValueHandling.Ignore)]
        public int gde_estado { get; set; }
    }

    #region Request & Response

    /// <summary>
    /// Autor: jlucero
    /// Fecha: 2015-06-14
    /// </summary>
    public class Request_Grupo_Detalle_Select
    {
        [JsonProperty("_a")]
        public int gru_id { get; set; }
    }

    #endregion
}
