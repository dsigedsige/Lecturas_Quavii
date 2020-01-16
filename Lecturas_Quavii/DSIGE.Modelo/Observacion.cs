using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using Newtonsoft;
using Newtonsoft.Json;

namespace DSIGE.Modelo
{
    public class Observacion : Auditoria
    {
        [JsonProperty("_a")]
        public int obs_id { get; set; }

        [JsonProperty("_b")]
        public int emp_id { get; set; }

        [JsonProperty("_c")]
        public string obs_abreviatura { get; set; }

        [JsonProperty("_d")]
        public string obs_descripcion { get; set; }

        [JsonProperty("_e")]
        public int gde_id { get; set; }

        [JsonProperty("_f")]
        public string gde_descripcion { get; set; }

        [JsonProperty("_g")]
        public int obs_estado { get; set; }

        [JsonProperty("_h")]
        public string obs_pideFoto { get; set; }

        [JsonProperty("_i")]
        public string obs_noPideFoto { get; set; }
    }

    #region Request & Response

    /// <summary>
    /// Autor: jlucero
    /// Fecha: 2015-06-14
    /// </summary>
    public class Request_CRUD_Observacion
    {
        [JsonProperty("_a")]
        public int obs_id { get; set; }

        [JsonProperty("_b")]
        public int emp_id { get; set; }

        [JsonProperty("_c")]
        public string obs_abreviatura { get; set; }

        [JsonProperty("_d")]
        public string obs_descripcion { get; set; }

        [JsonProperty("_e")]
        public int gde_id { get; set; }

        [JsonProperty("_f")]
        public int obs_estado { get; set; }

        [JsonProperty("_g")]
        public int obs_usuario { get; set; }

        [JsonProperty("_h")]
        public string obs_pideFoto { get; set; }

        [JsonProperty("_i")]
        public string obs_noPideFoto { get; set; }
    }

    #endregion
}
