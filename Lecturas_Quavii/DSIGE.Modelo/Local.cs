using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using Newtonsoft;
using Newtonsoft.Json;

namespace DSIGE.Modelo
{
    public class Local : Auditoria
    {
        [JsonProperty("_a")]
        public int loc_id { get; set; }

        [JsonProperty("_b")]
        public int emp_id { get; set; }

        [JsonProperty("_c")]
        public string loc_nombre { get; set; }

        [JsonProperty("_d")]
        public string loc_direccion { get; set; }

        [JsonProperty("_e")]
        public int loc_orden { get; set; }

        [JsonProperty("_f")]
        public string loc_latitud { get; set; }

        [JsonProperty("_g")]
        public string loc_longitud { get; set; }

        [JsonProperty("_h")]
        public int loc_estado { get; set; }
    }

    #region Request & Response

    /// <summary>
    /// Autor: jlucero
    /// Fecha: 2015-06-14
    /// </summary>
    public class Request_Local_Select
    {
        [JsonProperty("_a")]
        public int emp_id { get; set; }
    }


    /// <summary>
    /// Autor: rcontreras
    /// Fecha: 26-09-2015
    /// </summary>
    public class Request_CRUD_Local
    {
        [JsonProperty("_a")]
        public int loc_id { get; set; }

        [JsonProperty("_b")]
        public int emp_id { get; set; }

        [JsonProperty("_c")]
        public string loc_nombre { get; set; }

        [JsonProperty("_d")]
        public string loc_direccion { get; set; }

        [JsonProperty("_e")]
        public string loc_latitud { get; set; }

        [JsonProperty("_f")]
        public string loc_longitud { get; set; }

        [JsonProperty("_g")]
        public int loc_estado { get; set; }

        [JsonProperty("_h")]
        public int loc_usuario { get; set; }
    }

    #endregion
}
