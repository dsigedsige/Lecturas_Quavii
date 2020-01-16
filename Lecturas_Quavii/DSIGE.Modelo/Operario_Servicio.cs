using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using Newtonsoft;
using Newtonsoft.Json;

namespace DSIGE.Modelo
{
    public class Operario_Servicio : Auditoria
    {
        [JsonProperty("_a")]
        public int sop_id { get; set; }

        [JsonProperty("_b")]
        public int ser_id { get; set; }

        [JsonProperty("_c")]
        public int ope_id { get; set; }

        [JsonProperty("_d")]
        public int sop_estado { get; set; }
    }

    #region Request & Response

    /// <summary>
    /// Autor: jlucero
    /// Fecha: 2015-06-21
    /// </summary>
    public class Request_Operario_Servicio_Asignacion
    {
        [JsonProperty("_a")]
        public string ope_id { get; set; }

        [JsonProperty("_b")]
        public int ser_id { get; set; }

        [JsonProperty("_c")]
        public int sop_estado { get; set; }

        [JsonProperty("_d")]
        public int usu_id { get; set; }

        [JsonProperty("_e")]
        public int emp_id { get; set; }

        [JsonProperty("_f")]
        public int loc_id { get; set; }
    }

    #endregion
}
