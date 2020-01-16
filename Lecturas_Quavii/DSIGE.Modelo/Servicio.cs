using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using Newtonsoft;
using Newtonsoft.Json;

namespace DSIGE.Modelo
{
    public class Servicio : Auditoria
    {
        [JsonProperty("_a")]
        public int ser_id { get; set; }

        [JsonProperty("_b")]
        public string ser_descripcion { get; set; }

        [JsonProperty("_c")]
        public int ser_estado { get; set; }
    }

    #region Request & Response

    /// <summary>
    /// Autor: jlucero
    /// Fecha: 2015-06-14
    /// </summary>
    public class Request_CRUD_Servicio {
        [JsonProperty("_a")]
        public int id { get; set; }

        [JsonProperty("_b")]
        public string descripcion { get; set; }

        [JsonProperty("_c")]
        public int estado { get; set; }

        [JsonProperty("_d")]
        public int usuario { get; set; }
    }

    /// <summary>
    /// Autor: jlucero
    /// Fecha: 2015-06-19
    /// </summary>
    public class Request_Servicio_Operario
    {
        [JsonProperty("_a")]
        public string ope_id { get; set; }
    }

    #endregion
}