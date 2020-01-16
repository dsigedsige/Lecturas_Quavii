using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using Newtonsoft;
using Newtonsoft.Json;

namespace DSIGE.Modelo
{
    public class Operario : Auditoria
    {
        [JsonProperty("_a")]
        public int ope_id { get; set; }

        [JsonProperty("_b")]
        public int emp_id { get; set; }

        [JsonProperty("_c")]
        public int loc_id { get; set; }

        [JsonProperty("_d")]
        public string ope_documento { get; set; }

        [JsonProperty("_e")]
        public string ope_documento_tipo { get; set; }

        [JsonProperty("_f")]
        public string ope_apellido { get; set; }

        [JsonProperty("_g")]
        public string ope_nombre { get; set; }

        [JsonProperty("_h")]
        public string ope_foto { get; set; }

        [JsonProperty("_i")]
        public string ope_celular { get; set; }

        [JsonProperty("_j")]
        public string ope_usuario { get; set; }

        [JsonProperty("_k")]
        public string ope_contrasenia { get; set; }

        [JsonProperty("_l")]
        public string ope_online { get; set; }

        [JsonProperty("_m")]
        public int ope_estado { get; set; }

        [JsonProperty("_n")]
        public string ope_tipo_usuario { get; set; }

        [JsonProperty("_mail")]
        public string ope_email { get; set; }

        [JsonProperty("_x")]
        public string servicio { get; set; }

    }

    #region Request & Response

    /// <summary>
    /// Autor: jlucero
    /// Fecha: 2015-06-14
    /// </summary>
    public class Request_CRUD_Operario
    {
        [JsonProperty("_a")]
        public int ope_id { get; set; }

        [JsonProperty("_b")]
        public int emp_id { get; set; }

        [JsonProperty("_c")]
        public int loc_id { get; set; }

        [JsonProperty("_d")]
        public string ope_documento { get; set; }

        [JsonProperty("_e")]
        public string ope_documento_tipo { get; set; }

        [JsonProperty("_f")]
        public string ope_apellido { get; set; }

        [JsonProperty("_g")]
        public string ope_nombre { get; set; }

        [JsonProperty("_cargo")]
        public string ope_cargo { get; set; }


        [JsonProperty("_h")]
        public string ope_foto { get; set; }

        [JsonProperty("_i")]
        public string ope_celular { get; set; }

        [JsonProperty("_j")]
        public string ope_usuario { get; set; }

        [JsonProperty("_k")]
        public string ope_contrasenia { get; set; }

        [JsonProperty("_l")]
        public string ope_online { get; set; }

        [JsonProperty("_m")]
        public int ope_estado { get; set; }

        [JsonProperty("_n")]
        public int ope_crea { get; set; }

        [JsonProperty("_o")]
        public string ope_tipo_usuario { get; set; }

        [JsonProperty("_mail")]
        public string ope_email { get; set; }

    }

    /// <summary>
    /// Autor: jlucero
    /// Fecha: 2015-06-19
    /// </summary>
    public class Request_Operario_Empresa_Local_Opcion
    {
        [JsonProperty("_a")]
        public int emp_id { get; set; }

        [JsonProperty("_b")]
        public int loc_id { get; set; }

        [JsonProperty("_c")]
        public int opcion { get; set; }

        [JsonProperty("_d")]
        public int tip_serv { get; set; }
    }

    /// <summary>
    /// Autor: rcontreras
    /// Fecha: 23-09-2015
    /// </summary>
    public class Request_Local_Operario
    {
        [JsonProperty("_a")]
        public string ope_id { get; set; }
    }

    #endregion
}
