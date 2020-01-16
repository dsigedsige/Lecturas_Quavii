using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using Newtonsoft;
using Newtonsoft.Json;

namespace DSIGE.Modelo
{
    public class Sesion
    {
        public Sesion_Empresa empresa { get; set; }

        public Sesion_Usuario usuario { get; set; }

        public List<Sesion_Modulo> modulo { get; set; }
    }

    public class Sesion_Empresa
    {
        [JsonProperty("_a")]
        public int emp_id { get; set; }

        [JsonProperty("_b")]
        public string emp_descripcion { get; set; }

        [JsonProperty("_c")]
        public string emp_razon_social { get; set; }

        [JsonProperty("_d")]
        public string emp_imagen { get; set; }
    }

    public class Sesion_Usuario
    {
        [JsonProperty("_a")]
        public string aop_descripcion { get; set; }

        [JsonProperty("_b")]
        public int usu_id { get; set; }

        [JsonProperty("_c")]
        public string usu_nombre { get; set; }

        [JsonProperty("_d")]
        public int idPerfil { get; set; }
    }

    public class Sesion_Modulo
    {
        [JsonProperty("_a")]
        public int mod_id { get; set; }

        [JsonProperty("_b")]
        public string mod_descripcion { get; set; }

        [JsonProperty("_c")]
        public List<Sesion_Item> item { get; set; }

        [JsonProperty("_d")]
        public string nombre_usuario { get; set; }

        [JsonProperty("_e")]
        public string urlImagen { get; set; }

        [JsonProperty("_f")]
        public string dop_url { get; set; }
    }

    public class Sesion_Item {
        [JsonProperty("_a")]
        public int dop_id { get; set; }

        [JsonProperty("_b")]
        public string dop_descripcion { get; set; }

        [JsonProperty("_c")]
        public string dop_url { get; set; }

        [JsonProperty("_d")]
        public string urlImagen { get; set; }

        [JsonProperty("_e")]
        public int idPerfil { get; set; }
    }

    #region Request & Response

    /// <summary>
    /// Fecha: 2015-06-12
    /// </summary>
    public class Request_Sesion {
        [JsonProperty("_a")]
        public string usuario { get; set; }

        [JsonProperty("_b")]
        public string contrasenia { get; set; }
    }

    #endregion
}
