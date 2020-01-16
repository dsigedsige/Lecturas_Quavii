using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Newtonsoft;
using Newtonsoft.Json;

namespace DSIGE.Modelo
{
    public class AsignaLecturaReLectura
    {
        [JsonProperty("_a")]
        public int corre { get; set; }

        [JsonProperty("_b")]
        public int idLectura { get; set; }

        [JsonProperty("_c")]
        public string suministroLectura { get; set; }

        [JsonProperty("_d")]
        public string medidorLectura { get; set; }

        [JsonProperty("_e")]
        public string dirLectura { get; set; }

        [JsonProperty("_f")]
        public string lecAnte { get; set; }

        [JsonProperty("_g")]
        public string lecMin { get; set; }

        [JsonProperty("_h")]
        public string lecMax { get; set; }

        [JsonProperty("_i")]
        public string ope_nombre { get; set; }

        [JsonProperty("_j")]
        public string zonaLectura { get; set; }

        [JsonProperty("_k")]
        public string fecAsignaLectura { get; set; }

        [JsonProperty("_l")]
        public string clienLectura { get; set; }

        [JsonProperty("_m")]
        public string Estado { get; set; }

        [JsonProperty("_n")]
        public string lectura { get; set; }

        [JsonProperty("_o")]
        public string fecLectura { get; set; }

        [JsonProperty("_p")]
        public string obsLectura { get; set; }

        [JsonProperty("_q")]
        public string nroinstalacion { get; set; }
    }

    #region Request & Response
    /// <summary>
    /// Autor: rcontreras
    /// Fecha: 17-09-2015
    /// </summary>
    public class Request_Asigna_Lectura_ReLectura
    {
        [JsonProperty("_a")]
        public int emp_id { get; set; }

        [JsonProperty("_b")]
        public int tip_ser_id { get; set; }

        [JsonProperty("_c")]
        public int id_local { get; set; }

        [JsonProperty("_d")]
        public string suministro { get; set; }

        [JsonProperty("_e")]
        public string medidor { get; set; }

        [JsonProperty("_f")]
        public int tecnico_asig_id { get; set; }

        [JsonProperty("_g")]
        public int estado_asig { get; set; }

        [JsonProperty("_h")]
        public DateTime fecha_asig { get; set; }
    }

    /// <summary>
    /// Autor: rcontreras
    /// Fecha: 18-09-2015
    /// </summary>
    public class Request_actualiza_app_movil_lec_relec
    {
        [JsonProperty("_a")]
        public string idLectura { get; set; }

        [JsonProperty("_b")]
        public DateTime fechActu { get; set; }

        [JsonProperty("_c")]
        public int usuEdi { get; set; }

        [JsonProperty("_d")]
        public int idEmp { get; set; }

        [JsonProperty("_e")]
        public int idOperador { get; set; }

        [JsonProperty("_f")]
        public int opcion { get; set; }
    }
    #endregion

}
