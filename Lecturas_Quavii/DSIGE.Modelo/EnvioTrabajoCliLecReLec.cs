using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Newtonsoft;
using Newtonsoft.Json;

namespace DSIGE.Modelo
{
    public class EnvioTrabajoCliLecReLec
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

        [JsonProperty("_r")]
        public string foto { get; set; }

        [JsonProperty("_s")]
        public string lecMovil { get; set; }

        [JsonProperty("_t")]
        public string consuActu { get; set; }

        [JsonProperty("_u")]
        public string lecProm { get; set; }

        [JsonProperty("_vv")]
        public string lecPorcen { get; set; }

        [JsonProperty("_w")]
        public string lecVali { get; set; }

        [JsonProperty("_x")]
        public string lecManual { get; set; }

        [JsonProperty("_y")]
        public string idlecObs { get; set; }

        [JsonProperty("_z")]
        public string lecSucursal { get; set; }

        [JsonProperty("_aa")]
        public string lecSeccionLectura { get; set; }

        [JsonProperty("_ab")]
        public string lecCorrelativo { get; set; }

        [JsonProperty("_ac")]
        public string lecMarcaMedidor { get; set; }

        [JsonProperty("_ad")]
        public string confirLectura { get; set; }

        [JsonProperty("_ae")]
        public string blockLectura { get; set; }
    }

    public class EnvioTrabajoCliLecReLecLecturaPendiente
    {
        [JsonProperty("_a")]
        public int idOpeLectura { get; set; }

        [JsonProperty("_b")]
        public string nomOpe { get; set; }

        [JsonProperty("_c")]
        public string nomEstado { get; set; }

        [JsonProperty("_d")]
        public int cantidad { get; set; }
    }

    public class Edicion_Envio_Trabajo_Cliente_Lectura
    {

        public int id_Lectura { get; set; }
        public string suministro_lectura { get; set; }
        public string medidor_lectura { get; set; }
        public string marcaMedidor_lectura { get; set; }
        public string direccion_lectura { get; set; }
        public string id_Operario_Lectura { get; set; }
        public string nombre_operario { get; set; }

        public string Lectura_Manual_Lectura { get; set; }
        public string fechaLecturaMovil_lectura { get; set; }
        public string id_Observacion_Lectura { get; set; }
        public string descripcion_observacion { get; set; }

        public string Seccion_Lectura { get; set; }
        public string Zona_lectura { get; set; }

        public string abreviatura_observacion { get; set; }

        public string Block_lectura { get; set; }
        public string fechaAsignacion_Lectura { get; set; }
        public string confirmacion_Lectura { get; set; }
        public string Observacion_lectura { get; set; }

        public string lecturaMovil_lectura { get; set; }

    }

    public class Fotoselfie_reparto
    {
        public string RutaFoto { get; set; }
    }

}
