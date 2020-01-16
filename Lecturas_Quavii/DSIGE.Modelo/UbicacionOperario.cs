using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Newtonsoft;
using Newtonsoft.Json;

namespace DSIGE.Modelo
{

    public class UbicacionOperario_GPS
    {

       //''   id_Operario, ope_nombre, latitud_lectura, longitud_lectura, FechaGPS, tiempo_Espera

        [JsonProperty("_a")]
        public int id_Operario { get; set; }

        [JsonProperty("_b")]
        public string  ope_nombre { get; set; }

        [JsonProperty("_c")]
        public string  latitud_lectura { get; set; }

        [JsonProperty("_d")]
        public string  longitud_lectura { get; set; }

        [JsonProperty("_e")]
        public string  FechaGPS { get; set; }

        [JsonProperty("_f")]
        public int tiempo_Espera { get; set; }

        [JsonProperty("_g")]
        public string Fecha_Inicio_parada { get; set; }

        [JsonProperty("_h")]
        public string fecha_fin_parada { get; set; }

        [JsonProperty("_i")]
        public string Minutos_Espera_Conv { get; set; }

            }


    public class UbicacionOperario
    {
        [JsonProperty("_a")]
        public int id_ope { get; set; }

        [JsonProperty("_b")]
        public string latitud { get; set; }

        [JsonProperty("_c")]
        public string longitud { get; set; }

        [JsonProperty("_d")]
        public int id_gps { get; set; }

        [JsonProperty("_f")]
        public string ope_nombre { get; set; }

        [JsonProperty("_g")]
        public string ope_tipo { get; set; }

        [JsonProperty("_h")]
        public int totAsig { get; set; }

        [JsonProperty("_i")]
        public int totReali { get; set; }

        [JsonProperty("_j")]
        public int totPend { get; set; }

        [JsonProperty("_k")]
        public int pocenAvance { get; set; }


        [JsonProperty("_l")]
        public int idLectura { get; set; }

        [JsonProperty("_m")]
        public string suministro { get; set; }

        [JsonProperty("_n")]
        public string medidor { get; set; }

        [JsonProperty("_o")]
        public string direcLectura { get; set; }

        [JsonProperty("_p")]
        public string nom_cliente { get; set; }

        [JsonProperty("_q")]
        public string lectura { get; set; }

        [JsonProperty("_r")]
        public string foto_lectura { get; set; }

        [JsonProperty("_s")]
        public int lec_estado { get; set; }

        [JsonProperty("_t")]
        public string fasig { get; set; }

        [JsonProperty("_u")]
        public string fmovil { get; set; }

        [JsonProperty("_v")]
        public DateTime fechaRegistro { get; set; }

        [JsonProperty("_w")]
        public string isOnline { get; set; }


        [JsonProperty("_x")]
        public string gpsActivo { get; set; }

        [JsonProperty("_y")]
        public string estaBateria { get; set; }

        [JsonProperty("_z")]
        public string fechaAndroid { get; set; }

        [JsonProperty("_aa")]
        public string paro { get; set; }

        [JsonProperty("_bb")]
        public string PlanDatos { get; set; }


        [JsonProperty("_cc")]
        public string ModoAvion { get; set; }

        [JsonProperty("_dd")]
        public string HoraAndroid { get; set; }

        


        


    }


    public class Ubicacion_Lectura
    {

        public int id_TipoServicio { get; set; }
        public string nombre_tiposervicio { get; set; }
        public int estado { get; set; }

        public int id_Operario { get; set; }
        public string desc_operario { get; set; }


        public int id_ope { get; set; }
        public string latitud { get; set; }
        public string longitud { get; set; }
        public int id_gps { get; set; }
        public string ope_nombre { get; set; }
        public string ope_tipo { get; set; }
        public int totAsig { get; set; }
        public int totReali { get; set; }
        public int totPend { get; set; }
        public int pocenAvance { get; set; }
        public int idLectura { get; set; }
        public string suministro { get; set; }
        public string medidor { get; set; }
        public string direcLectura { get; set; }
        public string nom_cliente { get; set; }
        public string lectura { get; set; }
        public string foto_lectura { get; set; }
        public int lec_estado { get; set; }
        public string fasig { get; set; }
        public string fmovil { get; set; }
        public DateTime fechaRegistro { get; set; }
        public string isOnline { get; set; }
        public string gpsActivo { get; set; }
        public string estaBateria { get; set; }
        public string fechaAndroid { get; set; }
        public string paro { get; set; }
        public string PlanDatos { get; set; }
        public string ModoAvion { get; set; }
        public string HoraAndroid { get; set; }

    }

}
