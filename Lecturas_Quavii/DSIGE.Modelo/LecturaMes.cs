using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DSIGE.Modelo
{
  public  class LecturaMes
    {

        [JsonProperty("medidor_lectura", NullValueHandling = NullValueHandling.Ignore)]
        public String medidor_lectura { get; set; }

        [JsonProperty("mesUno_lectura", NullValueHandling = NullValueHandling.Ignore)]
        public string mesUno_lectura { get; set; }

        [JsonProperty("mesDos_lectura", NullValueHandling = NullValueHandling.Ignore)]
        public string mesDos_lectura { get; set; }

        [JsonProperty("mesTres_lectura", NullValueHandling = NullValueHandling.Ignore)]
        public string mesTres_lectura { get; set; }

        [JsonProperty("mesCuatro_lectura", NullValueHandling = NullValueHandling.Ignore)]
        public string mesCuatro_lectura { get; set; }

        [JsonProperty("mesCinco_lectura", NullValueHandling = NullValueHandling.Ignore)]
        public string mesCinco_lectura { get; set; }

        [JsonProperty("consumo1", NullValueHandling = NullValueHandling.Ignore)]
        public string consumo1 { get; set; }

        [JsonProperty("consumo2", NullValueHandling = NullValueHandling.Ignore)]
        public string consumo2 { get; set; }

        [JsonProperty("consumo3", NullValueHandling = NullValueHandling.Ignore)]
        public string consumo3 { get; set; }

        [JsonProperty("consumo4", NullValueHandling = NullValueHandling.Ignore)]
        public string consumo4 { get; set; }

        [JsonProperty("promedioConsumo", NullValueHandling = NullValueHandling.Ignore)]
        public decimal? promedioConsumo { get; set; }

        [JsonProperty("parametroMaximo", NullValueHandling = NullValueHandling.Ignore)]
        public Decimal? parametroMaximo { get; set; }

        [JsonProperty("parametroMinimo", NullValueHandling = NullValueHandling.Ignore)]
        public Decimal? parametroMinimo { get; set; }

        [JsonProperty("lecturaMaximo", NullValueHandling = NullValueHandling.Ignore)]
        public Decimal? lecturaMaximo { get; set; }

        [JsonProperty("lecturaMinimo", NullValueHandling = NullValueHandling.Ignore)]
        public Decimal? LecturaMinimo { get; set; }


    }
}
