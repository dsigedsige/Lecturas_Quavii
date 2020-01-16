using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using Newtonsoft;
using Newtonsoft.Json;

namespace DSIGE.Modelo
{
    public class Corte : Auditoria
    {
        [JsonProperty("_a")]
        public int emp_id { get; set; }

        [JsonProperty("_b")]
        public int ser_id { get; set; }

        [JsonProperty("_c")]
        public DateTime cor_importacion_fecha { get; set; }

        [JsonProperty("_d")]
        public string cor_importacion_archivo { get; set; }

        [JsonProperty("_e")]
        public int loc_id { get; set; }

        [JsonProperty("_f")]
        public string cor_aviso_clase { get; set; }

        [JsonProperty("_g")]
        public string cor_aviso { get; set; }

        [JsonProperty("_h")]
        public string cor_aviso_fecha { get; set; }

        [JsonProperty("_i")]
        public string cor_documento_bloqueo { get; set; }

        [JsonProperty("_j")]
        public string cor_suministro { get; set; }

        [JsonProperty("_k")]
        public string cor_interlocutor { get; set; }

        [JsonProperty("_l")]
        public string cor_clase_cuenta { get; set; }

        [JsonProperty("_m")]
        public string cor_deuda { get; set; }

        [JsonProperty("_n")]
        public string cor_cantidad { get; set; }

        [JsonProperty("_o")]
        public string cor_instalacion_numero { get; set; }

        [JsonProperty("_p")]
        public string cor_medidor { get; set; }

        [JsonProperty("_q")]
        public string cor_direccion { get; set; }

        [JsonProperty("_r")]
        public string cor_distrito { get; set; }

        [JsonProperty("_s")]
        public string cor_seccion { get; set; }

        [JsonProperty("_t")]
        public int ope_id { get; set; }

        [JsonProperty("_u")]
        public int cor_orden { get; set; }

        [JsonProperty("_v")]
        public int cor_estado { get; set; }

        [JsonProperty("_w")]
        public int usu_id { get; set; }

        [JsonProperty("_x")]
        public DateTime cor_asignacion_fecha { get; set; }
    }


    public class CorteTemporalCorte
    {
  
        public int emp_id { get; set; }

       
        public int ser_id { get; set; }

        public int id_tiposervicio { get; set; }
 
        public DateTime cor_importacion_fecha { get; set; }

   
        public string cor_importacion_archivo { get; set; }

    
        public int loc_id { get; set; }

  
        public string cor_aviso_clase { get; set; }

      
        public string cor_aviso { get; set; }

     
        public string cor_aviso_fecha { get; set; }

        public string cor_documento_bloqueo { get; set; }

  
        public string cor_suministro { get; set; }

        public string cor_interlocutor { get; set; }

    
        public string cor_clase_cuenta { get; set; }

       
        public string cor_deuda { get; set; }

       
        public string cor_cantidad { get; set; }

      
        public string cor_instalacion_numero { get; set; }

        public string cor_medidor { get; set; }

      
        public string cor_direccion { get; set; }

        public string unidad_corte { get; set; }  
        public string cor_distrito { get; set; }

       
        public string cor_seccion { get; set; }

      
        public int tecnico { get; set; }
        public string nombre_ope { get; set; }
      
        public int cor_orden { get; set; }

       
        public int cor_estado { get; set; }

    
        public int usu_id { get; set; }
        public int contador { get; set; }

        public int cant_correcto { get; set; }
        public int cant_erroneo { get; set; }

        
        public DateTime cor_asignacion_fecha { get; set; }
    }


    public class corteSumnistro
    {

        public string cor_suministro { get; set; }

        public string cor_medidor { get; set; }

        public string crea_nombre { get; set; }

        public string cor_direccion { get; set; }

    }

    public class Reparto
    {
        public int total { get; set; }
        public string unidad_lectura { get; set; }
        public string nombre_UnidadLectura { get; set; }
        public int id_operario { get; set; }
        public string id_operario_cambiar { get; set; }
        
        public string id_operario_aux { get; set; }        
        public string NRORECIBO_REPARTO { get; set; }
        public string EMPRESA { get; set; }
        public string CLIENTE { get; set; }
        public string ORDEN { get; set; }
        public string CTA { get; set; }
        public string FECHA_MAX { get; set; }
        public string DIRECCION { get; set; }
        public string DISTRITO { get; set; }
        public string TIPO_REPARTO { get; set; }
        public string CICLO { get; set; }
        public string UL { get; set; }
        public string flag_compartido { get; set; }
         

    }
    #region Request & Response

    /// <summary>
    /// Autor: jlucero
    /// Fecha: 2015-07-17
    /// </summary>
    public class Request_Corte_Validacion
    {
        [JsonProperty("_a")]
        public string archivo { get; set; }
    }

    /// <summary>
    /// Autor: jlucero
    /// Fecha: 2015-07-17
    /// </summary>
    public class Request_Corte_Inserta_Masivo
    {
        [JsonProperty("_a")]
        public string cor_archivo { get; set; }

        [JsonProperty("_b")]
        public string cor_asignacion_fecha { get; set; }

        [JsonProperty("_c")]
        public string cor_id { get; set; }

        [JsonProperty("_d")]
        public int usu_id { get; set; }
    }

    #endregion
}