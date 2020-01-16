using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using Newtonsoft;
using Newtonsoft.Json;

namespace DSIGE.Modelo
{
    public class Lectura : Auditoria
    {
        [JsonProperty("_a")]
        public int emp_id { get; set; }

        [JsonProperty("_b")]
        public int ser_id { get; set; }

        [JsonProperty("_c")]
        public DateTime lec_importacion_fecha { get; set; }

        [JsonProperty("_d")]
        public string lec_importacion_archivo { get; set; }

        [JsonProperty("_e")]
        public string lec_suministro { get; set; }

        [JsonProperty("_f")]
        public string lec_medidor { get; set; }

        [JsonProperty("_g")]
        public string lec_medidor_marca { get; set; }

        [JsonProperty("_h")]
        public string lec_numero_instalacion { get; set; }

        [JsonProperty("_i")]
        public string lec_direccion { get; set; }

        [JsonProperty("_j")]
        public int ubi_id { get; set; }

        [JsonProperty("_k")]
        public string lec_distrito { get; set; }

        [JsonProperty("_l")]
        public string lec_cliente_nombre { get; set; }

        [JsonProperty("_m")]
        public string lec_cliente_tipo { get; set; }

        [JsonProperty("_n")]
        public string lec_interlocutor { get; set; }

        [JsonProperty("_o")]
        public string lec_categoria { get; set; }

        [JsonProperty("_p")]
        public string lec_anterior { get; set; }

        [JsonProperty("_q")]
        public string lec_zona { get; set; }

        [JsonProperty("_r")]
        public DateTime lec_asignacion_fecha { get; set; }

        [JsonProperty("_s")]
        public int ope_id { get; set; }

        [JsonProperty("_t")]
        public int lec_orden { get; set; }

        [JsonProperty("_u")]
        public int lec_reordenamiento { get; set; }

        [JsonProperty("_v")]
        public int lec_estado { get; set; }

        [JsonProperty("_w")]
        public int usu_id { get; set; }

        [JsonProperty("_x")]
        public int loc_id { get; set; }

        [JsonProperty("_y")]
        public int lec_seccion { get; set; }

        [JsonProperty("_z")]
        public int lec_correlativo { get; set; }



        [JsonProperty("_ab")]
        public string dniOpe { get; set; }

        [JsonProperty("_ac")]
        public string bloque { get; set; }

        [JsonProperty("_ad")]
        public int orden { get; set; }


        [JsonProperty("MesUno")]
        public string MesUno { get; set; }

        [JsonProperty("MesDos")]
        public string MesDos { get; set; }

        [JsonProperty("MesTres")]
        public string MesTres { get; set; }

        [JsonProperty("MesCuatro")]
        public string MesCuatro { get; set; }

        [JsonProperty("MesCinco")]
        public string MesCinco { get; set; }

        [JsonProperty("Promedio")]
        public int Promedio { get; set; }

        [JsonProperty("Minimo")]
        public int Minimo { get; set; }

        [JsonProperty("Maximo")]
        public int Maximo { get; set; }

        [JsonProperty("Lect_Minima")]
        public int Lect_Minima { get; set; }

        [JsonProperty("Lect_Maxima")]
        public int Lect_Maxima { get; set; }

        [JsonProperty("Ultimas_Lect")]
        public int Ultimas_Lect { get; set; }

        [JsonProperty("id_UsuarioExpor")]
        public int id_UsuarioExpor { get; set; }

        [JsonProperty("nombre_ope")]
        public string nombre_ope { get; set; }

    }

    public class Cls_Entidad_ReporteDiarioNew
    {
        public int id_Operario { get; set; }
        public string apellidos_operario { get; set; }
        public string Efectivo { get; set; }
        public string totalHora { get; set; }
        public string Fecha { get; set; }

    }

    public class LecturaHistorico
    {
        [JsonProperty("_a")]
        public int id_ope { get; set; }

        [JsonProperty("_b")]
        public string des_ope { get; set; }

        [JsonProperty("_c")]
        public double total { get; set; }

        [JsonProperty("_d")]
        public double realizado { get; set; }

        [JsonProperty("_e")]
        public double conFoto { get; set; }

        [JsonProperty("_f")]
        public double pendiente { get; set; }

        [JsonProperty("_g")]
        public double avance { get; set; }

        [JsonProperty("_h")]
        public string f_ini { get; set; }

        [JsonProperty("_i")]
        public string f_fin { get; set; }

        [JsonProperty("_j")]
        public string horas { get; set; }

        [JsonProperty("_k")]
        public string obs { get; set; }

        [JsonProperty("_l")]
        public int sinFoto { get; set; }

    }

    public class urlLectura
    {
        [JsonProperty("_a")]
        public string url_lectura { get; set; }
    }



    public class LecturaTomadas
    {

        //[JsonProperty("_a")]
        //public int id_lectura { get; set; }

        //[JsonProperty("_b")]
        //public int orden_lectura { get; set; }

        //[JsonProperty("_c")]
        //public string suministro_lectura { get; set; }

        //[JsonProperty("_d")]
        //public string medidor_lectura { get; set; }

        //[JsonProperty("_e")]
        //public string marcaMedidor_lectura { get; set; }

        //[JsonProperty("_f")]
        //public string direccion_lectura { get; set; }


        //[JsonProperty("_g")]
        //public string lectura_7 { get; set; }

        //[JsonProperty("_h")]
        //public string lectura_8 { get; set; }

        //[JsonProperty("_i")]
        //public string lectura_9 { get; set; }

        //[JsonProperty("_j")]
        //public string lectura_10 { get; set; }

        //[JsonProperty("_k")]
        //public string lectura_movil { get; set; }

        //[JsonProperty("_l")]
        //public string operario { get; set; }

        //[JsonProperty("_m")]
        //public string fechaLectura { get; set; }

        //[JsonProperty("_n")]
        //public string obs { get; set; }

        //[JsonProperty("_o")]
        //public string ubicacion { get; set; }

        //[JsonProperty("_p")]
        //public string notas { get; set; }

        //[JsonProperty("_q")]
        //public string tieneFoto_lectura { get; set; }

        //[JsonProperty("_r")]
        //public string nombre_foto { get; set; }

        //[JsonProperty("_s")]
        //public string latitud_lectura { get; set; }

        //[JsonProperty("_t")]
        //public string longitud_lectura { get; set; }

        //[JsonProperty("_fe")]
        //public string fechaData { get; set; }

        //[JsonProperty("_Z")]
        //public int id_TipoServicio { get; set; }


        [JsonProperty("_a")]
        public int id_lectura { get; set; }

        [JsonProperty("_b")]
        public int orden_lectura { get; set; }

        [JsonProperty("_c")]
        public string suministro_lectura { get; set; }

        [JsonProperty("_d")]
        public string medidor_lectura { get; set; }

        [JsonProperty("_e")]
        public string marcaMedidor_lectura { get; set; }

        [JsonProperty("_f")]
        public string direccion_lectura { get; set; }


        [JsonProperty("_g")]
        public string lectura_7 { get; set; }

        [JsonProperty("_h")]
        public string lectura_8 { get; set; }

        [JsonProperty("_i")]
        public string lectura_9 { get; set; }

        [JsonProperty("_j")]
        public string lectura_10 { get; set; }

        [JsonProperty("_k")]
        public string lectura_movil { get; set; }

        [JsonProperty("_l")]
        public string operario { get; set; }

        [JsonProperty("_m")]
        public string fechaLectura { get; set; }

        [JsonProperty("_xx")]
        public string fechaAsignacion_Lectura { get; set; }

        [JsonProperty("_n")]
        public string obs { get; set; }

        [JsonProperty("_o")]
        public string ubicacion { get; set; }

        [JsonProperty("_p")]
        public string notas { get; set; }

        [JsonProperty("_q")]
        public string tieneFoto_lectura { get; set; }

        [JsonProperty("_r")]
        public string nombre_foto { get; set; }

        [JsonProperty("_s")]
        public string latitud_lectura { get; set; }

        [JsonProperty("_t")]
        public string longitud_lectura { get; set; }

        [JsonProperty("_fe")]
        public string fechaData { get; set; }

        [JsonProperty("_Z")]
        public int id_TipoServicio { get; set; }

        [JsonProperty("_kb")]
        public string size_foto { get; set; }


        [JsonProperty("_fa")]
        public string fotoActa { get; set; }


    }


    public class LecturaCortesTomadas
    {
        
 
        public int id_lectura { get; set; } 
        public int orden_lectura { get; set; } 
        public string suministro_lectura { get; set; } 
        public string medidor_lectura { get; set; } 
        public string marcaMedidor_lectura { get; set; } 
        public string direccion_lectura { get; set; } 
        public string lectura_7 { get; set; } 
        public string lectura_8 { get; set; } 
        public string lectura_9 { get; set; }         
        public string lectura_10 { get; set; } 
        public string lectura_movil { get; set; } 
        public string operario { get; set; } 
        public string fechaLectura { get; set; } 
        public string obs { get; set; } 
        public string ubicacion { get; set; } 
        public string notas { get; set; } 
        public string tieneFoto_lectura { get; set; } 
        public string nombre_foto { get; set; } 
        public string latitud_lectura { get; set; } 
        public string longitud_lectura { get; set; } 
        public string fechaData { get; set; }
    }


    public class ResumenLecturasCab
    {


        public int id_servicio { get; set; }
        public int id_ope { get; set; }
        public string des_ope { get; set; }
        public double total { get; set; }
        public double realizado { get; set; }
        public double conFoto { get; set; }
        public double pendiente { get; set; }
        public double avance { get; set; }
        public string f_ini { get; set; }
        public string f_fin { get; set; }
        public string horas { get; set; }
        public string obs { get; set; }
        public int sinFoto { get; set; }

    }

    public class ResumenLecturas
    {

        public int id_Lectura { get; set; }
        public int id_servicio { get; set; }
        public string suministro_lectura { get; set; }
        public string medidor_lectura { get; set; }
        public string direccion_lectura { get; set; }
        public string nombreCliente_lectura { get; set; }
        public string id_Operario_Lectura { get; set; }
        public string operario { get; set; }
        public string Sector_lectura { get; set; }
        public string Zona_lectura { get; set; }
        public string DNS_lectura { get; set; }
        public string LecturaMovil_Lectura { get; set; }
        public string confirmacion_Lectura { get; set; }        
        public string abreviatura_observacion { get; set; }
        public string descripcion_estado { get; set; }
        public string foto { get; set; }
        public string latitud_lectura { get; set; }
        public string longitud_lectura { get; set; }
        public string mapa { get; set; }
               public string url { get; set; }
        public string latitudMovil_lectura { get; set; }
        public string longitudMovil_lectura { get; set; }

        public int estado { get; set; }

    }


    public class Fotos
    {
       public int id_Lectura { get; set; }
        public string foto { get; set; }
    }


    public class Lecturas_Tomadas
    {

        public int id_Lectura { get; set; }
        public string Suministro_lectura { get; set; }
        public string Medidor_lectura { get; set; }

        public string Confirmacion_lectura { get; set; }
        public string lectura_Anterior { get; set; }
        public decimal promedioConsumo_Lectura { get; set; }
        public string Operario { get; set; }

        public string observacion_lectura { get; set; }
        public string descripcion_observacion { get; set; }
        public string fechaLecturaMovil_Lectura { get; set; }
        public string tieneFoto_lectura { get; set; }

        public string MEDIDOR { get; set; }
        public string CTA_CTO { get; set; }
        public string FECHA_PLAN_LECTURA { get; set; }
        public string MES { get; set; }
        public string LECTURA { get; set; }
    }



    #region Request & Response

    /// <summary>
    /// Autor: jlucero
    /// Fecha: 2015-07-07
    /// </summary>
    public class Request_Lectura_Validacion
    {
        [JsonProperty("_a")]
        public string archivo { get; set; }
    }

    /// <summary>
    /// Autor: jlucero
    /// Fecha: 2015-07-09
    /// </summary>
    public class Request_Lectura_Inserta_Masivo
    {
        [JsonProperty("_a")]
        public string lec_archivo { get; set; }

        [JsonProperty("_b")]
        public string lec_asignacion_fecha { get; set; }

        [JsonProperty("_c")]
        public string lec_id { get; set; }

        [JsonProperty("_d")]
        public int usu_id { get; set; }
    }


    /// <summary>
    /// Autor: rcontreras
    /// Fecha: 19-11-2015
    /// </summary>
    public class Request_Lectura_Historico
    {
        [JsonProperty("_a")]
        public int local { get; set; }

        [JsonProperty("_c")]
        public string f_ini { get; set; }

        [JsonProperty("_d")]
        public string f_fin { get; set; }

        [JsonProperty("_e")]
        public string lista { get; set; }
    }

    /// <summary>
    /// Autor: rcontreras
    /// Fecha: 20-11-2015
    /// </summary>
    public class Request_Lectura_Tomadas
    {
        [JsonProperty("_a")]
        public int local { get; set; }

        [JsonProperty("_b")]
        public int tipo { get; set; }

        [JsonProperty("_c")]
        public string f_ini { get; set; }

        [JsonProperty("_d")]
        public string f_fin { get; set; }

        [JsonProperty("_e")]
        public string suministro { get; set; }

        [JsonProperty("_f")]
        public string medidor { get; set; }

        [JsonProperty("_o")]
        public int operario { get; set; }

        [JsonProperty("_g")]
        public string lista { get; set; }

    }
    //objLectura.lec_item = row["item"].ToString();
    //                objLectura.lec_instalacion = row["instalacion"].ToString();
    //                objLectura.lec_Equipo = Convert.ToInt32(row["Equipo"].ToString());
    //                objLectura.lec_Aparato = row["Aparato"].ToString();
    //                objLectura.lec_Tipocalle = row["Tipo de calle"].ToString();
    //                objLectura.lec_NombreCalle = row["Nombre Calle"].ToString();
    //                objLectura.lec_AlturaCalle = row["Altura de Calle"].ToString();
    //                objLectura.NumeroEdificio = row["Numero de Edificio"].ToString();
    //                objLectura.NumeroDepart = row["Numero de Departamento"].ToString();
    //                objLectura.PuntoSumin = row["Detalle adicional de ubicacion (Punto suministo)"].ToString();
    //                objLectura.Piso = row["Piso"].ToString();
    //                objLectura.VivPrin = row["Vivienda Principal (Objeto conexion)"].ToString();
    //                objLectura.DetallCons = row["Detalle Construccion (Objeto de conexion)"].ToString();
    //                objLectura.COnjViv = row["Conjunto de Vivienda (Objeto de conexion)"].ToString();

    public class LeerLecturas
    {



        [JsonProperty("_a")]
        public int INT { get; set; }

        [JsonProperty("_b")]
        public int id_TipoServicio { get; set; }

        [JsonProperty("_c")]
        public DateTime fechaImportacion_corte { get; set; }

        [JsonProperty("_d")]
        public string archivoImportacion_corte { get; set; }

        [JsonProperty("_e")]
        public string suministro_corte { get; set; }

        [JsonProperty("_f")]
        public string medidor_corte { get; set; }

        [JsonProperty("_g")]
        public string claseAviso_corte { get; set; }

        [JsonProperty("_h")]
        public string aviso_corte { get; set; }

        [JsonProperty("_i")]
        public DateTime fechaAviso_corte { get; set; }

        [JsonProperty("_j")]
        public string docBloqueo_corte { get; set; }

        [JsonProperty("_k")]
        public string nombreInterlocutor_corte { get; set; }

        [JsonProperty("_l")]
        public string claseCuenta_corte { get; set; }

        [JsonProperty("_m")]
        public string deudaSoles_corte { get; set; }

        [JsonProperty("_n")]
        public string cantidad_corte { get; set; }


        [JsonProperty("_o")]
        public string nroInstalacion_corte { get; set; }

        [JsonProperty("_p")]
        public string direccion_corte { get; set; }

        [JsonProperty("_q")]
        public int id_ubicacion { get; set; }

        [JsonProperty("_r")]
        public int id_Local { get; set; }


        [JsonProperty("_s")]
        public string unidad_corte { get; set; }

        [JsonProperty("_t")]
        public string ejecutante_corte { get; set; }

        [JsonProperty("_u")]
        public int creadoPor_corte { get; set; }

        [JsonProperty("_v")]
        public String statusSistem_corte { get; set; }


        [JsonProperty("_w")]
        public string codificacion_corte { get; set; }

        [JsonProperty("_x")]
        public DateTime fechaAsignacion_corte { get; set; }

        [JsonProperty("_y")]
        public DateTime id_Operario_corte { get; set; }

        [JsonProperty("_z")]
        public string limite_Minimo_corte { get; set; }


        [JsonProperty("_ab")]
        public string limite_Maximo_corte { get; set; }

        [JsonProperty("_ac")]
        public int orden_corte { get; set; }

        [JsonProperty("_ad")]
        public int reordenamiento_corte { get; set; }


        [JsonProperty("_ae")]
        public DateTime fechaLecturaMovil_corte { get; set; }

        [JsonProperty("_af")]
        public DateTime fechaRecepcion_corte { get; set; }

        [JsonProperty("_ag")]
        public string LecturaMovil_corte { get; set; }

        [JsonProperty("_ah")]
        public string confirmacion_corte { get; set; }

        [JsonProperty("_ai")]
        public string Lectura_Manual_corte { get; set; }

        [JsonProperty("_aj")]
        public string Observacion_corte { get; set; }

        [JsonProperty("_ak")]
        public string nombreCliente_corte { get; set; }

        [JsonProperty("_al")]
        public string tieneFoto_corte { get; set; }


        [JsonProperty("_am")]
        public string latitud_corte { get; set; }


        [JsonProperty("_an")]
        public string longitud_corte { get; set; }

        [JsonProperty("_ao")]
        public DateTime fechaEnvioCliente_corte { get; set; }

        [JsonProperty("_ap")]
        public string id_Observacion_corte { get; set; }

        [JsonProperty("_aq")]
        public string id_Operario_ReConexion { get; set; }



        [JsonProperty("_ar")]
        public int estado { get; set; }

        [JsonProperty("_as")]
        public int usuario_creacion { get; set; }

        [JsonProperty("_at")]
        public int usuario_edicion { get; set; }

        [JsonProperty("_au")]
        public int fecha_edicion { get; set; }

        [JsonProperty("_aw")]
        public string distrito_corte { get; set; }



        [JsonProperty("_br")]
        public int contador { get; set; }

        [JsonProperty("bt")]
        public int tecnico { get; set; }

        [JsonProperty("_bu")]
        public int secuencia { get; set; }

        [JsonProperty("_cy")]
        public string lec_importacion_archivo { get; set; }

    }



    ///// <summary>
    ///// Autor: jlucero
    ///// Fecha: 2015-07-09
    ///// </summary>
    //public class Request_Elimina_Masivo
    //{
    //    [JsonProperty("_a")]
    //    public string archivo { get; set; }

    //    [JsonProperty("_b")]
    //    public int opcion { get; set; }
    //}

    #endregion
}