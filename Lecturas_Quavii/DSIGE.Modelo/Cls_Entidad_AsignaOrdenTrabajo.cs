using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DSIGE.Modelo
{
    public class Cls_Entidad_AsignaOrdenTrabajo
    {
        public class Locales
        {
            public int loc_id { get; set; }
            public int emp_id { get; set; }
            public string loc_nombre { get; set; }
            public string loc_direccion { get; set; }
            public int loc_orden { get; set; }
            public string loc_latitud { get; set; }
            public string loc_longitud { get; set; }
            public int loc_estado { get; set; }
        }

        public class validacion
        {           
            public bool checkeado { get; set; }
            public string  estado { get; set; }
             public string  id_Lectura { get; set; }
             public string  foto { get; set; }
             public string  fecha_lectura { get; set; } 
             public string  suministro_lectura { get; set; } 
             public string  medidor_lectura { get; set; } 
             public string  limite_Minimo_lectura { get; set; } 
             public string  limite_Maximo_lectura { get; set; } 
             public string  operario { get; set; } 
             public string  id_observacion { get; set; } 
             public string  descripcion_observacion { get; set; } 
             public string Consu_02 { get; set; } 
             public string  Consu_act { get; set; } 
             public string  Porcen { get; set; } 
             public string  L_ant { get; set; } 
             public string  L_ { get; set; } 
             public string  Consu_a { get; set; }
            public string lectura_movil { get; set; }
            
            public string lectura_max { get; set; }
            public string lectura_min { get; set; }
            public string comparativa_max { get; set; }
            public string comparativa_min { get; set; }

            public string  Prom_3mes { get; set; } 
            public string  Consu_01 { get; set; } 
            public string  porce { get; set; }
            public string  Relecturas { get; set; }
            public string resultado { get; set; }
            public string Observacion_lectura { get; set; }

            
        }


        public class Servicios
        { 
            public string id_Configuracion_UL { get; set; }
            public string Dia_Configuracion_UL { get; set; }
            public string responsable { get; set; }
            public string cantidad_lectura { get; set; }
            public string Estado { get; set; }

            public string usuario_creacion { get; set; }
            public string fecha_creacion { get; set; }
            public string usuario_modificacion { get; set; }
            public string fecha_modificacion { get; set; }

            public string Cod_UnidadLectura { get; set; }
            public string nombre_UnidadLectura { get; set; }
            public string Distrito_UnidadLectura { get; set; }

            public string id_anio { get; set; }
            public string descripcion_anio { get; set; }
            public string id_mes { get; set; }
            public string descripcion_mes { get; set; }


            public bool checkeado { get; set; }
            public int id_Usuario_Responsable { get; set; }
            public int id_Usuario { get; set; }
            public string supervisor { get; set; }

            public string usuario_registra { get; set; }
            public string usuario_edicion { get; set; }

            public int id_Operario { get; set; }
            public string desc_operario { get; set; }

            public int id_TipoServicio { get; set; }
            public string nombre_tiposervicio { get; set; }
            public int estado { get; set; }
            public int cantidad { get; set; }

        }

        public class Estados
        {
            public int id_Estado { get; set; }
            public string descripcion_estado { get; set; }
        }

        public class Observaciones
        {
            public int id_Observacion { get; set; }
            public string descripcion_observacion { get; set; }

            public int id_Lectura { get; set; }

            public int id_lectura_origen { get; set; }
            public string unidad_lectura    { get; set; }
            public string suministro_lectura { get; set; }
            public string medidor_lectura   { get; set; }
            public string direccion_lectura  { get; set; }
            public string manzana_lectura   { get; set; }
            public string cod_observacion { get; set; }
            public string lectura_anterior { get; set; }
            public string cod_observacion_anterior { get; set; }            

            public string lectura { get; set; }
            public string descripcion_validacion { get; set; }
            public string consumo1 { get; set; }
            public string consumo2 { get; set; }

            public string operario { get; set; }
            public string id_observacion { get; set; }

            public string lectura_movil { get; set; }
            public string lectura_mxa { get; set; }
            public string lectura_min { get; set; }

            public string comparativa_max { get; set; }
            public string comparativa_min { get; set; }
            public string resultado { get; set; }



        }

        public class Tecnico
        {
            public int ope_id { get; set; }
            public string ope_nombre { get; set; }
        }

        public class OrdenesTrabajo
        {
            public bool checkeado { get; set; }
            public int id_Lectura { get; set; }
            public string suministro_lectura { get; set; }
            public string medidor_lectura { get; set; }
            public string direccion_lectura { get; set; }
            public string lectura_Anterior { get; set; }
            public string limite_Minimo_lectura { get; set; }
            public string limite_Maximo_lectura { get; set; }


            public string id_Operario_Lectura { get; set; }
            public string ope_nombre { get; set; }
            public string Zona_lectura { get; set; }
            public string fechaAsignacion_Lectura { get; set; }
            public string nombreCliente_lectura { get; set; }
            public int estado { get; set; }
            public string abreviatura_estado { get; set; }
            public string nroInstalacion_lectura { get; set; }



            public int ord_asig { get; set; }
            public string fecha_importacion { get; set; }
            public string solicitud { get; set; }
            public string cuenta_contrato { get; set; }
            public string nro_serie_medidor { get; set; }
            public string cliente { get; set; }
            public string direc_instalacion { get; set; }
            public string distrito { get; set; }
            public string UL { get; set; }
            public string apellido_lector { get; set; }
            public string grand_total { get; set; }
            public string count { get; set; }
            public string accion { get; set; }
            public string resultado_corte { get; set; }
            public string fecha_corte { get; set; }
            public string mes_corte { get; set; }
            public string hora_corte { get; set; }
            public string lectura_corte { get; set; }
            public string tipo_corte { get; set; }
            public string cod_OBS { get; set; }
            public string observacion_lectura { get; set; }
            public string detalle_observaciones { get; set; }

            public int id_Operario_corte { get; set; }
            public int id_Corte { get; set; }
            public string suministro_corte { get; set; }

        }

        public class LecturaEnvio
        {

            public string CL { get; set; }
            public string OPERARIO { get; set; }
            public decimal CANT_ASIGNADO { get; set; }
            public decimal CANT_REALIZADO { get; set; }
            public decimal CANT_PENDIENTE { get; set; }

            public string direccion { get; set; }
            public string medidor_corte { get; set; }
            public string fechLectura { get; set; }
            public string lecMovil { get; set; }
            public string consuActu { get; set; }
            public string lecProm { get; set; }
            public string lecPorcen { get; set; }
            public string lecVali { get; set; }
            public string lecManual { get; set; }
            public string idLecObs { get; set; }
            public string lecObs { get; set; }

            public bool checkeado { get; set; }
            public int id_Lectura { get; set; }
            public string foto { get; set; }
            public string suministro_lectura { get; set; }
            public string medidor_lectura { get; set; }
            public string direccion_lectura { get; set; }
            public string lectura_Anterior { get; set; }
            public string limite_Minimo_lectura { get; set; }
            public string limite_Maximo_lectura { get; set; }

            public string ope_nombre { get; set; }
            public string Zona_lectura { get; set; }
            public string fechaAsignacion_Lectura { get; set; }
            public string nombreCliente_lectura { get; set; }
            public int estado { get; set; }
            public string abreviatura_estado { get; set; }
            public string nroInstalacion_lectura { get; set; }

            public int ord_asig { get; set; }
            public string fecha_importacion { get; set; }
            public string solicitud { get; set; }
            public string cuenta_contrato { get; set; }
            public string nro_serie_medidor { get; set; }
            public string cliente { get; set; }
            public string direc_instalacion { get; set; }
            public string distrito { get; set; }
            public string UL { get; set; }
            public string apellido_lector { get; set; }
            public decimal grand_total { get; set; }
            public int count { get; set; }
            public string accion { get; set; }
            public string resultado_corte { get; set; }
            public string fecha_corte { get; set; }
            public string mes_corte { get; set; }
            public string hora_corte { get; set; }
            public string lectura_corte { get; set; }
            public string tipo_corte { get; set; }
            public string cod_OBS { get; set; }
            public string observacion_lectura { get; set; }
            public string detalle_observaciones { get; set; }

            public int id_Operario_corte { get; set; }
            public int id_Corte { get; set; }
            public string suministro_corte { get; set; }

            public string nroEquipo_lectura { get; set; }
            public string CuentaContrato { get; set; }
            public string FechaPlanLectura { get; set; }
            public string FechaRealLectura { get; set; }
            public string HoraLecturaReal { get; set; }
            public string Lectura { get; set; }
            public string NotaLecturista { get; set; }
            public string CodigoComentario { get; set; }
            public string Comentario { get; set; }
            public string id_Operario_Lectura { get; set; }
            public string CodigoLector { get; set; }

            public string unidadLectura { get; set; }
            public string ejecutante { get; set; }
            public string orden { get; set; }
            public string creado { get; set; }
            public string estatusSistema { get; set; }
            public string codCodificacion { get; set; }
            public string codificacion { get; set; }
            public string contador { get; set; }
            public string tcos { get; set; }
            public string hora { get; set; }
            public string lectura { get; set; }
            public string imposibilidad { get; set; }
            public string observacion { get; set; }
            public string Deuda_Soles { get; set; }


            public string claseAviso { get; set; }
            public string aviso { get; set; }
            public string fechaAviso { get; set; }
            public string docBloqueo { get; set; }
            public string ctaContrato { get; set; }
            public string nombreInterlocutor { get; set; }
            public string claseCuenta { get; set; }

            public string cantRecibos { get; set; }
            public string instalacion { get; set; }
            public string nroSerieMedidor { get; set; }
            public string direcInstalacion { get; set; }
            public string distritoInstalacion { get; set; }

            public string Ejecutante { get; set; }
            public string creadoPor { get; set; }
            public string estadoInstalacion { get; set; }
            public string estatusUsuario { get; set; }
            public string nombreFoto { get; set; }
            public string fotourl { get; set; }

            public string nombreCliente { get; set; }
            public string id_observacionResultado_corte { get; set; }
            public string descripcion_observacion { get; set; }

            public string id_Observacion_corte { get; set; }
            public string id_resultadoObsCorte { get; set; }

            public string fecha { get; set; }
            public string desplazamiento { get; set; }
            public string ubicacion_Medidor { get; set; }
            public string tieneFoto { get; set; }            

            public string Lectura1 { get; set; }
            public string Consumo1 { get; set; }
            public string Nota_Lect1 { get; set; }
            public string Estado1 { get; set; }
            public string Lectura2 { get; set; }
            public string Consumo2 { get; set; }
            public string Nota_Lect2 { get; set; }
            public string Estado2 { get; set; }
            public string Lectura3 { get; set; }
            public string Consumo3 { get; set; }
            public string Nota_Lect3 { get; set; }
            public string Estado3 { get; set; }
            public string Lectura4 { get; set; }
            public string Consumo4 { get; set; }
            public string Nota_Lect4 { get; set; }
            public string Estado4 { get; set; }
            public string Lectura5 { get; set; }
            public string Consumo5 { get; set; }
            public string Nota_Lect5 { get; set; }
            public string Estado5 { get; set; }
            public string Lectura6 { get; set; }
            public string Consumo6 { get; set; }
            public string Nota_Lect6 { get; set; }
            public string Estado6 { get; set; }

            public decimal Promedio { get; set; }
            public decimal Menos20 { get; set; }
            public decimal Mas20 { get; set; }
            public string Resultado { get; set; }
            public string ConsumoProm { get; set; }
            

            public int id_Reparto { get; set; }
            public string DocImpresion_Reparto { get; set; }
            public string NroRecibo_Reparto { get; set; }
            public string CuentaContrato_Reparto { get; set; }
            public string nombreCliente_Reparto { get; set; }
            public string direccion_Reparto { get; set; }
            public string fecha_reparto { get; set; }
            public string aviso_Reparto { get; set; }
            public string cargo_Reparto { get; set; }
            public string muestra_Reparto { get; set; }
            public string ConsumoActual { get; set; }
            public string Porcentaje { get; set; }
            public string PorcentajeMayo { get; set; }
            public string LecturaMenor { get; set; }
 
        }


        public class VisualizarPendiente
        {
            public string fecha { get; set; }
            public string nota_lectura { get; set; }
            public string estimacion { get; set; }
            public string parametros { get; set; }
            public string comercio { get; set; }
            public string total { get; set; }

            public string id_operario { get; set; }
            public string Operario { get; set; }
            public string id_estado { get; set; }
            public string Estado { get; set; }
            public int Cantidad { get; set; }
            public string fechaAsignacion { get; set; }

            public string Orden { get; set; }
            public string Cuenta_Contrato { get; set; }
            public string Nro_Instalacion { get; set; }
            public string Nro_Serie { get; set; }
            public string Direccion { get; set; }
            public string Ubicacion_Med { get; set; }
            public string Fecha_Alta { get; set; }
            public string Cliente { get; set; }
            public string Inter_Locutor { get; set; }


        }


        public class Historico_LecturaRelectura
        {
            public string medidor_lectura { get; set; }
            public string fechaLecturaMovil_lectura { get; set; }
            public string lectura_Anterior { get; set; }
            public string ope_nombre { get; set; }
            public string Observacion_lectura { get; set; }
            public string abreviatura_estado { get; set; }
        }

        public class Historico_CortesReconexiones
        {
            public string medidor_corte { get; set; }
            public string fechaAsignacion_corte { get; set; }
            public string id_Corte { get; set; }
            public string ope_nombre { get; set; }
            public string Observacion_corte { get; set; }
            public string abreviatura_estado { get; set; }
        }

        public class Trabajos_Cliente
        {
            public int ITEM { get; set; }
            public string INSTALACION { get; set; }
            public string EQUIPO { get; set; }
            public string MEDIDOR { get; set; }
            public string CTACTO { get; set; }
            public string FecPlanLectura { get; set; }
            public string FecRealPlanLectura { get; set; }
            public string Hora_Plantilla { get; set; }
            public string Lectura { get; set; }
            public string Nota { get; set; }
            public string Cod_Comentario { get; set; }
            public string Comentario { get; set; }
            public string Codigo_Lector { get; set; }
        }
    }
}
