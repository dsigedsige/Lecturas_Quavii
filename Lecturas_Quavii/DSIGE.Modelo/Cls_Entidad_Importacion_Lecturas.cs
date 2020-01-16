using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DSIGE.Modelo
{
    public class Cls_Entidad_Importacion_Lecturas
    {
        public string CLIENTE { get; set; }
        public string DIRECCION { get; set; }
        public string Sc { get; set; }
        public string Zn { get; set; }
        public string CRR { get; set; }
        public string MEDIDOR { get; set; }
        public string MARCA { get; set; }
        public string MesUno { get; set; }
        public string MesDos { get; set; }
        public string MesTres { get; set; }
        public string MesCuatro { get; set; }
        public string MesCinco { get; set; }
        public string DNI { get; set; }
        public string Nombre_Operador { get; set; }
        public string Bloque { get; set; }
        public int Orden { get; set; }
        public int loc_id { get; set; }
        public int id_UsuarioExpor { get; set; }
        public int Promedio { get; set; }
        public int Minimo { get; set; }
        public int Maximo { get; set; }
        public int Lect_Minima { get; set; }
        public int Lect_Maxima { get; set; }
        public int Ultimas_Lect_3 { get; set; }
        public string FLAG_LECTURA { get; set; }
        public int cant_correcto { get; set; }
        public int cant_erroneo { get; set; }

        public int reg_importado { get; set; }
        public int reg_correctos { get; set; }
        public int reg_incorrecto { get; set; }
        public int diferencia { get; set; }

        public string SUMINISTRO { get; set; }
        public string CAMBIO_FECHA { get; set; }











    }
    public class Cls_Entidad_Lecturas_Dia
    {
        public string Tipo { get; set; }
        public int Estado { get; set; }
        public int id_tipoServicio { get; set; }
        public int id_lectura { get; set; }
        public string suministro_lectura { get; set; }
        public string latitud_lectura { get; set; }
        public string longitud_lectura { get; set; }
        public string Color { get; set; }
        public string Medidor { get; set; }
        public string Direccion { get; set; }
        public string Operador { get; set; }
        public string Cliente { get; set; }
        public string Lectura { get; set; }
        public string FechaHora { get; set; }
        public string TomaFoto { get; set; }
        public string Url { get; set; }
        

    }

    public class Cls_Entidad_Lectura_OperarioZona
    {
        
        public int id_operario_lectura { get; set; }
        public int cant_zona { get; set; }
        public int cant_sum { get; set; }
        public string nroDoc_Operario { get; set; }

    }

    public class Cls_Entidad_Temporal_Sum
    {
        public string id_Suministro { get; set; }
        public int id_Operario { get; set; }
        public int Usuario { get; set; }
    }

    public class Cls_Entidad_Resumen_Personal
    {
        public string Cliente { get; set; }
        public string Operario { get; set; }
        public string Asignado { get; set; }
        public string Pendientes { get; set; }
        public string Ejecutado { get; set; }
        public string Efectivos { get; set; }
        public string NoEfectivos { get; set; }
        public string Avance { get; set; }
        public string HoraInicio { get; set; }
        public string HoraFinal { get; set; }
        public string HorasGeneral { get; set; }
        public string Promedio { get; set; }

    }

    public class Cls_Entidad_Lecturas_Trabajo
    {
        public int id_Operario { get; set; }
        public string apellidos_operario { get; set; }
        public string Asignado { get; set; }
        public string Ejecutado { get; set; }
        public string PendDia { get; set; }
        public string Porcentaje { get; set; }
        public string Latitud { get; set; }
        public string Longitud { get; set; }
        public string Minimo { get; set; }
        public string Maximo { get; set; }
        public string totalHora { get; set; }
        public string nroDoc_Operario { get; set; }
        

    }
    public class Cls_Entidad_Lecturas_Relecturas
    {
        public string item { get; set; }
        public string instalacion { get; set; }
        public string equipo { get; set; }
        public string aparato { get; set; }
        public string tipoCalle { get; set; }
        public string direccion_lectura { get; set; }
        public string altura_calle { get; set; }
        public string numeroEdificio { get; set; }
        public string numeroDepartamento { get; set; }
        public string detalleAdicionalUbicacion { get; set; }
        public string piso { get; set; }
        public string viviendaPrincipal { get; set; }
        public string detalleConstruccion { get; set; }
        public string conjuntoVivienda { get; set; }
        public string manzana { get; set; }
        public string cor_distrito { get; set; }
        public string codigoPostal { get; set; }
        public string poblacion { get; set; }
        public string emplazamiento { get; set; }
        public string sumplementoEmplazamiento { get; set; }
        public string lecturaAnterior { get; set; }
        public DateTime fechaPlanLectAngterior { get; set; }
        public DateTime fechaPlanLectActual { get; set; }
        public string fechaPlanLectProxima { get; set; }
        public string interlocutor { get; set; }
        public string cuentaContrato { get; set; }
        public string tipocliente { get; set; }
        public string categoria { get; set; }
        public string secuenciaLectura { get; set; }
        public string unidadLectura { get; set; }
        public string numeroLecturaEstimada { get; set; }
        public string marcaPrimeraLectura { get; set; }
        public string EmpresaLectora { get; set; }
        public string nota1UbicacionAparato { get; set; }
        public string nota2UbicacionAparato { get; set; }
        public int tecnico { get; set; }
        public string nombre_ope { get; set; }
        public int secuencia { get; set; }
        public int grupo { get; set; }
        public int idUsuarioExport { get; set; }
        public int loc_id { get; set; }
        public DateTime fechaAsignacion { get; set; }
        public int cant_correcto { get; set; }
        public int cant_erroneo { get; set; }
        public string suministro_lectura { get; set; }
        public int id_tiposervicio { get; set; }
        public string distritoLectura { get; set; }
        
    }
    public class Cls_Entidad_Temporal_Lectura_Sum
    {
        public string Suministro { get; set; }
        public string Sector { get; set; }
        public string LectAnterior { get; set; }
        public string Promedio { get; set; }
        public string LecturaMinima { get; set; }
        public string LecturaMaxima { get; set; }
        public string TipoServicio { get; set; }
        public string fechaAsignacion { get; set; }
        public int cantidad { get; set; }
    }

}
