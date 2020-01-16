using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DSIGE.Modelo
{
    public class DistribuirLecturas_E
    {

        public int index { get; set; }
        public bool checkeado { get; set; }
        public bool check { get; set; }
        public bool flag_operario { get; set; }
        public string id_Lectura { get; set; }
        public string colorItem { get; set; }

        public bool disabled { get; set; }
        public string Unidad_Lecturas  { get; set; }
        public string Cantidad_Manzana { get; set; }
        public string Cantidad { get; set; }
        public string Codigo_operario { get; set; }
        public string fechaAsignacion_Lectura { get; set; }
        public string Manzana { get; set; }
        public string cod_observacion { get; set; }      
        
        public string SECTOR { get; set; }
        public string CORR { get; set; }
        public string ZONA { get; set; }
        public int CANTIDAD { get; set; }
        public string FECHA { get; set; }
        public int id_Local { get; set; }
        public int LOCAL { get; set; }
        public string nombre_local { get; set; }
        public int id_TipoServicio { get; set; }
        public string flag_relectura { get; set; }
        
        public string nombre_tiposervicio { get; set; }
        public int estado { get; set; }
        public int id_Operario { get; set; }
        public int CANT_REGISTROS { get; set; }


        public string suministro_lectura { get; set; }
        public string medidor_lectura { get; set; }
        public string LecturaMovil_Lectura { get; set; }
        public string Apellidos_Operario { get; set; }
        public string Operario { get; set; }
        public string direccion_lectura { get; set; }
        public string nombreCliente_lectura { get; set; }
        public string email_operario { get; set; }

        

        public int id_grupo { get; set; }
        public int id_Operario_Lectura { get; set; }
        public string operario { get; set; }
        public string R1 { get; set; }
        public string SUMINISTRO { get; set; }
        public string MEDIDOR { get; set; }
        public string CORRELATIVO { get; set; }
        public string LECTURA { get; set; }
        public string DIRECCION { get; set; }
        public string CLIENTE { get; set; }

        public string id_anio { get; set; }
        public string descripcion_anio { get; set; }

        public string id_mes { get; set; }
        public string descripcion_mes { get; set; }

        public string r1_lectura { get; set; }
        public string R1_lectura { get; set; }
        public string S1_lectura { get; set; }
        public int cod_lector { get; set; }
        public string CC_lectura { get; set; }
        public string Sector_lectura { get; set; }
        public string Zona_lectura { get; set; }
        public string fecha_Asignacion { get; set; }

        public decimal totalLibro { get; set; }
        public decimal TotalRuta { get; set; }



    }
}
