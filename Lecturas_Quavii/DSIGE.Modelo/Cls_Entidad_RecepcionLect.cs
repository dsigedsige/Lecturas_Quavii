using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DSIGE.Modelo
{
    public class Cls_Entidad_RecepcionLect
    {
            public string id_Lectura { get; set; }
            public int  id_TipoServicio { get; set; }
            public string nombre_tiposervicio { get; set; }
            public int  estado { get; set; }
            public int id_Operario { get; set; }
            public string operario { get; set; }
            public string suministro_lectura { get; set; }
            public string medidor_lectura { get; set; }
            public string direccion_lectura { get; set; }
            public string nombreCliente_lectura { get; set; }
            public string operario_lectura { get; set; }
            public string orden_lectura { get; set; }
            public string lectura_Anterior { get; set; }
            public string LecturaMovil_Lectura { get; set; }
            public string id_Observacion_Lectura { get; set; }
            public string r1_lectura { get; set; }
            public string s1_lectura { get; set; }
            public string Sector_lectura { get; set; }
            public string Zona_lectura { get; set; }
            public string descripcion_estado { get; set; }
            public string DNS_lectura { get; set; }
            public string CodUbicacion_lectura { get; set; }
            public string DeviceId { get; set; }
            public string VersionApp { get; set; }
            public int flag_limpiar_historial { get; set; }

            public string lectura_bak { get; set; }
            public string suministro { get; set; }
            public string fecha_movil { get; set; }
            public string fecha_servidor { get; set; }

            public int id_Estado { get; set; }

            public string usuario_modificacion { get; set; }
            public string fecha_modificacion { get; set; }
            public string URL { get; set; }
 
    }

}
