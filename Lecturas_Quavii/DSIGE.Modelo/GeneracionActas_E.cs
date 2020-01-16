using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DSIGE.Modelo
{
    public class GeneracionActas_E
    {

        public int id_TipoServicio { get; set; }
        public string nombre_tiposervicio { get; set; }

        public string fecha  { get; set; }
        public string hora { get; set; }
        public string fechaLecturaMovil_corte { get; set; }
        public string cliente_suministro { get; set; }
        public string instalacion { get; set; }
        public string medidor { get; set; }
        public string ubicacion_medidor_interno { get; set; }
        public string ubicacion_medidor_externo { get; set; } 

        public string nombre_cliente { get; set; }
        public string direccion_suministro { get; set; }
        public string distrito { get; set; }
        public string nombre_personal { get; set; }
        public string dni_personal { get; set; }
        public string lectura_cierre { get; set; }
        
        public string motivo_cierre_deuda { get; set; }
        public string motivo_cierre_fuga_gas { get; set; }
        public string motivo_cierre_pedido_cliente { get; set; }
        public string motivo_cierre_seguridad { get; set; }

        public string resultado_cierre_ejecutado { get; set; }
        public string resultado_cierre_ausente { get; set; }
        public string resultado_cierre_acceso { get; set; }
        public string resultado_cierre_resistencia { get; set; }

        public string observaciones { get; set; } 
        
        public string ruta_foto_1 { get; set; }
        public string ruta_foto_2 { get; set; }
        public string ruta_foto_3 { get; set; }
        public string ruta_foto_4 { get; set; }

        public string motivo_deuda { get; set; }
        public string motivo_apc { get; set; }
        public string motivo_tecnico { get; set; }
        public string motivo_otros { get; set; }
        public string recone_exitosa { get; set; }
        public string exitosa_lectura { get; set; }
        public string artefac_cocina { get; set; }
        public string artefac_terma { get; set; } 
        public string artefac_otro { get; set; }

        public string recone_infrutuosa { get; set; }
        public string motivo_cliente_ausente { get; set; }
        public string motivo_cliente_impide { get; set; }

        public string motivo_sin_acceso { get; set; }
        public string motivo_cliente_otroCheck { get; set; }


        public string motivo_cliente_otros { get; set; }
        public string primera_visita_recone { get; set; }
        public string proxima_visita_fecha_check { get; set; }        
        public string proxima_visita_fecha { get; set; }
        public string proxima_visita_hora { get; set; }       
        public string proxima_visita_solicita { get; set; }

    }
}
