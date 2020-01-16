using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DSIGE.Modelo
{
  public  class Cls_Entidad_Export_trabajos_lectura
    {

        public int id_Lectura { get; set; }
        public string Instalacion { get; set; }
        public string Equipo { get; set; }
        public string Aparato { get; set; }


        public string direccion_lectura { get; set; }

        public string Tipo_calle { get; set; }
        public string Nombre_Calle { get; set; }
        public string Altura_Calle { get; set; }
        public string Numero_Edificio { get; set; }
        public string Numero_Departamento { get; set; }
        public string Detalle_adicional_ubicacion { get; set; }
        public string Piso { get; set; }
        public string Vivienda_Principal { get; set; }
        public string Detalle_Construccion { get; set; }
        public string Conjunto_Vivienda { get; set; }
        public string Manzana_Lote { get; set; }

        public string Distrito { get; set; }
        public string Codigo_postal { get; set; }
        public string Poblacion { get; set; }
        public string Emplazamiento { get; set; }
        public string Suplemento_emplazamiento { get; set; }
        public string Lectura_anterior { get; set; }
        public string Fecha_planificada_lectura_anterior { get; set; }
        public string Fecha_planificada_lectura_actual { get; set; }
        public string Fecha_planificada_lectura_proxima { get; set; }
        public string Interlocutor_comercial { get; set; }
        public string Cuenta_contrato { get; set; }
        public string Tipo_Cliente { get; set; }
        public string Categoria { get; set; }
        public string Secuencia_lectura { get; set; }
        public string Unidad_lectura { get; set; }
        public string Numero_lecturas_estimadas_consecutivas { get; set; }
        public string Marca_primera_lectura { get; set; }
        public string Empresa_Lectora { get; set; }
        public string Nota_1_ubicacion_aparato { get; set; }
        public string Nota_2_ubicacion_aparato { get; set; }
        public string Tecnico { get; set; }
        public string Secuencia { get; set; }
        public string Grupo { get; set; }

    }
}
