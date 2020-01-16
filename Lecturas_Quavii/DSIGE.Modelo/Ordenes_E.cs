using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DSIGE.Modelo
{
    public class Ordenes_E
    {
        public string Unidad_Lecturas { get; set; }
        public string Codigo_operario { get; set; }
        public string opcion { get; set; }
    }
    //public class OrdenesDetalle_E
    //{
    //    public int id_Lectura { get; set; }
    //    public string Codigo_operario { get; set; }
    //    public string opcion { get; set; }
    //}

    public class OrdenesDetalle_E
    {
        public string id_Lectura { get; set; }
        public string Unidad_Lecturas { get; set; }
        public string Codigo_operario { get; set; }
        public string Manzana { get; set; }
        public string opcion { get; set; }
        public string flag_detallado { get; set; }
        
    }

    public class Archivo_E
    {
        public string email { get; set; }
        public string ruta { get; set; }
        public string nombreFile { get; set; }
 
    }

    public class RepartoDetalle
    {
        public int id_Reparto { get; set; }
        public int id_operario { get; set; }
        public int id_operario_inicial { get; set; }
        public string manzana { get; set; }
        public string unidad_distrito { get; set; }

    }

    public class CorteReconexionDetalle
    {
        public int id_Corte { get; set; }
        public int id_operario { get; set; }

    }

    public class GrandesClientesDetalle
    {
        public int Id_GrandeCliente { get; set; }
        public int id_operario { get; set; }
        public int id_operario_inicial { get; set; }
        public string direccion_lectura { get; set; }
 
    }


}
 