using DSIGE.Dato;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DSIGE.Negocio
{
    public class GeneracionActas_BL
    {
        public object Capa_Negocio_Servicio()
        {
            GeneracionActas_DAO Objeto_Dato = new GeneracionActas_DAO();
            return Objeto_Dato.Capa_Dato_Servicios();
        }

        public object Capa_Negocio_Mostrando_informacion_general(int local, int servicio, int operario, string fecha)
        {
            GeneracionActas_DAO Objeto_Dato = new GeneracionActas_DAO();
            return Objeto_Dato.Capa_Dato_Mostrando_informacion_general(local, servicio, operario, fecha);
        }


        public object Capa_Negocio_MostrandoActa(int idCorte, int idServicio)
        {
            GeneracionActas_DAO Objeto_Dato = new GeneracionActas_DAO();
            return Objeto_Dato.Capa_Dato_MostrandoActa(idCorte, idServicio);
        }



        //public object Capa_Negocio_Mostrando_informacion_general(int servicio, int operario, string fecha)
        //{
        //    GeneracionActas_DAO Objeto_Dato = new GeneracionActas_DAO();
        //    return Objeto_Dato.getCrearPDF_actas();
        //}


        public object Capa_Negocio_Mostrando_informacion_Inspecciones(int servicio, int operario, string fecha, int tipoReporte)
        {
            GeneracionActas_DAO Objeto_Dato = new GeneracionActas_DAO();
            return Objeto_Dato.Capa_Dato_Mostrando_informacion_Inspecciones(servicio, operario, fecha, tipoReporte);
        }

        public object Capa_negocio_get_generacionPdf_inspecciones(int id_inspeccion)
        {
            GeneracionActas_DAO Objeto_Dato = new GeneracionActas_DAO();
            return Objeto_Dato.Capa_dato_get_generacionPdf_inspecciones(id_inspeccion);
        }


    }
}
