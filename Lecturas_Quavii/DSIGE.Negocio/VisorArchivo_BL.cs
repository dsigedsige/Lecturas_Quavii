using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DSIGE.Modelo;
using DSIGE.Dato;

namespace DSIGE.Negocio
{
    public class VisorArchivo_BL 
    {
      public string Capa_Negocio_GuardarArchivos(int id_empresa, int usuario, string local, string servicio, string tipodoc, string contrato, string fecharegistro, string nombreArchivo, string url)
        {
            VisorArchivo_DAO Objeto_Dato = new VisorArchivo_DAO();
            return Objeto_Dato.Capa_Dato_GuardarArchivos(id_empresa, usuario, local, servicio, tipodoc, contrato, fecharegistro, nombreArchivo, url);
        }      

        public List<ImportarArchivo> Capa_Negocio_Get_Listado_Locales()
        {
            try
            {
                VisorArchivo_DAO Objeto_Dato = new VisorArchivo_DAO();
                return Objeto_Dato.Capa_Dato_Get_Locales();
            }
            catch (Exception e)
            {
                throw e;
            }
        }

        public List<ImportarArchivo> Capa_Negocio_Get_Listado_Servicios()
        {
            try
            {
                VisorArchivo_DAO Objeto_Dato = new VisorArchivo_DAO();
                return Objeto_Dato.Capa_Dato_Get_Servicio();
            }
            catch (Exception e)
            {
                throw e;
            }
        }

        public List<ImportarArchivo> Capa_Negocio_Get_ListandoDetallesDocumentos(int id_empresa, string id_local, string id_servicio, string fecha_ini, string fecha_fin, string contrato)
        {
            try
            {
                   VisorArchivo_DAO Objeto_Dato = new VisorArchivo_DAO();
                   return Objeto_Dato.Capa_Dato_Get_ListandoDetallesDocumentos(id_empresa, id_local, id_servicio, fecha_ini, fecha_fin, contrato);
            }
            catch (Exception e)
            {
                throw e;
            }
        }

        public List<ImportarArchivo> Capa_Negocio_Get_RutaArchivos(string id_Archivo)
        {
            try
            {
                VisorArchivo_DAO Objeto_Dato = new VisorArchivo_DAO();
                return Objeto_Dato.Capa_Dato_Get_RutaArchivos(id_Archivo);
            }
            catch (Exception e)
            {
                throw e;
            }
        }



    }
}
