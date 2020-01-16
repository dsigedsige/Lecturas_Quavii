using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DSIGE.Modelo;
using DSIGE.Dato;

namespace DSIGE.Negocio
{
    public class ImportarArchivo_BL
    {
      public string Capa_Negocio_GuardarArchivos(int id_empresa, int usuario, string local, string servicio, string tipodoc, string contrato, string fecharegistro, string nombreArchivo, string url)
        {
            ImportarArchivos_DAO Objeto_Dato = new ImportarArchivos_DAO();
            return Objeto_Dato.Capa_Dato_GuardarArchivos(id_empresa, usuario, local, servicio, tipodoc, contrato, fecharegistro, nombreArchivo, url);
        }
      

        public List<ImportarArchivo> Capa_Negocio_Get_Listado_Locales()
        {
            try
            {
                ImportarArchivos_DAO Objeto_Dato = new ImportarArchivos_DAO();
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
                ImportarArchivos_DAO Objeto_Dato = new ImportarArchivos_DAO();
                return Objeto_Dato.Capa_Dato_Get_Servicio();
            }
            catch (Exception e)
            {
                throw e;
            }
        }

        public List<ImportarArchivo> Capa_Negocio_Get_TipoDocumentos()
        {
            try
            {
                ImportarArchivos_DAO Objeto_Dato = new ImportarArchivos_DAO();
                return Objeto_Dato.Capa_Dato_Get_TipoDocumento();
            }
            catch (Exception e)
            {
                throw e;
            }
        }



    }
}
