using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DSIGE.Modelo;
using DSIGE.Dato;

namespace DSIGE.Negocio
{
  public class VisorNovedades_BL
    {

        public string Capa_Negocio_GuardarArchivos(int id_empresa, int usuario, string local, string servicio, string tipodoc, string contrato, string fecharegistro, string nombreArchivo, string url)
        {
            VisorNovedades_DAO Objeto_Dato = new VisorNovedades_DAO();
            return Objeto_Dato.Capa_Dato_GuardarArchivos(id_empresa, usuario, local, servicio, tipodoc, contrato, fecharegistro, nombreArchivo, url);
        }

        public List<Novedades> Capa_Negocio_Get_Listado_Operarios()
        {
            try
            {
                VisorNovedades_DAO Objeto_Dato = new VisorNovedades_DAO();
                return Objeto_Dato.Capa_Dato_Get_Operarios();
            }
            catch (Exception e)
            {
                throw e;
            }
        }

        public List<Novedades> Capa_Negocio_Get_ListandoDetallesDocumentos(string id_operario, string fecha_ini, string fecha_fin, string contrato)
        {
            try
            {
                VisorNovedades_DAO Objeto_Dato = new VisorNovedades_DAO();
                return Objeto_Dato.Capa_Dato_Get_ListandoDetallesDocumentos(id_operario, fecha_ini, fecha_fin, contrato);
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
                VisorNovedades_DAO Objeto_Dato = new VisorNovedades_DAO();
                return Objeto_Dato.Capa_Dato_Get_RutaArchivos(id_Archivo);
            }
            catch (Exception e)
            {
                throw e;
            }
        }
    }
}
