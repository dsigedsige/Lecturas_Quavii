using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DSIGE.Modelo;
using DSIGE.Dato;

namespace DSIGE.Negocio
{
    public class ImportarArchivoPlano_BL
    {
        public string Capa_Negocio_GuardarArchivos(List<ImportarArchivoPlano> obj_list, int id_user, int empresa)
        {
            ImportarArchivoPlano_DAO Objeto_Dato = new ImportarArchivoPlano_DAO();
            return Objeto_Dato.Capa_Dato_GuardarArchivos(obj_list, id_user, empresa);
        }
        public string Capa_Negocio_GuardarArchivos_new(List<ImportarArchivoPlano_new> obj_list, int id_user, int empresa)
        {
            try
            {
                ImportarArchivoPlano_DAO Objeto_Dato = new ImportarArchivoPlano_DAO();
                return Objeto_Dato.Capa_Dato_GuardarArchivos_new(obj_list, id_user, empresa);
            }
            catch (Exception)
            {
                throw;
            }
        }
               
        public List<ImportarArchivoPlano> Capa_Negocio_Get_ListaArchivosAlmacenados(string fechaAsignada, int TipoServicio)
        {
            try
            {
                ImportarArchivoPlano_DAO Objeto_Dato = new ImportarArchivoPlano_DAO();
                return Objeto_Dato.Capa_Dato_Get_ListaArchivosAlmacenados(fechaAsignada, TipoServicio);
            }
            catch (Exception e)
            {
                throw e;
            }

        }

    }
}
