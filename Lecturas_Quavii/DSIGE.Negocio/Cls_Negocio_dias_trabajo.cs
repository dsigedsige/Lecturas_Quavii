 
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DSIGE.Dato;
using DSIGE.Modelo;

namespace DSIGE.Negocio
{
    public class Cls_Negocio_dias_trabajo
    {
        //Listado

        public List<Cls_Entidad_dias_trabajo> Capa_Negocio_Get_Listado_dias_trabajo()
        {
            try
            {
                Cls_Dato_dias_trabajo Objeto_Dato = new Cls_Dato_dias_trabajo();
                return Objeto_Dato.Capa_Dato_Get_Listado_dias_trabajo();
            }
            catch (Exception e)
            {
                throw e;
            }

        }

        public List<Cls_Entidad_dias_trabajo> Capa_Negocio_Get_Listado_operario()
        {
            try
            {
                Cls_Dato_dias_trabajo Objeto_Dato = new Cls_Dato_dias_trabajo();
                return Objeto_Dato.Capa_Dato_Get_Listado_operario();
            }
            catch (Exception e)
            {
                throw e;
            }

        }








        //Guardando 
        public bool Capa_Negocio_Guardar_Informacion(Cls_Entidad_dias_trabajo Entidad)
        {
            try
            {
                Cls_Dato_dias_trabajo Objeto_Dato = new Cls_Dato_dias_trabajo();
                return Objeto_Dato.Capa_Dato_Guardar_Informacion(Entidad);
            }
            catch (Exception)
            {
                
                throw;
            }
        }

        // modificando

        public bool Capa_Negocio_Modificar_Informacion(Cls_Entidad_dias_trabajo Entidad)
        {
            try
            {
                Cls_Dato_dias_trabajo Objeto_Dato = new Cls_Dato_dias_trabajo();
                return Objeto_Dato.Capa_Dato_Modificar_Informacion (Entidad);
            }
            catch (Exception)
            {

                throw;
            }
        }


        // anulando

        public bool Capa_Negocio_Anular_Informacion(Cls_Entidad_dias_trabajo Entidad)
        {
            try
            {
                Cls_Dato_dias_trabajo Objeto_Dato = new Cls_Dato_dias_trabajo();
                return Objeto_Dato.Capa_Dato_Anular_Informacion (Entidad);
            }
            catch (Exception)
            {

                throw;
            }
        }

        //auditoria


        public List<Cls_Entidad_dias_trabajo> Capa_Negocio_Get_Listado_Auditoria(int id_dia)
        {
            try
            {
                Cls_Dato_dias_trabajo Objeto_Dato = new Cls_Dato_dias_trabajo();
                return Objeto_Dato.Capa_Dato_Get_Listado_Auditoria(id_dia);
            }
            catch (Exception e)
            {

                throw e;
            }

        }


    }
}
