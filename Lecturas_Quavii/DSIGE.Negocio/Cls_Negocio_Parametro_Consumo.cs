using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DSIGE.Dato;
using DSIGE.Modelo;

namespace DSIGE.Negocio
{
    public class Cls_Negocio_Parametro_Consumo
    {
        //Listado

        public List<Cls_Entidad_Parametro_Consumo> Capa_Negocio_Get_Listado_Parametros()
        {
            try
            {
                Cls_Dato_Parametro_Consumo Objeto_Dato = new Cls_Dato_Parametro_Consumo();
                return Objeto_Dato.Capa_Dato_Get_Listado_Parametros ();
            }
            catch (Exception e)
            {
                throw e;
            }

        }
        //Guardando 
        public bool Capa_Negocio_Guardar_Informacion(Cls_Entidad_Parametro_Consumo Entidad)
        {
            try
            {
                Cls_Dato_Parametro_Consumo Objeto_Dato = new Cls_Dato_Parametro_Consumo();
                return Objeto_Dato.Capa_Dato_Guardar_Informacion(Entidad);
            }
            catch (Exception)
            {
                
                throw;
            }
        }

        // modificando

        public bool Capa_Negocio_Modificar_Informacion(Cls_Entidad_Parametro_Consumo Entidad)
        {
            try
            {
                Cls_Dato_Parametro_Consumo Objeto_Dato = new Cls_Dato_Parametro_Consumo();
                return Objeto_Dato.Capa_Dato_Modificar_Informacion (Entidad);
            }
            catch (Exception)
            {

                throw;
            }
        }


        // anulando

        public bool Capa_Negocio_Anular_Informacion(Cls_Entidad_Parametro_Consumo Entidad)
        {
            try
            {
                Cls_Dato_Parametro_Consumo Objeto_Dato = new Cls_Dato_Parametro_Consumo();
                return Objeto_Dato.Capa_Dato_Anular_Informacion (Entidad);
            }
            catch (Exception)
            {

                throw;
            }
        }

        //auditoria


        public List<Cls_Entidad_Parametro_Consumo> Capa_Negocio_Get_Listado_Parametros_Auditoria(int id_parametro)
        {
            try
            {
                Cls_Dato_Parametro_Consumo Objeto_Dato = new Cls_Dato_Parametro_Consumo();
                return Objeto_Dato.Capa_Dato_Get_Listado_Parametros_Auditoria(id_parametro);
            }
            catch (Exception e)
            {

                throw e;
            }

        }


    }
}
