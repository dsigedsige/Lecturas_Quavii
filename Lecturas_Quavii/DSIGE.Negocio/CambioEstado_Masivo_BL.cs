using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DSIGE.Modelo;
using DSIGE.Dato;

namespace DSIGE.Modelo
{
    public class CambioEstado_Masivo_BL
    {
        public string Capa_Negocio_get_CambioEstado_Masivo(int local, int servicio, int operario,  string fecha_Asigna, int usuario)
       {
           try
           {
               Cls_Dato_CambioEstado Objeto_Dato = new Cls_Dato_CambioEstado();
               return Objeto_Dato.Capa_Dato_Set_CambioEstado_Masivo(local, servicio, operario, fecha_Asigna, usuario);
           }
           catch (Exception e)
           {
               throw e;
           }

       }


        public string Capa_Negocio_GenerarAnular_Masivo(int servicio, string fecha_Asigna, int estado, int operario, int usuario)
        {
            try
            {
                Cls_Dato_CambioEstado Objeto_Dato = new Cls_Dato_CambioEstado();
                return Objeto_Dato.Capa_Dato_GenerarAnular_Masivo(servicio, fecha_Asigna, estado, operario, usuario);
            }
            catch (Exception e)
            {
                throw e;
            }

        }

        public List<Cambio_Estado_Masivo_E> Capa_Negocio_Get_ListaServicios()
        {
            try
            {
                Cls_Dato_CambioEstado Objeto_Dato = new Cls_Dato_CambioEstado();
                return Objeto_Dato.Capa_Dato_Get_ListaServicio();
            }
            catch (Exception e)
            {
                throw e;
            }
        }


        public List<Cambio_Estado_Masivo_E> Capa_Negocio_Get_ListaServicios_usuario(int id_usuario)
        {
            try
            {
                Cls_Dato_CambioEstado Objeto_Dato = new Cls_Dato_CambioEstado();
                return Objeto_Dato.Capa_Dato_Get_ListaServicio_usuario(id_usuario);
            }
            catch (Exception e)
            {
                throw e;
            }
        }

        public List<Cambio_Estado_Masivo_E> Capa_Negocio_Get_ListaLocales()
        {
            try
            {
                Cls_Dato_CambioEstado Objeto_Dato = new Cls_Dato_CambioEstado();
                return Objeto_Dato.Capa_Dato_Get_ListaLocales();
            }
            catch (Exception e)
            {
                throw e;
            }

        }


        
    }
}
