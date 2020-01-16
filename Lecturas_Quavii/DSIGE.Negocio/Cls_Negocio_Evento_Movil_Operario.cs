using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DSIGE.Modelo;
using DSIGE.Dato;

namespace DSIGE.Negocio
{
    public class Cls_Negocio_Evento_Movil_Operario
    {

        //Listado

        public List<Cls_Entidad_Evento_Movil_Operario> Capa_Negocio_Locales()
        {
            try
            {
                Cls_Dato_Evento_Movil_Operario Objeto_Dato = new Cls_Dato_Evento_Movil_Operario();
                return Objeto_Dato.Capa_Dato_Locales();
            }
            catch (Exception e)
            {
                throw e;
            }
        }

        public List<Cls_Entidad_Evento_Movil_Operario> Capa_Negocio_Operarios(int id_local)
        {
            try
            {
                Cls_Dato_Evento_Movil_Operario Objeto_Dato = new Cls_Dato_Evento_Movil_Operario();
                return Objeto_Dato.Capa_Dato_Operarios(id_local);
            }
            catch (Exception e)
            {
                throw e;
            }
        }

        public List<Cls_Entidad_Evento_Movil_Operario> Capa_Negocio_Listando_Datos_Cabecera(string id_local, string id_operario, string fecha_ini, string fecha_fin, string lista)
        {
            try
            {
                Cls_Dato_Evento_Movil_Operario Objeto_Dato = new Cls_Dato_Evento_Movil_Operario();
                return Objeto_Dato.Capa_Dato_Listando_Datos_Cabecera(id_local, id_operario, fecha_ini, fecha_fin,lista);
            }
            catch (Exception e)
            {
                throw e;
            }
        }

        public List<Cls_Entidad_Evento_Movil_Operario> Capa_Negocio_Datos_Operario_Detallado(int id_operario, string fecha)
        {
            try
            {
              Cls_Dato_Evento_Movil_Operario objeto_dato= new Cls_Dato_Evento_Movil_Operario();
              return objeto_dato.Capa_Dato_Datos_Operario_Detallado(id_operario, fecha);
            }
            catch (Exception e )
            {
                
                throw e;
            }
        
        }


        //   Eventos movil Operario seguimiento




        public List<Cls_Entidad_Evento_Movil_Operario> Capa_Negocio_Listando_Datos_Cabecera_Seguimiento_Operario(string id_local, int id_operario, string fecha_ini, string fecha_fin, string lista)
        {
            try
            {
                Cls_Dato_Evento_Movil_Operario Objeto_Dato = new Cls_Dato_Evento_Movil_Operario();
                return Objeto_Dato.Capa_Dato_Listando_Datos_Cabecera_Seguimiento_Operario(id_local, id_operario, fecha_ini, fecha_fin,lista);
            }
            catch (Exception e)
            {
                throw e;
            }
        }





        //   fin de Eventos movil Operario seguimiento


    }
}
