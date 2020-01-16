using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DSIGE.Modelo;
using DSIGE.Dato;
using System.Web;

namespace DSIGE.Modelo
{
    public class DistribuirLecturas_BL
    {

       public List<DistribuirLecturas_E> Capa_Negocio_Get_ListaLocales()
       {
           try
           {
               Cls_Dato_DistrilbuirLecturas Objeto_Dato = new Cls_Dato_DistrilbuirLecturas();
               return Objeto_Dato.Capa_Dato_Get_ListaLocales ();
           }
           catch (Exception e)
           {
               throw e;
           }

       }

       public List<DistribuirLecturas_E> Capa_Negocio_Get_ListaServicios()
       {
           try
           {
               Cls_Dato_DistrilbuirLecturas Objeto_Dato = new Cls_Dato_DistrilbuirLecturas();
               return Objeto_Dato.Capa_Dato_Get_ListaServicio();
           }
           catch (Exception e)
           {
               throw e;
           }
       }


       public List<DistribuirLecturas_E> Capa_Negocio_Get_ListaInformacionLecturas(int local, string fechaAsigna, int servicio, string opcion,  int id_supervisor, int id_operario_supervisor, string tipoCliente)
        {
           try
           {
               Cls_Dato_DistrilbuirLecturas Objeto_Dato = new Cls_Dato_DistrilbuirLecturas();
               return Objeto_Dato.Capa_Dato_Get_ListarInformacionLecturas(local, fechaAsigna, servicio, opcion, id_supervisor, id_operario_supervisor, tipoCliente);
           }
           catch (Exception e)
           {
               throw e;
           }

       }

       public List<DistribuirLecturas_E> Capa_Negocio_Get_ListaInformacionLecturasDetalle(int id_local, string fechaAsigna, int id_servicio, string unidad_Lectura, int id_operario, string opcion)
       {
           try
           {
               Cls_Dato_DistrilbuirLecturas Objeto_Dato = new Cls_Dato_DistrilbuirLecturas();
               return Objeto_Dato.Capa_Dato_Get_ListarInformacionLecturasDetalle(id_local, fechaAsigna, id_servicio, unidad_Lectura, id_operario, opcion);
           }
           catch (Exception e)
           {
               throw e;
           }
       }

        public List<DistribuirLecturas_E> Capa_Negocio_Get_ListaInformacionLecturasDetalle_general(int id_local, string fechaAsigna, int id_servicio, string unidad_Lectura, int id_operario, string opcion)
        {
            try
            {
                Cls_Dato_DistrilbuirLecturas Objeto_Dato = new Cls_Dato_DistrilbuirLecturas();
                return Objeto_Dato.Capa_Dato_Get_ListarInformacionLecturasDetalle_general(id_local, fechaAsigna, id_servicio, unidad_Lectura, id_operario, opcion);
            }
            catch (Exception e)
            {
                throw e;
            }
        }


        public List<DistribuirLecturas_E> Capa_Negocio_Get_ListaInformacionLecturasDetalleDistribucion(int local, string fechaAsigna, int servicio, int id_operario, int zona, int sector)
       {
           try
           {
               Cls_Dato_DistrilbuirLecturas Objeto_Dato = new Cls_Dato_DistrilbuirLecturas();
               return Objeto_Dato.Capa_Dato_Get_ListarInformacionLecturasDetalleDistribucion(local, fechaAsigna, servicio, id_operario, zona, sector);
           }
           catch (Exception e)
           {
               throw e;
           }

       }

       public string Capa_Negocio_GenerandoReasignacionTrabajos(List<string> ListaTrabajos, int id_usuario, string FechaMovil, string servicio)
       {
           try
           {
               Cls_Dato_DistrilbuirLecturas Objeto_Dato = new Cls_Dato_DistrilbuirLecturas();
               return Objeto_Dato.Capa_Dato_Set_ReasignacionTrabajos(ListaTrabajos, id_usuario, FechaMovil, servicio);
           }
           catch (Exception e)
           {
               throw e;
           }

       }

       public string Capa_Negocio_GenerandoReasignacionTrabajosDetalle(List<string> ListaTrabajos, int id_usuario, string FechaMovil, string servicio)
       {
           try
           {
               Cls_Dato_DistrilbuirLecturas Objeto_Dato = new Cls_Dato_DistrilbuirLecturas();
               return Objeto_Dato.Capa_Dato_Set_ReasignacionTrabajosDetalle(ListaTrabajos, id_usuario, FechaMovil, servicio);
           }
           catch (Exception e)
           {
               throw e;
           }

       }

       public List<DistribuirLecturas_E> Capa_Negocio_get_ValidarOperario(int CodigoOperario)
       {
           try
           {
               Cls_Dato_DistrilbuirLecturas Objeto_Dato = new Cls_Dato_DistrilbuirLecturas();
               return Objeto_Dato.Capa_Dato_Get_ValidarOperario(CodigoOperario);
           }
           catch (Exception e)
           {
               throw e;
           }

       }


       public string Capa_Negocio_Get_DistribuirOrdenes(int id_local, string fechaAsignacion, int id_servicio, string unidad_lectura, int id_operario)
       {
           try
           {
               Cls_Dato_DistrilbuirLecturas Objeto_Dato = new Cls_Dato_DistrilbuirLecturas();
               return Objeto_Dato.Capa_Dato_Get_DistribuirOrdenes(id_local, fechaAsignacion, id_servicio, unidad_lectura, id_operario);
           }
           catch (Exception e)
           {
               throw e;
           }

       }


       public string Capa_Negocio_Set_Distribucion_EnviarMovil(List<Ordenes_E> ListaOrdenes, int id_local, string FechaAsigna, int servicio, string FechaMovil, string opcion, int usuario)
       {
           try
           {
               Cls_Dato_DistrilbuirLecturas Objeto_Dato = new Cls_Dato_DistrilbuirLecturas();
               return Objeto_Dato.Capa_Dato_Set_Distribucion_EnviarMovil(ListaOrdenes, id_local, FechaAsigna, servicio,FechaMovil, opcion, usuario);
           }
           catch (Exception e)
           {
               throw e;
           }

       }


       public string Capa_Negocio_Set_Reasignacion_EnviarMovil(List<Ordenes_E> ListaOrdenes, int id_local, string FechaAsigna, int servicio, string FechaMovil, string opcion, int usuario)
       {
           try
           {
               Cls_Dato_DistrilbuirLecturas Objeto_Dato = new Cls_Dato_DistrilbuirLecturas();
               return Objeto_Dato.Capa_Dato_Set_Reasignacion_EnviarMovil(ListaOrdenes, id_local, FechaAsigna, servicio, FechaMovil, opcion, usuario);
           }
           catch (Exception e)
           {
               throw e;
           }

       }



       public string Capa_Negocio_Set_Distribucion_EnviarMovil_detallado(List<OrdenesDetalle_E> ListaOrdenes, int id_local, string FechaAsigna, int servicio, string FechaMovil, string opcion, int usuario, string Flag_detallado)
       {
           try
           {
               Cls_Dato_DistrilbuirLecturas Objeto_Dato = new Cls_Dato_DistrilbuirLecturas();
               return Objeto_Dato.Capa_Dato_Set_Distribucion_EnviarMovil_detallado(ListaOrdenes, id_local, FechaAsigna, servicio, FechaMovil, opcion, usuario, Flag_detallado);
           }
           catch (Exception e)
           {
               throw e;
           }

       }


       public string Capa_Negocio_Set_Reasignacion_EnviarMovil_detallado(List<OrdenesDetalle_E> ListaOrdenes, int id_local, string FechaAsigna, int servicio, string FechaMovil, string opcion, int usuario, string Flag_detallado)
       {
           try
           {
               Cls_Dato_DistrilbuirLecturas Objeto_Dato = new Cls_Dato_DistrilbuirLecturas();
               return Objeto_Dato.Capa_Dato_Set_Reasignacion_EnviarMovil_detallado(ListaOrdenes, id_local, FechaAsigna, servicio, FechaMovil, opcion, usuario, Flag_detallado);
           }
           catch (Exception e)
           {
               throw e;
           }

       }


       public string Capa_Negocio_get_ProcesoConsumo(string FechaEnvio, int id_usuario, int id_servicio)
       {
           try
           {
               Cls_Dato_DistrilbuirLecturas Objeto_Dato = new Cls_Dato_DistrilbuirLecturas();
               return Objeto_Dato.Capa_Dato_Set_ProcesarConsumo(FechaEnvio, id_usuario, id_servicio);
           }
           catch (Exception e)
           {
               throw e;
           }
       }


        public string Capa_Negocio_get_ProcesoConsumo_new(string FechaEnvio, int id_usuario, int id_servicio)
        {
            try
            {
                Cls_Dato_DistrilbuirLecturas Objeto_Dato = new Cls_Dato_DistrilbuirLecturas();
                return Objeto_Dato.Capa_Dato_Set_ProcesarConsumo_new(FechaEnvio, id_usuario, id_servicio);
            }
            catch (Exception e)
            {
                throw e;
            }
        }


        public string Capa_Negocio_Proceso_lectura_pendiente(string FechaEnvio, int id_usuario, string fecha_relectura)
        {
            try
            {
                Cls_Dato_DistrilbuirLecturas Objeto_Dato = new Cls_Dato_DistrilbuirLecturas();
                return Objeto_Dato.Capa_Dato_Proceso_lectura_pendiente(FechaEnvio, id_usuario, fecha_relectura);
            }
            catch (Exception e)
            {
                throw e;
            }
        }


        public string Capa_Negocio_DetalleLecturas_Correo(List<string> ListaTrabajos, int id_usuario, int id_local, string FechaAsigna,int  servicio)
       {
           try
           {
               Cls_Dato_DistrilbuirLecturas Objeto_Dato = new Cls_Dato_DistrilbuirLecturas();
               return Objeto_Dato.Capa_Dato_DetalleLecturas_Correo(ListaTrabajos, id_usuario, id_local, FechaAsigna, servicio);
           }
           catch (Exception e)
           {
               throw e;
           }

       }

       public string Capa_Negocio_Inserta_Lecturas(HttpPostedFileBase file, int idlocal, string idfechaAsignacion, int idServicio, int idusuario, string fecha_lectura)
       {
           Cls_Dato_DistrilbuirLecturas Objeto_Dato = new Cls_Dato_DistrilbuirLecturas();
           return Objeto_Dato.Capa_Dato_GuardarArchivo(file, idlocal, idfechaAsignacion, idServicio, idusuario , fecha_lectura);
       }


    }
}
