using DSIGE.Modelo;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using Excel = OfficeOpenXml;
using Style = OfficeOpenXml.Style;
 

namespace DSIGE.Web.Controllers
{
    public class DistribuirLecturasController : Controller
    {
        //
        // GET: /DistribucionTrabajos/

        public ActionResult Distribuirlecturas_index()
        {
            return View();
        }

        public static string _Serialize(object value, bool ignore = false)
        {
            var SerializerSettings = new JsonSerializerSettings()
            {
                MaxDepth = Int32.MaxValue,
                NullValueHandling = (ignore == true ? NullValueHandling.Ignore : NullValueHandling.Include)
            };
            return JsonConvert.SerializeObject(value, Formatting.Indented, SerializerSettings);
        }
        
        [HttpPost]
        public string ListandoServicios()
        {
            object loDatos;
            try
            {
                DistribuirLecturas_BL obj_negocio = new DistribuirLecturas_BL();
                loDatos = obj_negocio.Capa_Negocio_Get_ListaServicios();
                return _Serialize(loDatos, true);
            }
            catch (Exception ex)
            {
                return _Serialize(ex.Message, true);
            }

        }

        [HttpPost]
        public string ListandoLocal()
        {
            object loDatos;
            try
            {
                DistribuirLecturas_BL obj_negocio = new DistribuirLecturas_BL();
                loDatos = obj_negocio.Capa_Negocio_Get_ListaLocales();
                return _Serialize(loDatos, true);
            }
            catch (Exception ex)
            {
                return _Serialize(ex.Message, true);
            }
        }

        [HttpPost]
        public string ListandoInformacionLecturas(int local, string fechaAsigna, int servicio, string opcion, int id_supervisor, int id_operario_supervisor, string tipoCliente)
        {
            object loDatos;
            try
            {
                DistribuirLecturas_BL obj_negocio = new DistribuirLecturas_BL();
                loDatos = obj_negocio.Capa_Negocio_Get_ListaInformacionLecturas(local, fechaAsigna, servicio, opcion, id_supervisor, id_operario_supervisor, tipoCliente);
                return _Serialize(loDatos, true);
            }
            catch (Exception ex)
            {
                return _Serialize(ex.Message, true);
            }
        }


        [HttpPost]
        public string ListandoInformacionLecturasDetalle(int id_local, string fechaAsigna, int id_servicio, string unidad_Lectura, int id_operario, string opcion)
        {         
            object loDatos;
            try
            {
                DistribuirLecturas_BL obj_negocio = new DistribuirLecturas_BL();
                loDatos = obj_negocio.Capa_Negocio_Get_ListaInformacionLecturasDetalle(id_local, fechaAsigna, id_servicio, unidad_Lectura, id_operario, opcion);
                return _Serialize(loDatos, true);
            }
            catch (Exception ex)
            {
                return _Serialize(ex.Message, true);
            }
        }

        [HttpPost]
        public string ListandoInformacionLecturasDetalle_general(int id_local, string fechaAsigna, int id_servicio, string unidad_Lectura, int id_operario, string opcion)
        {
            object loDatos;
            try
            {
                DistribuirLecturas_BL obj_negocio = new DistribuirLecturas_BL();
                loDatos = obj_negocio.Capa_Negocio_Get_ListaInformacionLecturasDetalle_general(id_local, fechaAsigna, id_servicio, unidad_Lectura, id_operario, opcion);
                return _Serialize(loDatos, true);
            }
            catch (Exception ex)
            {
                return _Serialize(ex.Message, true);
            }
        }


        [HttpPost]
        public string GenerandoReasignacionTrabajos(List<string> ListaTrabajos, string FechaMovil, string servicio)
        {
            object loDatos;
            try
            {
                DistribuirLecturas_BL obj_negocio = new DistribuirLecturas_BL();
                loDatos = obj_negocio.Capa_Negocio_GenerandoReasignacionTrabajos(ListaTrabajos, ((Sesion)Session["Session_Usuario_Acceso"]).usuario.usu_id, FechaMovil, servicio);
                return _Serialize(loDatos, true);
            }
            catch (Exception ex)
            {
                return _Serialize(ex.Message, true);
            }
        }

        [HttpPost]
        public string GenerandoReasignacionTrabajosDetalle(List<string> ListaTrabajos, string FechaMovil, string servicio)
        {
            object loDatos;
            try
            {
                DistribuirLecturas_BL obj_negocio = new DistribuirLecturas_BL();
                loDatos = obj_negocio.Capa_Negocio_GenerandoReasignacionTrabajosDetalle(ListaTrabajos, ((Sesion)Session["Session_Usuario_Acceso"]).usuario.usu_id, FechaMovil, servicio);
                return _Serialize(loDatos, true);
            }
            catch (Exception ex)
            {
                return _Serialize(ex.Message, true);
            }
        }


        [HttpPost]
        public string validarExistenciaOperario(int id_operario)
        {
            object loDatos;
            try
            {
                DistribuirLecturas_BL obj_negocio = new DistribuirLecturas_BL();
                loDatos = obj_negocio.Capa_Negocio_get_ValidarOperario(id_operario);
                return _Serialize(loDatos, true);
            }
            catch (Exception ex)
            {
                return _Serialize(ex.Message, true);
            }
        }


        [HttpPost]
        public string validarOperario(int CodigoOperario)
        {
            object loDatos;
            try
            {
                DistribuirLecturas_BL obj_negocio = new DistribuirLecturas_BL();
                loDatos = obj_negocio.Capa_Negocio_get_ValidarOperario(CodigoOperario);
                return _Serialize(loDatos, true);
            }
            catch (Exception ex)
            {
                return _Serialize(ex.Message, true);
            }
        }


        [HttpPost]
        public string DistribuirOrdenes(int id_local, string fechaAsignacion, int id_servicio, string unidad_lectura, int id_operario)
        {
            object loDatos;
            try
            {
                DistribuirLecturas_BL obj_negocio = new DistribuirLecturas_BL();
                loDatos = obj_negocio.Capa_Negocio_Get_DistribuirOrdenes(id_local, fechaAsignacion, id_servicio, unidad_lectura, id_operario);
                return _Serialize(loDatos, true);
            }
            catch (Exception ex)
            {
                return _Serialize(ex.Message, true);
            }
        }

        [HttpPost]
        public string Generando_EnvioMovil_Distribucion_Reasignacion(List<Ordenes_E> ListaOrdenes, int id_local, string FechaAsigna, int servicio,string FechaMovil , string opcion)
        {
            object loDatos=null;
            try
            {
                DistribuirLecturas_BL obj_negocio = new DistribuirLecturas_BL();

                if (opcion == "D")
                {
                    loDatos = obj_negocio.Capa_Negocio_Set_Distribucion_EnviarMovil(ListaOrdenes, id_local, FechaAsigna, servicio, FechaMovil, opcion, ((Sesion)Session["Session_Usuario_Acceso"]).usuario.usu_id);  
                }
                else if (opcion == "R")
                {
                    loDatos = obj_negocio.Capa_Negocio_Set_Reasignacion_EnviarMovil(ListaOrdenes, id_local, FechaAsigna, servicio, FechaMovil, opcion, ((Sesion)Session["Session_Usuario_Acceso"]).usuario.usu_id);
                }

                return _Serialize(loDatos, true);
            }
            catch (Exception ex)
            {
                return _Serialize(ex.Message, true);
            }
        }


        [HttpPost]
        public string Generando_EnvioMovil_Distribucion_Reasignacion_Detallado(List<OrdenesDetalle_E> ListaOrdenes, int id_local, string FechaAsigna, int servicio, string FechaMovil, string opcion, string Flag_detallado)
        {
            object loDatos = null;
            try
            {
                DistribuirLecturas_BL obj_negocio = new DistribuirLecturas_BL();

                if (opcion == "D")
                {
                    loDatos = obj_negocio.Capa_Negocio_Set_Distribucion_EnviarMovil_detallado(ListaOrdenes, id_local, FechaAsigna, servicio, FechaMovil, opcion, ((Sesion)Session["Session_Usuario_Acceso"]).usuario.usu_id, Flag_detallado);
                }
                else if (opcion == "R")
                {
                    loDatos = obj_negocio.Capa_Negocio_Set_Reasignacion_EnviarMovil_detallado(ListaOrdenes, id_local, FechaAsigna, servicio, FechaMovil, opcion, ((Sesion)Session["Session_Usuario_Acceso"]).usuario.usu_id, Flag_detallado);
                }

                return _Serialize(loDatos, true);
            }
            catch (Exception ex)
            {
                return _Serialize(ex.Message, true);
            }
        }

        [HttpPost]
        public string ProcesandoConsumos(string fechaAsigna, int id_servicio)
        {
            object loDatos;
            try
            {
                DistribuirLecturas_BL obj_negocio = new DistribuirLecturas_BL();
                loDatos = obj_negocio.Capa_Negocio_get_ProcesoConsumo(fechaAsigna, ((Sesion)Session["Session_Usuario_Acceso"]).usuario.usu_id, id_servicio);
                return _Serialize(loDatos, true);
            }
            catch (Exception ex)
            {
                return _Serialize(ex.Message, true);
            }
        }

        [HttpPost]
        public string ProcesandoConsumos_new(string fechaAsigna, int id_servicio)
        {
            object loDatos;
            try
            {
                DistribuirLecturas_BL obj_negocio = new DistribuirLecturas_BL();
                loDatos = obj_negocio.Capa_Negocio_get_ProcesoConsumo_new(fechaAsigna, ((Sesion)Session["Session_Usuario_Acceso"]).usuario.usu_id, id_servicio);
                return _Serialize(loDatos, true);
            }
            catch (Exception ex)
            {
                return _Serialize(ex.Message, true);
            }
        }

        [HttpPost]
        public string Proceso_lectura_pendiente(string fechaAsigna, string fecha_relectura)
        {
            object loDatos;
            try
            {
                DistribuirLecturas_BL obj_negocio = new DistribuirLecturas_BL();
                loDatos = obj_negocio.Capa_Negocio_Proceso_lectura_pendiente(fechaAsigna, ((Sesion)Session["Session_Usuario_Acceso"]).usuario.usu_id, fecha_relectura);
                return _Serialize(loDatos, true);
            }
            catch (Exception ex)
            {
                return _Serialize(ex.Message, true);
            }
        }

        [HttpPost]
        public string Generando_Enviar_Correo(List<string> ListaOrdenes, int id_local, string FechaAsigna, int servicio)
        {
            object loDatos=null;
            try
            {
                  DistribuirLecturas_BL obj_negocio = new DistribuirLecturas_BL();
                  loDatos = obj_negocio.Capa_Negocio_DetalleLecturas_Correo(ListaOrdenes, ((Sesion)Session["Session_Usuario_Acceso"]).usuario.usu_id, id_local, FechaAsigna, servicio);
                 return _Serialize(loDatos, true);
            }
            catch (Exception ex)
            {
                return _Serialize("0|" + ex.Message, true);
            }
        }


        [HttpPost]
        public string Inserta_Excel_Lecturas(HttpPostedFileBase file, int idlocal, string idfechaAsignacion, int idServicio, string fecha_lectura)
        {
            try
            {
                object loDatos = null;
                DistribuirLecturas_BL obj_negocio = new DistribuirLecturas_BL();
                loDatos = obj_negocio.Capa_Negocio_Inserta_Lecturas(file, idlocal, idfechaAsignacion, idServicio, ((Sesion)Session["Session_Usuario_Acceso"]).usuario.usu_id, fecha_lectura);
     
                return _Serialize(loDatos, true);
            }
            catch (Exception ex)
            {
                return _Serialize(ex.Message, true);
            }
        }





    }
}
