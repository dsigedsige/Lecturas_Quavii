using DSIGE.Modelo;
using DSIGE.Negocio;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;

namespace DSIGE.Web.Controllers
{
    public class ResultadosLecturaController : Controller
    {
        //
        // GET: /CambioEstado/

        public ActionResult index_ResultadosLecturas()
        {
            return View();
        }


        public ActionResult index_ResultadosLecturas_Detallado(int servicio, string fecha, string sector)
        {
            ViewBag.servicio_global = servicio;
            ViewBag.fecha_global = fecha;
            ViewBag.sector_global = sector;
            return View();
        }
        

        public ActionResult index_ResultadosLecturas_II()
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
                ResultadoLecturas_BL obj_negocio = new ResultadoLecturas_BL();
                loDatos = obj_negocio.Capa_Negocio_Get_ListaServicios();
                return _Serialize(loDatos, true);
            }
            catch (Exception ex)
            {
                return _Serialize(ex.Message, true);
            }

        }

        [HttpPost]
        public string ListandoOperarios()
        {
            object loDatos;
            try
            {
                ResultadoLecturas_BL obj_negocio = new ResultadoLecturas_BL();
                loDatos = obj_negocio.Capa_Negocio_Get_ListaOperarios();
                return _Serialize(loDatos, true);
            }
            catch (Exception ex)
            {
                return _Serialize(ex.Message, true);
            }

        }

        [HttpPost]
        public string ListandoSectores()
        {
            object loDatos;
            try
            {
                ResultadoLecturas_BL obj_negocio = new ResultadoLecturas_BL();
                loDatos = obj_negocio.Capa_Negocio_Get_ListaSector();
                return _Serialize(loDatos, true);
            }
            catch (Exception ex)
            {
                return _Serialize(ex.Message, true);
            }

        }

 
        [HttpPost]
        public string ListandoResumenLecturas( int Local, string FechaAsignacion, int id_tiposervicio,int id_supervisor, int id_operario_supervisor, string ciclo)
        {
            object loDatos;
            try
            {
                ResultadoLecturas_BL obj_negocio = new ResultadoLecturas_BL();
                loDatos = obj_negocio.Capa_Negocio_Get_ListaResumenLectura(Local, FechaAsignacion, id_tiposervicio, id_supervisor, id_operario_supervisor, ciclo);
                return _Serialize(loDatos, true);
            }
            catch (Exception ex)
            {
                return _Serialize(ex.Message, true);
            }
        }
                

        [HttpPost]
        public string ListandoResumenLecturas_Observacion(int  anio, int mes, string sector)
        {
            object loDatos;
            try
            {
                ResultadoLecturas_BL obj_negocio = new ResultadoLecturas_BL();
                loDatos = obj_negocio.Capa_Negocio_Get_ListaLectura_Observacion(anio, mes, sector);
                return _Serialize(loDatos, true);
            }
            catch (Exception ex)
            {
                return _Serialize(ex.Message, true);
            }
        }
        

        [HttpPost]
        public string ListandoResumenLecturas_Detallado(int Local , string FechaAsignacion, int id_tiposervicio,int id_supervisor, int id_operario_supervisor, string ciclo)
        {
            object loDatos;
            try
            {
                ResultadoLecturas_BL obj_negocio = new ResultadoLecturas_BL();
                loDatos = obj_negocio.Capa_Negocio_Get_ListaLectura_Detallado(Local, FechaAsignacion, id_tiposervicio, id_supervisor, id_operario_supervisor, ciclo);
                return _Serialize(loDatos, true);
            }
            catch (Exception ex)
            {
                return _Serialize(ex.Message, true);
            }
        }
        

        [HttpPost]
        public string VerificacionFotos(string lecturas)
        {
            object loDatos;
            try
            {
                ResultadoLecturas_BL obj_negocio = new ResultadoLecturas_BL();
                loDatos = obj_negocio.Capa_Negocio_VerificacionLecturas(lecturas);
                return _Serialize(loDatos, true);
            }
            catch (Exception ex)
            {
                return _Serialize(ex.Message, true);
            }

        }


        [HttpPost]
        public string ListandoFotoSelfie(int anio, int mes, string sector, int operario)
        {
            object loDatos;
            try
            {
                ResultadoLecturas_BL obj_negocio = new ResultadoLecturas_BL();
                loDatos = obj_negocio.Capa_Negocio_Get_FotoSelfie(anio, mes, sector, operario);
                return _Serialize(loDatos, true);
            }
            catch (Exception ex)
            {
                return _Serialize(ex.Message, true);
            }
        }



        [HttpPost]
        public string ListandoNotasOperario(string fechaAsignacion, int idServicio, int operario)
        {
            object loDatos;
            try
            {
                ResultadoLecturas_BL obj_negocio = new ResultadoLecturas_BL();
                loDatos = obj_negocio.Capa_Negocio_Get_NotasOperario(fechaAsignacion, idServicio, operario);
                return _Serialize(loDatos, true);
            }
            catch (Exception ex)
            {
                return _Serialize(ex.Message, true);
            }
        }

        [HttpPost]
        public string Listando_Observaciones(string fechaAsignacion, int idServicio, int operario ,  int id_supervisor, int id_operario_supervisor)
        {
            object loDatos;
            try
            {
                ResultadoLecturas_BL obj_negocio = new ResultadoLecturas_BL();
                loDatos = obj_negocio.Capa_Negocio_Get_Observaciones(fechaAsignacion, idServicio, operario, id_supervisor, id_operario_supervisor);
                return _Serialize(loDatos, true);
            }
            catch (Exception ex)
            {
                return _Serialize(ex.Message, true);
            }
        }


        [HttpPost]
        public string Guardando_NotasOperario(string fechaAsignacion, int idServicio, int operario, string observacion)
        {
            object loDatos;
            try
            {
                ResultadoLecturas_BL obj_negocio = new ResultadoLecturas_BL();
                loDatos = obj_negocio.Capa_Negocio_Guardar_NotasOperario(fechaAsignacion, idServicio, operario, observacion, ((Sesion)Session["Session_Usuario_Acceso"]).usuario.usu_id);
                return _Serialize(loDatos, true);
            }
            catch (Exception ex)
            {
                return _Serialize(ex.Message, true);
            }
        }


        [HttpPost]
        public string Listando_UltimoSector()
        {
            object loDatos;
            try
            {
                ResultadoLecturas_BL obj_negocio = new ResultadoLecturas_BL();
                loDatos = obj_negocio.Capa_Negocio_Get_ListaUltimoSector();
                return _Serialize(loDatos, true);
            }
            catch (Exception ex)
            {
                return _Serialize(ex.Message, true);
            }
        }



        [HttpPost]
        public string Descargar_efectividad(int Local,  int servicio, string fecha, int id_supervisor, int id_operario_supervisor)
        {
            object loDatos;
            try
            {
                ResultadoLecturas_BL obj_negocio = new ResultadoLecturas_BL();
                loDatos = obj_negocio.Capa_Negocio_Descargar_efectividad(Local, servicio, fecha, ((Sesion)Session["Session_Usuario_Acceso"]).usuario.usu_id, id_supervisor, id_operario_supervisor);
                return _Serialize(loDatos, true);
            }
            catch (Exception ex)
            {
                return _Serialize(ex.Message, true);
            }
        }


        [HttpPost]
        public string Descargar_efectividad_detallado(int Local, int servicio, string fecha)
        {
            object loDatos;
            try
            {
                ResultadoLecturas_BL obj_negocio = new ResultadoLecturas_BL();
                loDatos = obj_negocio.Capa_Negocio_Descargar_efectividad_detallado(Local, servicio, fecha, ((Sesion)Session["Session_Usuario_Acceso"]).usuario.usu_id);
                return _Serialize(loDatos, true);
            }
            catch (Exception ex)
            {
                return _Serialize(ex.Message, true);
            }
        }



    }
}
