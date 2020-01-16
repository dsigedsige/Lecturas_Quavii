using DSIGE.Modelo;
using DSIGE.Negocio;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;

namespace DSIGE.Web.Controllers
{
    public class SeguimientoLecturasController : Controller
    {
        //
        // GET: /Seguimiento_Operarios/

        public ActionResult SeguimientoLecturas()
        {
            return View();
        }


        public ActionResult index_SeguimientoLecturas_II(string fechaAsignacion, int Servicio, int nroDocOperario)
        {
            ViewBag.FechaAsignacion_global = fechaAsignacion;
            ViewBag.Servicio_global = Servicio;
            ViewBag.operario_global = nroDocOperario;
            return View();
        }

        /// <summary>
        /// Autor: rcontreras
        /// Fecha: 30-09-2015
        /// </summary>
        /// <returns></returns>
        [HttpPost]
        public ActionResult JsonSeguimiento_OperariosGPS(string __a, string __b, string __c, string __d, string lista)
        {
            return new ContentResult
            {
                Content = MvcApplication._Serialize(new NUbicacion_Lecturas().NListaLecturasOperariosGPS(
                Convert.ToInt32(__a),
                Convert.ToString(__b),
                __c,
                __d,
                ((Sesion)Session["Session_Usuario_Acceso"]).empresa.emp_id, lista)),
                ContentType = "application/json"
            };
        }




        [HttpPost]
        public ActionResult JsonSeguimiento_lectura_II(int Local, string FechaAsiga, int servicio, int operario, string suministro, string medidor)
        {
            List<UbicacionOperario> list = new List<UbicacionOperario>();
            try
            {
                list = new NUbicacion_Lecturas().NSeguimiento_Lecturas_II(Local, FechaAsiga, servicio, operario, suministro, medidor);
            }
            catch (Exception e)
            {
                return new ContentResult
                {
                    Content = MvcApplication._Serialize(e.Message),
                    ContentType = "application/json"
                };
            }
            return new ContentResult
            {
                Content = MvcApplication._Serialize(list),
                ContentType = "application/json"
            };
        }



        [HttpPost]
        public ActionResult JsonSeguimiento_lectura_III(string FechaAsiga, int servicio, int operario, string suministro, string medidor)
        {
            List<UbicacionOperario> list = new List<UbicacionOperario>();
            try
            {
                list = new NUbicacion_Lecturas().NSeguimiento_Lecturas_III(FechaAsiga, servicio, operario, suministro, medidor);
            }
            catch (Exception e)
            {
                return new ContentResult
                {
                    Content = MvcApplication._Serialize(e.Message),
                    ContentType = "application/json"
                };
            }
            return new ContentResult
            {
                Content = MvcApplication._Serialize(list),
                ContentType = "application/json"
            };
        }


        [HttpPost]
        public ActionResult JsonSeguimiento_lectura_Resumen(string FechaAsiga, int servicio, int operario)
        {
            List<UbicacionOperario> list = new List<UbicacionOperario>();
            try
            {
                list = new NUbicacion_Lecturas().NSeguimiento_Lecturas_Resumen(FechaAsiga, servicio, operario);
            }
            catch (Exception e)
            {
                return new ContentResult
                {
                    Content = MvcApplication._Serialize(e.Message),
                    ContentType = "application/json"
                };
            }
            return new ContentResult
            {
                Content = MvcApplication._Serialize(list),
                ContentType = "application/json"
            };
        }


 

    }
}
