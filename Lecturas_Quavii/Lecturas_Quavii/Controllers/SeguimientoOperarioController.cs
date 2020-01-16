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
    public class SeguimientoOperarioController : Controller
    {
        //
             public ActionResult SeguimientoOperario()
        {
            return View();
        }

        public ActionResult SeguimientoOperario_reparto()
        {
            return View();
        }

        public ActionResult efectividadOperarios_index()
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


        public ActionResult index_SeguimientoOperario_II(string fechaAsignacion, int Servicio, int nroDocOperario)
        {
            ViewBag.FechaAsignacion_global = fechaAsignacion;
            ViewBag.Servicio_global = Servicio;
            ViewBag.operario_global = nroDocOperario;
            return View();
        }

        [HttpPost]
        public ActionResult JsonSeguimiento_Operarios_GPS2(string __a, string __b, string lista)
        {
            List<UbicacionOperario_GPS> list = new List<UbicacionOperario_GPS>();
            try
            {

                list = new NSeguimiento_Operario().NSeguimientoOperario_GPS(Convert.ToInt32(__a), __b, lista);


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
        public ActionResult JsonSeguimiento_Operarios_GPS_II( int Local, string FechaAsiga, int servicio, int operario, string suministro, string medidor)
        {
            List<UbicacionOperario_GPS> list = new List<UbicacionOperario_GPS>();
            try
            {
                list = new NSeguimiento_Operario().NSeguimientoOperario_GPS_II(Local,FechaAsiga, servicio, operario, suministro, medidor);
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
        public ActionResult JsonSeguimiento_Operarios_GPS_Reparto(string FechaAsiga, int servicio, int operario, string suministro, string medidor)
        {
            List<UbicacionOperario_GPS> list = new List<UbicacionOperario_GPS>();
            try
            {
                list = new NSeguimiento_Operario().NSeguimientoOperario_GPS_reparto(FechaAsiga, servicio, operario, suministro, medidor);
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
        public ActionResult JsonSeguimiento_Operarios_Resumen(string FechaAsiga, int servicio, int  operario)
        {
            List<UbicacionOperario_GPS> list = new List<UbicacionOperario_GPS>();
            try
            {
                list = new NSeguimiento_Operario().NSeguimientoOperario_GPS_Resumen(FechaAsiga, servicio,  operario);
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
        public ActionResult descargar_informacion_reparto_excel(string FechaAsiga, int servicio, int operario, string suministro, string medidor)
        {
            object list;
            try
            {
                list = new NSeguimiento_Operario().Ndescargar_informacion_reparto_excel(FechaAsiga, servicio, operario, suministro, medidor, ((Sesion)Session["Session_Usuario_Acceso"]).usuario.usu_id);
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
        public string get_Operario_Gps(string FechaAsiga, int servicio, int operario)
        {
            object loDatos;
            try
            {
                loDatos = new NSeguimiento_Operario().Capa_Negocio_get_Operario_Gps(FechaAsiga, servicio, operario);
                return _Serialize(loDatos, true);
            }
            catch (Exception ex)
            {
                return _Serialize(ex.Message, true);
            }
        }


        [HttpPost]
        public string get_Suministros_Mapa(string FechaAsiga, int servicio, int operario)
        {
            object loDatos;
            try
            {
                loDatos = new NSeguimiento_Operario().Capa_Negocio_get_Suministros_Mapa(FechaAsiga, servicio, operario);
                return _Serialize(loDatos, true);
            }
            catch (Exception ex)
            {
                return _Serialize(ex.Message, true);
            }
        }


        [HttpPost]
        public string get_Suministros_Trabajados(string FechaAsiga, int servicio, int operario)
        {
            object loDatos;
            try
            {
                loDatos = new NSeguimiento_Operario().Capa_Negocio_get_Suministros_Trabajados(FechaAsiga, servicio, operario);
                return _Serialize(loDatos, true);
            }
            catch (Exception ex)
            {
                return _Serialize(ex.Message, true);
            }
        }



    }
}
