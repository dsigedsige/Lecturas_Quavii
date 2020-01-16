using DSIGE.Modelo;
using DSIGE.Negocio;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;

namespace DSIGE.Web.Controllers
{
    public class Seguimiento_OperariosController : Controller
    {
        //
        // GET: /Seguimiento_Operarios/

        public ActionResult Inicio()
        {
            return View();
        }

        /// <summary>
        /// Autor: rcontreras
        /// Fecha: 30-09-2015
        /// </summary>
        /// <returns></returns>
        [HttpPost]
        public ActionResult JsonSeguimiento_OperariosGPS(string __a, string __b, string __c, string __d,string lista)
        {
            return new ContentResult
            {
                Content = MvcApplication._Serialize(new NUbicacion_Lecturas().NListaLecturasOperariosGPS(
                        Convert.ToInt32(__a),
                        Convert.ToString(__b),
                        __c,
                        __d,
                        ((Sesion)Session["Session_Usuario_Acceso"]).empresa.emp_id, lista
                )),
                ContentType = "application/json"
            };
        }

        /// <summary>
        /// Autor: rcontreras
        /// Fecha: 30-09-2015
        /// </summary>
        /// <returns></returns>
        [HttpPost]
        public ActionResult JsonSeguimiento_OperariosGPS2(string __a)
        {
            return new ContentResult
            {
                Content = MvcApplication._Serialize(new NSeguimiento_Operario().NSeguimientoOperariosGPS(
                        Convert.ToInt32(__a)
                )),
                ContentType = "application/json"
            };
        }


        [HttpPost]
        public ActionResult JsonSeguimiento_Operarios_GPS2(string __a, string __b, string lista)
        {
            List<UbicacionOperario_GPS> list = new List<UbicacionOperario_GPS>();
            try
            {
                
                list = new NSeguimiento_Operario().NSeguimientoOperario_GPS(Convert.ToInt32(__a), __b,lista);


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
