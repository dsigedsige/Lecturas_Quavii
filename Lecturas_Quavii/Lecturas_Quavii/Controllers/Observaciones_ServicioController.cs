using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Web;
using System.Web.Mvc;

using DSIGE;
using DSIGE.Modelo;
using DSIGE.Negocio;

namespace DSIGE.Web.Controllers
{
    public class Observaciones_ServicioController : Controller
    {
        //
        // GET: /Observaciones_Servicio/

        public ActionResult Inicio()
        {
            return View();
        }

        /// <summary>
        /// Autor: rcontreras
        /// Fecha: 23-09-2015
        /// </summary>
        /// <returns></returns>
        [HttpGet]
        public ActionResult JsonObservacionServicio(string __a)
        {
            return new ContentResult
            {
                Content = MvcApplication._Serialize(
                        new NObservacion_Servicio().NLista(
                               ((Sesion)Session["Session_Usuario_Acceso"]).empresa.emp_id,
                                Convert.ToInt32(__a),
                                Convert.ToInt32('0')
                        )
                    , true),
                ContentType = "application/json"
            };
        }


        /// <summary>
        /// Autor: rcontreras
        /// Fecha: 26-09-2015
        /// </summary>
        /// <param name="__a"></param>
        /// <returns></returns>
        [HttpGet]
        public ActionResult JsonServicioObservacion(string __a)
        {
            return new ContentResult
            {
                Content = MvcApplication._Serialize(new NObservacion_Servicio().NObservacionServicio(
                                __a
                ), true),
                ContentType = "application/json"
            };
        }



        /// <summary>
        /// Autor: rcontreras
        /// Fecha: 26-09-2015
        /// </summary>
        /// <returns></returns>
        [HttpPost]
        public ActionResult AsignaServicio(string __a, string __b, string __c)
        {
            Int32 respuesta = new NObservacion_Servicio().NAsignaServicio(
                            __a,
                            Convert.ToInt32(__b),
                            Convert.ToInt32(__c),
                            ((Sesion)Session["Session_Usuario_Acceso"]).usuario.usu_id,
                            ((Sesion)Session["Session_Usuario_Acceso"]).empresa.emp_id
                    );

            return new ContentResult
            {
                Content = "{ \"__a\": " + respuesta + " }",
                ContentType = "application/json"
            };
        }

    }
}
