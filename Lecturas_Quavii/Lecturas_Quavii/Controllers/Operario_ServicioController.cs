using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;

using DSIGE;
using DSIGE.Modelo;
using DSIGE.Negocio;

namespace DSIGE.Web.Controllers
{
    public class Operario_ServicioController : Controller
    {
        //
        // GET: /Operario_Servicio/

        /// <summary>
        /// Autor: jlucero
        /// Fecha: 2015-06-21
        /// </summary>
        /// <returns></returns>
        public ActionResult Inicio()
        {
            return View();
        }

        /// <summary>
        /// Autor: jlucero
        /// Fecha: 2015-06-21
        /// </summary>
        /// <param name="__a"></param>
        /// <param name="__b"></param>
        /// <param name="__c"></param>
        /// <returns></returns>
        [HttpPost]
        public ActionResult AsignaServicio(string __a, string __b, string __c, string __d) {
            Int32 respuesta = new NOperario_Servicio().NAsignaServicio(
                        new Request_Operario_Servicio_Asignacion()
                        {
                            ope_id = __a,
                            ser_id = Convert.ToInt32(__b),
                            usu_id = ((Sesion)Session["Session_Usuario_Acceso"]).usuario.usu_id,
                            sop_estado = Convert.ToInt32(__c),
                            emp_id = ((Sesion)Session["Session_Usuario_Acceso"]).empresa.emp_id,
                            loc_id = Convert.ToInt32(__d)
                        }
                    );

            return new ContentResult
            {
                Content = "{ \"__a\": " + respuesta + " }",
                ContentType = "application/json"
            };
        }
    }
}
