using DSIGE.Modelo;
using DSIGE.Negocio;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;

namespace DSIGE.Web.Controllers
{
    public class Resumen_Lecturas_ObservacionController : Controller
    {
        //
        // GET: /Resumen_Lecturas_Observacion/

        public ActionResult Inicio()
        {
            return View();
        }

        /// <summary>
        /// Autor: rcontreras
        /// Fecha: 19-11-2015
        /// </summary>
        /// <returns></returns>
        [HttpPost]
        public ActionResult JsonListaLecturasResumenObs(string __a, string __c, string __d,string __e, string lista)
        {
            return new ContentResult
            {
                Content = MvcApplication._Serialize(
                        new NLectura().NListaLecturasResumenObs(
                            new Request_Lectura_Historico()
                            {
                                local = Convert.ToInt32(__a),
                                f_ini = Convert.ToString(__c),
                                f_fin = Convert.ToString(__d)                               
                            },lista
                        )
                    , true),
                ContentType = "application/json"
            };
        }

    }
}
