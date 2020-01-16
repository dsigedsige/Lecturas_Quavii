using DSIGE.Modelo;
using DSIGE.Negocio;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;

namespace DSIGE.Web.Controllers
{
    public class Historico_LecturasController : Controller
    {
        //
        // GET: /Historico_Lecturas/

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
        public ActionResult JsonListaLecturaHistoricoXSuministro(string __a, string __c, string __d, string __e, string __f,string __g)
        {
            return new ContentResult
            {
                Content = MvcApplication._Serialize(
                        new NLectura().NListaLecturaHistoricoXSuministro(
                            new Request_Lectura_Tomadas()
                            {
                                local = Convert.ToInt32(__a),
                                f_ini = Convert.ToString(__c),
                                f_fin = Convert.ToString(__d),
                                suministro = Convert.ToString(__e),
                                medidor = Convert.ToString(__f),
                                lista = Convert.ToString(__g)
                            }
                        )
                    , true),
                ContentType = "application/json"
            };
        }

        [HttpPost]
        public ActionResult FotosLectura(string __a, string tiposervicio)
        {
            return new ContentResult
            {
                Content = MvcApplication._Serialize(
                        new NEnvioTrabajoCliLecReLec().NListaFotos_Historico(
                            Convert.ToInt32(__a),
                            Convert.ToInt32(tiposervicio)
                        )
                    , true),
                ContentType = "application/json"
            };
        }


    }
}
