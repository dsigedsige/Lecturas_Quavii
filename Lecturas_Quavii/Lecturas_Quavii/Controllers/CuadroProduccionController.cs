using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using DSIGE;
using DSIGE.Modelo;
using DSIGE.Negocio;
using System.Data.OleDb;
using System.Data;

namespace DSIGE.Web.Controllers
{
    public class CuadroProduccionController : Controller
    {
        //
        // GET: /CuadroProduccion/

        public ActionResult Inicio()
        {
            return View();
        }


        [HttpPost]
        public ActionResult JsonListarLocales()
        {
            return new ContentResult
            {
                Content = MvcApplication._Serialize(new NimportarDiasTrabajo().getLocales()),
                ContentType = "application/json"
            };
        }

        [HttpPost]
        public ActionResult JsonGetOperarios(int idlocal,string lista)
        {
            return new ContentResult
            {
                Content = MvcApplication._Serialize(new NCuadro_Produccion().GetOperarios("", "", idlocal, "Operarios", lista)),
                ContentType = "application/json"
            };
        }

        [HttpPost]
        public ActionResult JsonGetFechas(string fechaInicial, string fechaFinal, int idlocal,string lista)
        {
            return new ContentResult
            {
                Content = MvcApplication._Serialize(new NCuadro_Produccion().GetFechas(fechaInicial, fechaFinal, idlocal, "Fechas", lista)),
                ContentType = "application/json"
            };
        }

        [HttpPost]
        public ActionResult JsonGetDatos(string fechaInicial, string fechaFinal, int idlocal,string lista)
        {
            return new ContentResult
            {
                Content = MvcApplication._Serialize(new NCuadro_Produccion().GetDatos(fechaInicial, fechaFinal, idlocal, "Datos",lista)),
                ContentType = "application/json"
            };
        }

    }
}
