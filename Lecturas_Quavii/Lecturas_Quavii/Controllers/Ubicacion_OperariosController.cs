using System;
using System.Collections.Generic;
using System.Configuration;
using System.IO;
using System.Linq;
using System.Web;
using System.Web.Mvc;

using DSIGE;
using DSIGE.Modelo;
using DSIGE.Negocio;

using Excel = OfficeOpenXml;
using Style = OfficeOpenXml.Style;
using System.Web.UI;
namespace DSIGE.Web.Controllers
{
    public class Ubicacion_OperariosController : Controller
    {
        //
        // GET: /Ubicacion_Operarios/

        public ActionResult Inicio()
        {
            ViewBag.IdPerfil = ((Sesion)Session["Session_Usuario_Acceso"]).usuario.idPerfil;
            return View();
        }

        public ActionResult Inicio_new()
        {
            ViewBag.IdPerfil = ((Sesion)Session["Session_Usuario_Acceso"]).usuario.idPerfil;
            return View();
        }

        /// <summary>
        /// Autor: rcontreras
        /// Fecha: 30-09-2015
        /// </summary>
        /// <returns></returns>

        [HttpPost]
        public ActionResult JsonUbicacion_OperariosGPS( int id_local, string fecha, string servicio, int id_operario)
        { 
            return new ContentResult
            {
                Content = MvcApplication._Serialize(new NUbicacion_Operarios().NListaOperariosGPS(fecha, id_operario, 0, servicio, 0,0, id_local)),
                ContentType = "application/json"
            };
        }

        [HttpPost]
        public JsonResult ListandoServicios()
        {            
            try
            {
                Cls_Negocio_AsignarOrdenTrabajo obj_negocio = new Cls_Negocio_AsignarOrdenTrabajo();
                var loDatos = obj_negocio.Capa_Negocio_Get_ListaServicioXusuario(((Sesion)Session["Session_Usuario_Acceso"]).usuario.usu_id);
                // loDatos = obj_negocio.Capa_Negocio_Get_ListaServicioXusuario();
                return Json(loDatos, JsonRequestBehavior.AllowGet);
            }
            catch (Exception ex)
            {
               throw ex;
            }

        }


        ///// <summary>
        ///// Autor: rcontreras
        ///// Fecha: 08-09-2015
        ///// </summary>
        ///// <param name="__a"></param>
        ///// <returns></returns>
        //[HttpGet]
        //public ActionResult ExportaExcel(string __a)
        //{
        //    ViewBag.DatosExcel = new NUbicacion_Operarios().NListaOperariosGPS(Convert.ToInt32(__a));

        //    string archivo = "Operarios_Pendientes_Por_Liquidar" + DateTime.Now.ToString("yyyyMMdd_hhmmss") + ".xls";
        //    Response.AddHeader("content-disposition", "attachment; filename=" + archivo);
        //    Response.ContentType = "application/x-ms-excel";

        //    return View();
        //}

    }
}
