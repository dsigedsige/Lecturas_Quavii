using System;
using System.Collections.Generic;
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
    public class ObservacionController : Controller
    {
        //
        // GET: /Observacion/

        public ActionResult Inicio()
        {
            return View(new NObservacion().NLista());
        }

        /// <summary>
        /// Autor: jlucero
        /// Fecha: 2015-06-16
        /// </summary>
        /// <returns></returns>
        [HttpGet]
        public ActionResult JsonObservacion()
        {
            return new ContentResult
            {
                Content = MvcApplication._Serialize(new NObservacion().NLista()),
                ContentType = "application/json"
            };
        }

        /// <summary>
        /// Autor: jlucero
        /// Fecha: 2015-06-14
        /// </summary>
        /// <param name="__a"></param>
        /// <returns></returns>
        [HttpPost]
        public ActionResult Busca(string __a)
        {
            return new ContentResult
            {
                Content = MvcApplication._Serialize(
                    new NObservacion().NObjeto(
                        new Request_CRUD_Observacion()
                        {
                            obs_id = Convert.ToInt32(__a)
                        }
                    )),
                ContentType = "application/json"
            };
        }

        /// <summary>
        /// Autor: jlucero
        /// Fecha: 2015-06-14
        /// </summary>
        /// <param name="__a"></param>
        /// <param name="__b"></param>
        /// <returns></returns>
        [HttpPost]
        public ActionResult Inserta(string __a, string __b, string __c, string __d, string __e, string __f)
        {
            Int32 respuesta = new NObservacion().NInserta(
                        new Request_CRUD_Observacion()
                        {
                            emp_id = ((Sesion)Session["Session_Usuario_Acceso"]).empresa.emp_id, 
                            obs_abreviatura = __a, 
                            obs_descripcion = __b, 
                            gde_id = Convert.ToInt32(__c),
                            obs_estado = Convert.ToInt32(__d), 
                            obs_usuario = ((Sesion)Session["Session_Usuario_Acceso"]).usuario.usu_id,
                            obs_pideFoto = Convert.ToString(__e),
                            obs_noPideFoto = Convert.ToString(__f)
                        }
                    );

            return new ContentResult
            {
                Content = "{ \"__a\": " + respuesta + " }",
                ContentType = "application/json"
            };
        }

        /// <summary>
        /// Autor: jlucero
        /// Fecha: 2015-06-14
        /// </summary>
        /// <param name="__a"></param>
        /// <param name="__b"></param>
        /// <param name="__c"></param>
        /// <returns></returns>
        [HttpPost]
        public ActionResult Actualiza(string __a, string __b, string __c, string __d, string __e, string __f, string __g)
        {
            Int32 respuesta = new NObservacion().NActualiza(
                        new Request_CRUD_Observacion()
                        {
                            obs_id = Convert.ToInt32(__a), 
                            emp_id = ((Sesion)Session["Session_Usuario_Acceso"]).empresa.emp_id, 
                            obs_abreviatura = __b, 
                            obs_descripcion = __c, 
                            gde_id = Convert.ToInt32(__d), 
                            obs_estado = Convert.ToInt32(__e), 
                            obs_usuario = ((Sesion)Session["Session_Usuario_Acceso"]).usuario.usu_id,
                            obs_pideFoto = Convert.ToString(__f),
                            obs_noPideFoto = Convert.ToString(__g)
                        }
                    );

            return new ContentResult
            {
                Content = "{ \"__a\": " + respuesta + " }",
                ContentType = "application/json"
            };
        }

        /// <summary>
        /// Autor: jlucero
        /// Fecha: 2015-06-14
        /// </summary>
        /// <param name="__a"></param>
        /// <param name="__b"></param>
        /// <returns></returns>
        [HttpPost]
        public ActionResult Anula(string __a, string __b)
        {
            Int32 respuesta = new NObservacion().NAnula(
                        new Request_CRUD_Observacion()
                        {
                            obs_id = Convert.ToInt32(__a),
                            obs_estado = Convert.ToInt32(__b),
                            obs_usuario = ((Sesion)Session["Session_Usuario_Acceso"]).usuario.usu_id
                        }
                    );

            return new ContentResult
            {
                Content = "{ \"__a\": " + respuesta + " }",
                ContentType = "application/json"
            };
        }

        /// <summary>
        /// Autor: jlucero
        /// Fecha: 2015-06-14
        /// </summary>
        /// <param name="__a"></param>
        /// <returns></returns>
        [HttpPost]
        public ActionResult Auditoria(string __a)
        {
            return new ContentResult
            {
                Content = MvcApplication._Serialize(
                    new NObservacion().NAuditoria(
                        new Request_CRUD_Observacion()
                        {
                            obs_id = Convert.ToInt32(__a)
                        }
                    )
                ),
                ContentType = "application/json"
            };
        }

        /// <summary>
        /// Autor: jlucero
        /// Fecha: 2015-06-16
        /// </summary>
        /// <param name="__a"></param>
        /// <returns></returns>
        //[HttpPost]
        //public JsonResult Descarga(string __a)
        //{
        //    Int32 _fila = 2;
        //    String _servidor;
        //    String _ruta;

        //    //if (__a.Length == 0)
        //    //{
        //    //    return View();
        //    //}



        //    List<Observacion> _lista = MvcApplication._Deserialize<List<Observacion>>(__a);

        //    _servidor = String.Format("{0:ddMMyyyy_hhmmss}.xlsx", DateTime.Now);
        //    _ruta = Path.Combine(Server.MapPath("/Temp"), _servidor);

        //    FileInfo _file = new FileInfo(_ruta);
        //    if (_file.Exists)
        //    {
        //        _file.Delete();
        //        _file = new FileInfo(_ruta);
        //    }

        //    using (Excel.ExcelPackage oEx = new Excel.ExcelPackage(_file))
        //    {
        //        Excel.ExcelWorksheet oWs = oEx.Workbook.Worksheets.Add("Observacion");

        //        oWs.Cells[1, 1].Value = "ID";
        //        oWs.Cells[1, 2].Value = "ABREVIATURA";
        //        oWs.Cells[1, 3].Value = "DESCRIPCION";
        //        oWs.Cells[1, 4].Value = "GRUPO";
        //        oWs.Cells[1, 5].Value = "";

        //        foreach (Observacion oBj in _lista)
        //        {
        //            oWs.Cells[_fila, 1].Value = oBj.obs_id;
        //            oWs.Cells[_fila, 2].Value = oBj.obs_abreviatura;
        //            oWs.Cells[_fila, 3].Value = oBj.obs_descripcion;
        //            oWs.Cells[_fila, 4].Value = oBj.gde_descripcion;
        //            oWs.Cells[_fila, 5].Value = (oBj.obs_estado == 2 ? "Inactivo" : "Activo");

        //            _fila++;
        //        }

        //        oWs.Row(1).Style.WrapText = true;
        //        oWs.Row(1).Style.Font.Bold = true;
        //        oWs.Row(1).Style.HorizontalAlignment = Style.ExcelHorizontalAlignment.Center;
        //        oWs.Row(1).Style.VerticalAlignment = Style.ExcelVerticalAlignment.Center;

        //        oWs.Column(1).Style.Font.Bold = true;
        //        oWs.Column(1).AutoFit();
        //        oWs.Column(2).AutoFit();
        //        oWs.Column(3).AutoFit();
        //        oWs.Column(4).AutoFit();
        //        oWs.Column(5).AutoFit();

        //        oEx.Save();
        //    }

        //    Response.ClearContent();
        //    Response.AddHeader("content-disposition", "attachment; filename=Cuenta_corriente.xls");
        //    Response.ContentType = "application/json";

        //    return Json(new { Archivo = _servidor });
        //    //return new ContentResult
        //    //{
        //    //    Content = "{ \"__a\": \"" + _servidor + "\" }",
        //    //    ContentType = "application/x-ms-excel"
        //    //};
        //}

        public ActionResult Descarga(string __a)
        {
            var grid = new System.Web.UI.WebControls.GridView();
            //var collection =  List<Observacion>(__a);
            var collection = new NObservacion().NLista();
            //var collection = CuentaCorrienteAD.Instancia.ListarAll("", 10);
            grid.DataSource = from d in collection
                              select new
                              {
                                  ID = d.obs_id,
                                  ABREVIATURA = d.obs_abreviatura,
                                  DESCRIPCION = d.obs_descripcion,
                                  GRUPO = d.gde_descripcion,
                                  ESTADO = (d.obs_estado == 2 ? "Inactivo" : "Activo"),
                                  PIDE_FOTO = (d.obs_pideFoto == "2" ? "Inactivo" : "Activo"),
                                  NO_PIDE_FOTO = (d.obs_noPideFoto == "2" ? "Inactivo" : "Activo")
                              };
            grid.DataBind();

            Response.ClearContent();
            Response.AddHeader("content-disposition", "attachment; filename=Observacion_Listado.xls");
            Response.ContentType = "application/x-ms-excel";
            StringWriter sw = new StringWriter();
            HtmlTextWriter htw = new HtmlTextWriter(sw);

            grid.RenderControl(htw);

            Response.Write(sw.ToString());

            Response.End();

            return null;
        }

    }
}