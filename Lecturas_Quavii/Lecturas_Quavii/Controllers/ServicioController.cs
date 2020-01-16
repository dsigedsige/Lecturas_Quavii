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
    public class ServicioController : Controller
    {
        //
        // GET: /Servicio/

        [HttpGet]
        public ActionResult Inicio()
        {
            return View(new NServicio().NLista());
        }


        /// <summary>
        /// Autor: jlucero
        /// Fecha: 2015-06-16
        /// </summary>
        /// <returns></returns>
        [HttpGet]
        public ActionResult JsonServicio() {
            return new ContentResult
            {
                Content = MvcApplication._Serialize(new NServicio().NLista()),
                ContentType = "application/json"
            };
        }

        /// <summary>
        /// Autor: jlucero
        /// Fecha: 2015-06-19
        /// </summary>
        /// <param name="__a"></param>
        /// <returns></returns>
        [HttpGet]
        public ActionResult JsonServicioOperario(string __a)
        {
            return new ContentResult
            {
                Content = MvcApplication._Serialize(new NServicio().NLista(new Request_Servicio_Operario() { ope_id = __a }), true),
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
        public ActionResult Busca(string __a) {
            return new ContentResult
            {
                Content = MvcApplication._Serialize(
                    new NServicio().NObjeto(
                        new Request_CRUD_Servicio() { 
                            id = Convert.ToInt32(__a)
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
        public ActionResult Inserta(string __a, string __b)
        {
            Int32 respuesta = new NServicio().NInserta(
                        new Request_CRUD_Servicio()
                        {
                            descripcion = __a,
                            estado = Convert.ToInt32(__b),
                            usuario = ((Sesion)Session["Session_Usuario_Acceso"]).usuario.usu_id
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
        public ActionResult Actualiza(string __a, string __b, string __c)
        {
            Int32 respuesta = new NServicio().NActualiza(
                        new Request_CRUD_Servicio()
                        {
                            id = Convert.ToInt32(__a),
                            descripcion = __b,
                            estado = Convert.ToInt32(__c),
                            usuario = ((Sesion)Session["Session_Usuario_Acceso"]).usuario.usu_id
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
            Int32 respuesta = new NServicio().NAnula(
                        new Request_CRUD_Servicio()
                        {
                            id = Convert.ToInt32(__a),
                            estado = Convert.ToInt32(__b),
                            usuario = ((Sesion)Session["Session_Usuario_Acceso"]).usuario.usu_id
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
                    new NServicio().NAuditoria(
                        new Request_CRUD_Servicio()
                        {
                            id = Convert.ToInt32(__a)
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

        public ActionResult Descarga()
        {
            var grid = new System.Web.UI.WebControls.GridView();
            var collection = new NServicio().NLista();
            grid.DataSource = from d in collection
                              select new
                              {
                                  ID = d.ser_id,
                                  DESCRIPCION = d.ser_descripcion,
                                  ESTADO = (d.ser_estado == 2 ? "Inactivo" : "Activo")
                              };
            grid.DataBind();

            Response.ClearContent();
            Response.AddHeader("content-disposition", "attachment; filename=Servicio_Listado.xls");
            Response.ContentType = "application/x-ms-excel";
            StringWriter sw = new StringWriter();
            HtmlTextWriter htw = new HtmlTextWriter(sw);

            grid.RenderControl(htw);

            Response.Write(sw.ToString());

            Response.End();

            return null;
        }

        //public JsonResult Descarga(string __a)
        //{
        //    Int32 _fila = 2;
        //    String _servidor;
        //    String _ruta;

        //    //if (__a.Length == 0) {
        //    //    return View();
        //    //}

        //    List<Servicio> _lista = MvcApplication._Deserialize<List<Servicio>>(__a);

        //    _servidor = String.Format("{0:ddMMyyyy_hhmmss}.xlsx", DateTime.Now);
        //    _ruta = Path.Combine(Server.MapPath("/Temp"), _servidor);

        //    FileInfo _file = new FileInfo(_ruta);
        //    if (_file.Exists)
        //    {
        //        _file.Delete();
        //        _file = new FileInfo(_ruta);
        //    }

        //    using (Excel.ExcelPackage oEx = new Excel.ExcelPackage(_file)) {
        //        Excel.ExcelWorksheet oWs = oEx.Workbook.Worksheets.Add("Servicio");

        //        oWs.Cells[1, 1].Value = "ID";
        //        oWs.Cells[1, 2].Value = "DESCRIPCION";
        //        oWs.Cells[1, 3].Value = "";

        //        foreach (Servicio oBj in _lista)
        //        {
        //            oWs.Cells[_fila, 1].Value = oBj.ser_id;
        //            oWs.Cells[_fila, 2].Value = oBj.ser_descripcion;
        //            oWs.Cells[_fila, 3].Value = (oBj.ser_estado == 2 ? "Inactivo" : "Activo");

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

        //        oEx.Save();
        //    }

        //    return Json(new { Archivo = _servidor });
        //    //return new ContentResult
        //    //{
        //    //    Content = "{ \"__a\": \"" + _servidor + "\" }",
        //    //    ContentType = "application/json"
        //    //};
        //}
    }
}