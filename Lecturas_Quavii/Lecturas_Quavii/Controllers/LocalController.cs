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
    public class LocalController : Controller
    {
        //
        // GET: /Local/

        public ActionResult Inicio()
        {
            return View();
        }

        /// <summary>
        /// Autor: rcontreras
        /// Fecha: 25-09-2015
        /// </summary>
        /// <returns></returns>
        [HttpGet]
        public ActionResult JsonLocal()
        {
            return new ContentResult
            {
                Content = MvcApplication._Serialize(new NLocal().NLista(
                    new Request_CRUD_Local()
                    {
                        emp_id = ((Sesion)Session["Session_Usuario_Acceso"]).empresa.emp_id
                    }
                )),
                ContentType = "application/json"
            };
        }

        /// <summary>
        /// Autor: rcontreras
        /// Fecha: 26-09-2015
        /// </summary>
        /// <param name="__a"></param>
        /// <returns></returns>
        [HttpPost]
        public ActionResult Busca(string __a)
        {
            return new ContentResult
            {
                Content = MvcApplication._Serialize(
                    new NLocal().NObjeto(
                        new Request_CRUD_Local()
                        {
                            loc_id = Convert.ToInt32(__a)
                        }
                    )),
                ContentType = "application/json"
            };
        }

        /// <summary>
        /// Autor: rcontreras
        /// Fecha: 26-09-2015
        /// </summary>
        /// <returns></returns>
        [HttpPost]
        public ActionResult Inserta(string __a, string __b, string __c, string __d, string __e)
        {
            Int32 respuesta = new NLocal().NInserta(
                        new Request_CRUD_Local()
                        {
                            emp_id = ((Sesion)Session["Session_Usuario_Acceso"]).empresa.emp_id,
                            loc_nombre = __a,
                            loc_direccion = __b,
                            loc_latitud = __c,
                            loc_longitud = __d,
                            loc_estado = Convert.ToInt32(__e),
                            loc_usuario = ((Sesion)Session["Session_Usuario_Acceso"]).usuario.usu_id
                        }
                    );

            return new ContentResult
            {
                Content = "{ \"__a\": " + respuesta + " }",
                ContentType = "application/json"
            };
        }


        /// <summary>
        /// Autor: rcontreras
        /// Fecha: 26-09-2015
        /// </summary>
        /// <returns></returns>
        [HttpPost]
        public ActionResult Actualiza(string __a, string __b, string __c, string __d, string __e, string __f)
        {
            Int32 respuesta = new NLocal().NActualiza(
                        new Request_CRUD_Local()
                        {
                            loc_id = Convert.ToInt32(__a),
                            emp_id = ((Sesion)Session["Session_Usuario_Acceso"]).empresa.emp_id,
                            loc_nombre = __b,
                            loc_direccion = __c,
                            loc_latitud = __d,
                            loc_longitud = __e,
                            loc_estado = Convert.ToInt32(__f),
                            loc_usuario = ((Sesion)Session["Session_Usuario_Acceso"]).usuario.usu_id

                        }
                    );

            return new ContentResult
            {
                Content = "{ \"__a\": " + respuesta + " }",
                ContentType = "application/json"
            };
        }


        /// <summary>
        /// Autor: rcontreras
        /// Fecha: 26-09-2015
        /// </summary>
        /// <param name="__a"></param>
        /// <param name="__b"></param>
        /// <returns></returns>
        [HttpPost]
        public ActionResult Anula(string __a, string __b)
        {
            Int32 respuesta = new NLocal().NAnula(
                        new Request_CRUD_Local()
                        {
                            loc_id = Convert.ToInt32(__a),
                            loc_estado = Convert.ToInt32(__b),
                            loc_usuario = ((Sesion)Session["Session_Usuario_Acceso"]).usuario.usu_id
                        }
                    );

            return new ContentResult
            {
                Content = "{ \"__a\": " + respuesta + " }",
                ContentType = "application/json"
            };
        }

        /// <summary>
        /// Autor: rcontreras
        /// Fecha: 26-09-2015
        /// </summary>
        /// <param name="__a"></param>
        /// <returns></returns>
        [HttpPost]
        public ActionResult Auditoria(string __a)
        {
            return new ContentResult
            {
                Content = MvcApplication._Serialize(
                    new NLocal().NAuditoria(
                        new Request_CRUD_Local()
                        {
                            loc_id = Convert.ToInt32(__a)
                        }
                    )
                ),
                ContentType = "application/json"
            };
        }

        /// <summary>
        /// Autor: rcontreras
        /// Fecha: 26-09-2015
        /// </summary>
        /// <param name="__a"></param>
        /// <returns></returns>
        /// 

 


        public ActionResult Descarga(string __a)
        {
            var grid = new System.Web.UI.WebControls.GridView();
            //List<Local> _lista = MvcApplication._Deserialize<List<Local>>(__a);
            var collection = new NLocal().NLista(

                new Request_CRUD_Local()
                {
                    emp_id = 0,
                    loc_direccion = "",
                    loc_estado = 0,
                    loc_id = 0,
                    loc_latitud = "",
                    loc_longitud = "",
                    loc_nombre = "",
                    loc_usuario = 0
                }
                );
            grid.DataSource = from d in collection
                              select new
                              {
                                  ID = d.loc_id,
                                  LOCAL = d.loc_nombre,
                                  DIRECCION = d.loc_direccion,
                                  LATITUD = d.loc_latitud,
                                  LONGITUD = d.loc_longitud,
                                  ESTADO = (d.loc_estado == 2 ? "Inactivo" : "Activo")
                              };
            grid.DataBind();

            Response.ClearContent();
            Response.AddHeader("content-disposition", "attachment; filename=Local_Listado.xls");
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

        //    //if (__a.Length == 0)
        //    //{
        //    //    return View();
        //    //}

        //    List<Local> _lista = MvcApplication._Deserialize<List<Local>>(__a);

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
        //        Excel.ExcelWorksheet oWs = oEx.Workbook.Worksheets.Add("Local");

        //        oWs.Cells[1, 1].Value = "ID";
        //        oWs.Cells[1, 2].Value = "LOCAL";
        //        oWs.Cells[1, 3].Value = "DIRECCION";
        //        oWs.Cells[1, 4].Value = "LATITUD";
        //        oWs.Cells[1, 5].Value = "LONGITUD";
        //        oWs.Cells[1, 6].Value = "";

        //        foreach (Local oBj in _lista)
        //        {
        //            oWs.Cells[_fila, 1].Value = oBj.loc_id;
        //            oWs.Cells[_fila, 2].Value = oBj.loc_nombre;
        //            oWs.Cells[_fila, 3].Value = oBj.loc_direccion;
        //            oWs.Cells[_fila, 4].Value = oBj.loc_latitud;
        //            oWs.Cells[_fila, 5].Value = oBj.loc_longitud;
        //            oWs.Cells[_fila, 6].Value = (oBj.loc_estado == 2 ? "Inactivo" : "Activo");

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
        //        oWs.Column(6).AutoFit();

        //        oEx.Save();
        //    }

        //    //return new ContentResult
        //    //{
        //    //    Content = "{ \"__a\": \"" + _servidor + "\" }",
        //    //    ContentType = "application/json"
        //    //};
        //    return Json(new { Archivo = _servidor });//_fileServer });
        //}

    }
}
