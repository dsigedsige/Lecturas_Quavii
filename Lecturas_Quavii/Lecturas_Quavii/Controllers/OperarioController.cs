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

using System.Net;
using System.Net.Mail;
using System.Threading.Tasks;


namespace DSIGE.Web.Controllers
{
    public class OperarioController : Controller
    {
        //
        // GET: /Operario/

        public ActionResult Inicio()
        {
            return View(new NOperario().NLista());
        }

        /// <summary>
        /// Autor: jlucero
        /// Fecha: 2015-06-16
        /// </summary>
        /// <returns></returns>
        [HttpGet]
        public ActionResult JsonOperario()
        {
            return new ContentResult
            {
                Content = MvcApplication._Serialize(new NOperario().NLista()),
                ContentType = "application/json"
            };
        }

        /// <summary>
        /// Autor: jlucero
        /// Fecha: 2015-06-19
        /// </summary>
        /// <returns></returns>
        [HttpGet]
        public ActionResult JsonOperarioLocal(string __a, string __b)
        {
            return new ContentResult
            {
                Content = MvcApplication._Serialize(
                        new NOperario().NLista(
                            new Request_Operario_Empresa_Local_Opcion() {
                                emp_id = ((Sesion)Session["Session_Usuario_Acceso"]).empresa.emp_id,
                                loc_id = Convert.ToInt32(__a),
                                opcion = Convert.ToInt32(__b)
                            }
                        )
                    , true),
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
                    new NOperario().NObjeto(
                        new Request_CRUD_Operario()
                        {
                            ope_id = Convert.ToInt32(__a)
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
        /// <param name="__c"></param>
        /// <param name="__d"></param>
        /// <param name="__e"></param>
        /// <param name="__f"></param>
        /// <param name="__g"></param>
        /// <param name="__h"></param>
        /// <param name="__i"></param>
        /// <param name="__j"></param>
        /// <param name="__k"></param>
        /// <returns></returns>
        [HttpPost]
        public ActionResult Inserta(string __a, string __b, string __c, string __d, string __e,string __m, string __f, string __g, string __h, string __i, string __j, string __k, string __l, string __mail)
        {
            Int32 respuesta = new NOperario().NInserta(
                        new Request_CRUD_Operario()
                        {
                            emp_id = ((Sesion)Session["Session_Usuario_Acceso"]).empresa.emp_id, 
                            loc_id = Convert.ToInt32(__a), 
                            ope_documento = __b, 
                            ope_documento_tipo = __c, 
                            ope_apellido = __d, 
                            ope_nombre = __e,
                            ope_foto = __f, 
                            ope_celular = __g, 
                            ope_usuario = __h, 
                            ope_contrasenia = __i, 
                            ope_online = __j, 
                            ope_estado = Convert.ToInt32(__k), 
                            ope_crea = ((Sesion)Session["Session_Usuario_Acceso"]).usuario.usu_id,
                            ope_tipo_usuario = __l,
                            ope_email = __mail
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
        /// <param name="__d"></param>
        /// <param name="__e"></param>
        /// <param name="__f"></param>
        /// <param name="__g"></param>
        /// <param name="__h"></param>
        /// <param name="__i"></param>
        /// <param name="__j"></param>
        /// <param name="__k"></param>
        /// <param name="__l"></param>
        /// <returns></returns>
        [HttpPost]
        public ActionResult Actualiza(string __a, string __b, string __c, string __d, string __e, string __f, string __g, string __h, string __i, string __j, string __k, string __l, string __m, string __mail)
        {
            Int32 respuesta = new NOperario().NActualiza(
                        new Request_CRUD_Operario()
                        {
                            ope_id = Convert.ToInt32(__a),
                            emp_id = ((Sesion)Session["Session_Usuario_Acceso"]).empresa.emp_id,
                            loc_id = Convert.ToInt32(__b),
                            ope_documento = __c,
                            ope_documento_tipo = __d,
                            ope_apellido = __e,
                            ope_nombre = __f,
                            ope_foto = __g,
                            ope_celular = __h,
                            ope_usuario = __i,
                            ope_contrasenia = __j,
                            ope_online = __k,
                            ope_estado = Convert.ToInt32(__l),
                            ope_crea = ((Sesion)Session["Session_Usuario_Acceso"]).usuario.usu_id,
                            ope_tipo_usuario = __m,
                            ope_email = __mail
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
            Int32 respuesta = new NOperario().NAnula(
                        new Request_CRUD_Operario()
                        {
                            ope_id = Convert.ToInt32(__a),
                            ope_estado = Convert.ToInt32(__b),
                            ope_crea = ((Sesion)Session["Session_Usuario_Acceso"]).usuario.usu_id
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
                    new NOperario().NAuditoria(
                        new Request_CRUD_Operario()
                        {
                            ope_id = Convert.ToInt32(__a)
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
        /// 

        public ActionResult Descarga()
        {
            var grid = new System.Web.UI.WebControls.GridView();
            var collection = new NOperario().NLista();
            grid.DataSource = from d in collection
                              select new
                              {
                                  ID = d.ope_id,
                                  DOCUMENTO = d.ope_documento,
                                  APELLIDOS_Y_NOMBRES = d.ope_nombre,
                                  TIPO_USUARIO = d.ope_tipo_usuario,
                                  CELULAR = d.ope_celular,
                                  ESTADO = (d.ope_estado == 2 ? "Inactivo" : "Activo")
                              };
            grid.DataBind();

            Response.ClearContent();
            Response.AddHeader("content-disposition", "attachment; filename=Operario_Listado.xls");
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

        //    List<Operario> _lista = MvcApplication._Deserialize<List<Operario>>(__a);

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
        //        Excel.ExcelWorksheet oWs = oEx.Workbook.Worksheets.Add("Operario");

        //        oWs.Cells[1, 1].Value = "ID";
        //        oWs.Cells[1, 2].Value = "DOCUMENTO";
        //        oWs.Cells[1, 3].Value = "APELLIDOS Y NOMBRES";
        //        oWs.Cells[1, 4].Value = "TIPO USUARIO";
        //        oWs.Cells[1, 5].Value = "CELULAR";
        //        oWs.Cells[1, 6].Value = "";

        //        foreach (Operario oBj in _lista)
        //        {
        //            oWs.Cells[_fila, 1].Value = oBj.ope_id;
        //            oWs.Cells[_fila, 2].Value = oBj.ope_documento;
        //            oWs.Cells[_fila, 3].Value = oBj.ope_nombre;
        //            oWs.Cells[_fila, 4].Value = oBj.ope_tipo_usuario;
        //            oWs.Cells[_fila, 5].Value = oBj.ope_celular;
        //            oWs.Cells[_fila, 6].Value = (oBj.ope_estado == 2 ? "Inactivo" : "Activo");

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
        //        oWs.Column(5).AutoFit();
        //        oEx.Save();
        //    }

        //    return Json(new { Archivo = _servidor });
        //    //return new ContentResult
        //    //{
        //    //    Content = "{ \"__a\": \"" + _servidor + "\" }",
        //    //    ContentType = "application/json"
        //    //};
        //}

        /// <summary>
        /// Autor: jlucero
        /// Fecha: 2015-06-16
        /// </summary>
        /// <param name="__a"></param>
        /// <returns></returns>
        [HttpPost]
        public ActionResult CargaImage(HttpPostedFileBase __a)
        {
            String _extension;
            String _servidor;
            String _ruta;
            Boolean _estado;

            _extension = "";
            _servidor = "";
            _ruta = "";
            _estado = false;

            try
            {
                _extension = Path.GetExtension(__a.FileName);
                _servidor = String.Format("{0:ddMMyyyy_hhmmss}" + _extension, DateTime.Now);
                //_ruta = Path.Combine(ConfigurationManager.AppSettings["servidor-foto"], _servidor);

                _ruta = Path.Combine(Server.MapPath("/Calidda/Content/foto"), _servidor);

                __a.SaveAs(_ruta);

                _estado = true;
            }
            catch (Exception) {
                _servidor = ConfigurationManager.AppSettings["foto"];
            }

            return new ContentResult
            {
                Content = "{ \"__a\": \"" + _servidor + "\", \"__b\": " + _estado.ToString().ToLower() + " }",
                ContentType = "application/json"
            };
        }



        /// <summary>
        /// Autor: rcontreras
        /// Fecha: 20-05-2015
        /// </summary>
        /// <returns></returns>
        [HttpPost]
        public ActionResult JsonRecuperaCorreo(string __a)
        {
            return new ContentResult
            {
                Content = MvcApplication._Serialize(
                        new NUsuario().NRecuperaClave(
                            new Request_Usuario_Recupera_Clave()
                            {
                                usu_email = __a
                            }
                        )
                    , true),
                ContentType = "application/json"
            };
        }

        [HttpPost]
        public ActionResult enviarCorreoRecuparaContrasena(string __a, string __b, string __c)
            
            {
            if (ModelState.IsValid)
            {
                var body = "<center><h1>Recuperación de Contraseña</h1></center> <p> Usuario : " + __a + ", su contraseña es : " + __c + 
                    "</p><p>Ingrese al sistema :</p><p></p><p></p><p>Atte.</p><p>Administrador Web</p><p>Lecturas</p>";
                var message = new MailMessage();
                message.To.Add(new MailAddress(__b));
                message.From = new MailAddress("cobralecturas@gmail.com");
                message.Subject = "Recuperación de Correo - Sistema de Lecturas";
                message.Body = body;
                message.IsBodyHtml = true;

                using (var smtp = new SmtpClient())
                {
                    var credential = new NetworkCredential
                    {
                        //UserName = "cobra.lecturas@hotmail.com",
                        //Password = "Aa.123456"

                        UserName = "cobralecturas@gmail.com",
                        Password = "A.123456"
                    };
                    smtp.Credentials = credential;
                    //smtp.Host = "smtp.live.com";
                    //smtp.Port = 25;
                    smtp.Host = "smtp.gmail.com";
                    smtp.Port = 587;

                    smtp.EnableSsl = true;
                    smtp.Send(message);
                    return new ContentResult
                    {
                        Content = MvcApplication._Serialize(message),
                        ContentType = "application/json"
                    };
                }
            }
            else
            {
                return View();
            }


        }

    }
}
