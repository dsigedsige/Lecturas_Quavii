using DSIGE.Modelo;
using DSIGE.Negocio;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.IO;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using Excel = OfficeOpenXml;
using Style = OfficeOpenXml.Style;

namespace DSIGE.Web.Controllers
{
    public class Envia_Trabajos_CliController : Controller
    {
        //
        // GET: /Envia_Trabajos_Cli/

        public ActionResult Inicio()
        {
            return View();
        }


        public ActionResult ExportarExcel_VisualizarEnvioAlCliente(int id_Local, int id_operario, string  suministro_lectura , string  medidor_lectura , string  fechaAsignacion_Lectura, int estado )
        {
                NEnvioTrabajoCliLecReLec obj_negocio = new NEnvioTrabajoCliLecReLec();
                ViewBag.DatosExcel = obj_negocio.Capa_Negocio_Excel_Envio_Trabajo_cliente(id_Local,id_operario, suministro_lectura, medidor_lectura, fechaAsignacion_Lectura, estado);
                
                string archivo = "ReporteExcel_TrabajoCliLecReLec" + DateTime.Now.ToString("yyyyMMdd_hhmmss") + ".xls";
                Response.AddHeader("content-disposition", "attachment; filename=" + archivo);
                Response.ContentType = "application/ms-excel";
                ViewBag.fecha = fechaAsignacion_Lectura;
                return View();
        }

        ////Edicion de Registros de Envio de trabajos cliente 

        public static string _Serialize(object value, bool ignore = false)
        {
            var SerializerSettings = new JsonSerializerSettings()
            {
                MaxDepth = Int32.MaxValue,
                NullValueHandling = (ignore == true ? NullValueHandling.Ignore : NullValueHandling.Include)
            };

            return JsonConvert.SerializeObject(value, Formatting.Indented, SerializerSettings);
        }


        // listando

        [HttpPost]
        public string ListandoOperario()
        {
            object loDatos;
            try
            {
                NEnvioTrabajoCliLecReLec obj_negocio = new NEnvioTrabajoCliLecReLec();
                loDatos = obj_negocio.Capa_Negocio_Listado_Operario();

                return _Serialize(loDatos, true);

            }
            catch (Exception ex)
            {
                throw ex;
            }


        }

        // listando 

        [HttpPost]
        public string ListandoObservaciones()
        {
            object loDatos;
            try
            {
                NEnvioTrabajoCliLecReLec obj_negocio = new NEnvioTrabajoCliLecReLec();
                loDatos = obj_negocio.Capa_Negocio_Listado_Observaciones();

                return _Serialize(loDatos, true);

            }
            catch (Exception ex)
            {
                throw ex;
            }


        }

      //  ListandoModoEdicion


        [HttpPost]
        public string ListandoModoEdicion(int id_lectura)
        {
            object loDatos;
            try
            {
                NEnvioTrabajoCliLecReLec obj_negocio = new NEnvioTrabajoCliLecReLec();
                loDatos = obj_negocio.Capa_Negocio_Listando_Registros_Modo_edicion(id_lectura);

                return _Serialize(loDatos, true);

            }
            catch (Exception ex)
            {
                throw ex;
            }


        }


        // acutualizando los datos
        

        [HttpPost]
        public string GuardandoInformacion(int id_Lectura, string id_Operario_Lectura, string Lectura_Manual_Lectura, DateTime fechaLecturaMovil_lectura, string id_Observacion_Lectura)
        {
            bool loDatos;
            try
            {
                NEnvioTrabajoCliLecReLec obj_negocio = new NEnvioTrabajoCliLecReLec();
                var id_user = ((Sesion)Session["Session_Usuario_Acceso"]).usuario.usu_id;

                loDatos = obj_negocio.Capa_Negocio_Guardar_Informacion(id_Lectura, id_Operario_Lectura, Lectura_Manual_Lectura, fechaLecturaMovil_lectura, id_Observacion_Lectura, id_user);

                return _Serialize(loDatos, true);

            }
            catch (Exception ex)
            {
                throw ex;
            }


        }
        







        

        /// <summary>
        /// Autor: rcontreras
        /// Fecha: 20-05-2015
        /// </summary>
        /// <returns></returns>
        [HttpPost]
        public ActionResult JsonListaEnviarClienteLecReLec(string __a, string __b, string __c, string __d, string __e, string __f, string __g)
        {
            return new ContentResult
            {
                Content = MvcApplication._Serialize(
                        new NEnvioTrabajoCliLecReLec().NLista(
                            new Request_Asigna_Lectura_ReLectura()
                            {
                                emp_id = ((Sesion)Session["Session_Usuario_Acceso"]).empresa.emp_id,
                                tip_ser_id = Convert.ToInt32(__a),
                                id_local = Convert.ToInt32(__b),
                                suministro = __c,
                                medidor = __d,
                                tecnico_asig_id = Convert.ToInt32(__e),
                                estado_asig = Convert.ToInt32(__f),
                                fecha_asig = Convert.ToDateTime(__g)
                            }
                        )
                    , true),
                ContentType = "application/json"
            };
        }

        /// <summary>
        /// Autor: rcontreras
        /// Fecha: 28-09-2015
        /// </summary>
        /// <returns></returns>
        [HttpPost]
        public ActionResult ActualizaEnvioCliente(string __a, string __b)
        {
            Int32 respuesta = new NEnvioTrabajoCliLecReLec().NActualizaEnviaCliente(
                            __a,
                            ((Sesion)Session["Session_Usuario_Acceso"]).usuario.usu_id,
                            ((Sesion)Session["Session_Usuario_Acceso"]).empresa.emp_id
                    );

            return new ContentResult
            {
                Content = "{ \"__a\": " + respuesta + " }",
                ContentType = "application/json"
            };
        }

        [HttpPost]
        public ActionResult FotosLectura(string __a)
        {
            return new ContentResult
            {
                Content = MvcApplication._Serialize(
                        new NEnvioTrabajoCliLecReLec().NListaFotos(
                            Convert.ToInt32(__a)
                        )
                    , true),
                ContentType = "application/json"
            };
        }

        [HttpPost]
        public ActionResult FotoSelfie_Reparto(string fecha, int id_operario)
        {
            return new ContentResult
            {
                Content = MvcApplication._Serialize( new NEnvioTrabajoCliLecReLec().NFoto_selfie_reparto(fecha, id_operario), true),
                ContentType = "application/json"
            };
        }


        /// <summary>
        /// Autor: rcontreras
        /// Fecha: 29-09-2015
        /// </summary>
        /// <returns></returns>
        [HttpPost]
        public ActionResult ListaLecturasPendientes(string __a)
        {
            return new ContentResult
            {
                Content = MvcApplication._Serialize(
                        new NEnvioTrabajoCliLecReLec().NListaLecturasPendientes(
                            Convert.ToDateTime(__a),
                           ((Sesion)Session["Session_Usuario_Acceso"]).empresa.emp_id
                        )
                    , true),
                ContentType = "application/json"
            };
        }


        /// <summary>
        /// Autor: rcontreras
        /// Fecha: 29-09-2015
        /// </summary>
        /// <param name="__a"></param>
        /// <returns></returns>
        [HttpPost]
        public ActionResult DescargaListaPreVisualizaEnvio(string __a, string __b)
        {
            string tituloHoja = "";
            if (__b == "1")
            {
                tituloHoja = "EnvioAlCliente";
            }
            else
            {
                tituloHoja = "PreVisualizaEnvioAlCliente";
            }
            Int32 _fila = 2;
            String _servidor;
            String _ruta;

            if (__a.Length == 0)
            {
                return View();
            }

            _servidor = String.Format("{0:ddMMyyyy_hhmmss}.xlsx", DateTime.Now);
            _ruta = Path.Combine(Server.MapPath("/Lecturas/Temp"), _servidor);

            FileInfo _file = new FileInfo(_ruta);
            if (_file.Exists)
            {
                _file.Delete();
                _file = new FileInfo(_ruta);
            }

            using (Excel.ExcelPackage oEx = new Excel.ExcelPackage(_file))
            {

                List<EnvioTrabajoCliLecReLec> _lista = new NEnvioTrabajoCliLecReLec().NListaPreVisualizaEnvio(
                        __a
                    );
                
                Excel.ExcelWorksheet oWs = oEx.Workbook.Worksheets.Add(tituloHoja);

                oWs.Cells[1, 1].Value = "ORDEN";
                oWs.Cells[1, 2].Value = "SECTOR";
                oWs.Cells[1, 3].Value = "ZONA";
                oWs.Cells[1, 4].Value = "CORRELATIVO";
                oWs.Cells[1, 5].Value = "SUMINISTRO";
                oWs.Cells[1, 6].Value = "MEDIDOR";
                oWs.Cells[1, 7].Value = "LECTURA";
                oWs.Cells[1, 8].Value = "CODIGO";
                oWs.Cells[1, 9].Value = "OBSERVACIONES";
                oWs.Cells[1, 10].Value = "DIRECCION";
                oWs.Cells[1, 11].Value = "MARCA";
                oWs.Cells[1, 12].Value = "BLOCK";
                oWs.Cells[1, 13].Value = "NOMBRE DEL LECTURISTA";
                oWs.Cells[1, 14].Value = "FECHA DE LECTURA";
                oWs.Cells[1, 15].Value = "CONFIRMACION DE LECTURA";
                oWs.Cells[1, 16].Value = "FOTO";

                foreach (EnvioTrabajoCliLecReLec oBj in _lista)
                {
                    oWs.Cells[_fila, 1].Value = oBj.corre;
                    oWs.Cells[_fila, 2].Value = (oBj.lecSeccionLectura);
                    oWs.Cells[_fila, 3].Value = (oBj.zonaLectura);
                    oWs.Cells[_fila, 4].Value = (oBj.lecCorrelativo);
                    oWs.Cells[_fila, 5].Value = (oBj.suministroLectura);
                    oWs.Cells[_fila, 6].Value = (oBj.medidorLectura);
                    oWs.Cells[_fila, 7].Value = (oBj.lecMovil);
                    oWs.Cells[_fila, 8].Value = (oBj.idlecObs);
                    oWs.Cells[_fila, 9].Value = (oBj.obsLectura);
                    oWs.Cells[_fila, 10].Value = (oBj.dirLectura);
                    oWs.Cells[_fila, 11].Value = (oBj.lecMarcaMedidor);
                    oWs.Cells[_fila, 12].Value = (oBj.blockLectura);
                    oWs.Cells[_fila, 13].Value = (oBj.ope_nombre);
                    oWs.Cells[_fila, 14].Value = (oBj.fecLectura);
                    oWs.Cells[_fila, 15].Value = (oBj.confirLectura);
                    oWs.Cells[_fila, 16].Value = (oBj.foto);                   
                    _fila++;
                }

                oWs.Row(1).Style.WrapText = true;
                oWs.Row(1).Style.Font.Bold = true;
                oWs.Row(1).Style.HorizontalAlignment = Style.ExcelHorizontalAlignment.Center;
                oWs.Row(1).Style.VerticalAlignment = Style.ExcelVerticalAlignment.Center;
                oWs.Column(1).Style.Font.Bold = true;
                oWs.Column(1).AutoFit();
                oWs.Column(2).AutoFit();
                oWs.Column(3).AutoFit();
                oWs.Column(4).AutoFit();
                oWs.Column(5).AutoFit();
                oWs.Column(6).AutoFit();
                oWs.Column(7).AutoFit();
                oWs.Column(8).AutoFit();
                oWs.Column(9).AutoFit();
                oWs.Column(10).AutoFit();
                oWs.Column(11).AutoFit();
                oWs.Column(12).AutoFit();
                oWs.Column(13).AutoFit();
                oWs.Column(14).AutoFit();
                oWs.Column(15).AutoFit();
                oWs.Column(16).AutoFit();
                oEx.Save();
            }

            return new ContentResult
            {
                Content = "{ \"__a\": \"" + _servidor + "\" }",
                ContentType = "application/json"
            };
        }


        /// <summary>
        /// Autor: rcontreras
        /// Fecha: 29-09-2015
        /// </summary>
        /// <param name="__a"></param>
        /// <returns></returns>
        [HttpPost]
        public ActionResult DescargaPendientes(string __a)
        {
            Int32 _fila = 2;
            String _servidor;
            String _ruta;

            if (__a.Length == 0)
            {
                return View();
            }

            List<EnvioTrabajoCliLecReLecLecturaPendiente> _lista = MvcApplication._Deserialize<List<EnvioTrabajoCliLecReLecLecturaPendiente>>(__a);

            _servidor = String.Format("{0:ddMMyyyy_hhmmss}.xlsx", DateTime.Now);
            _ruta = Path.Combine(Server.MapPath("/Lecturas/Temp"), _servidor);

            FileInfo _file = new FileInfo(_ruta);
            if (_file.Exists)
            {
                _file.Delete();
                _file = new FileInfo(_ruta);
            }

            using (Excel.ExcelPackage oEx = new Excel.ExcelPackage(_file))
            {
                Excel.ExcelWorksheet oWs = oEx.Workbook.Worksheets.Add("Lecturas Pendientes");

                oWs.Cells[1, 1].Value = "OPERARIO";
                oWs.Cells[1, 2].Value = "ESTADO";
                oWs.Cells[1, 3].Value = "CANTIDAD";

                foreach (EnvioTrabajoCliLecReLecLecturaPendiente oBj in _lista)
                {
                    oWs.Cells[_fila, 1].Value = oBj.nomOpe;
                    oWs.Cells[_fila, 2].Value = oBj.nomEstado;
                    oWs.Cells[_fila, 3].Value = (oBj.cantidad);

                    _fila++;
                }

                oWs.Row(1).Style.WrapText = true;
                oWs.Row(1).Style.Font.Bold = true;
                oWs.Row(1).Style.HorizontalAlignment = Style.ExcelHorizontalAlignment.Center;
                oWs.Row(1).Style.VerticalAlignment = Style.ExcelVerticalAlignment.Center;

                oWs.Column(1).Style.Font.Bold = true;
                oWs.Column(1).AutoFit();
                oWs.Column(2).AutoFit();
                oWs.Column(3).AutoFit();

                oEx.Save();
            }

            return new ContentResult
            {
                Content = "{ \"__a\": \"" + _servidor + "\" }",
                ContentType = "application/json"
            };
        }

    }
}
