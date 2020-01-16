using DSIGE.Modelo;
using DSIGE.Negocio;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Web;
using System.Web.Mvc;

using Excel = OfficeOpenXml;
using Style = OfficeOpenXml.Style;

namespace DSIGE.Web.Controllers
{
    public class Asigna_LecturaController : Controller
    {
        //
        // GET: /Asigna_Lectura/

        public ActionResult Inicio()
        {
            return View();
        }

        /// <summary>
        /// Autor: rcontreras
        /// Fecha: 17-05-2015
        /// </summary>
        /// <returns></returns>
        [HttpPost]
        public ActionResult JsonListaAsigLecturaReLectura(string __a, string __b, string __c, string __d, string __e, string __f, string __g)
        {
            return new ContentResult
            {
                Content = MvcApplication._Serialize(
                        new NAsignaLecturaReLectura().NLista(
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
        /// Fecha: 18-09-2015
        /// </summary>
        /// <returns></returns>
        [HttpPost]
        public ActionResult ActualizaEnvioMovil(string __a, string __b)
        {
            Int32 respuesta = new NAsignaLecturaReLectura().NActualizaEnviaMovil(
                        new Request_actualiza_app_movil_lec_relec()
                        {
                            idLectura = __a,
                            fechActu = Convert.ToDateTime(__b),
                            usuEdi = ((Sesion)Session["Session_Usuario_Acceso"]).usuario.usu_id,
                            idEmp = ((Sesion)Session["Session_Usuario_Acceso"]).empresa.emp_id
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
        /// Fecha: 19-09-2015
        /// </summary>
        /// <returns></returns>
        [HttpGet]
        public ActionResult JsonListaEstados(string __a)
        {
            return new ContentResult
            {
                Content = MvcApplication._Serialize(new NAsignaLecturaReLectura().NListaEstados(Convert.ToInt32(__a))),
                ContentType = "application/json"
            };
        }

        /// <summary>
        /// Autor: rcontreras
        /// Fecha: 19-09-2015
        /// </summary>
        /// <returns></returns>
        [HttpPost]
        public ActionResult ListaTecnicos(string __a, string __b)
        {
            return View(new NTecnico().NLista(
                new Request_Operario_Empresa_Local_Opcion(){
                    loc_id = Convert.ToInt32(__a),
                    tip_serv = Convert.ToInt32(__b),
                    emp_id = ((Sesion)Session["Session_Usuario_Acceso"]).empresa.emp_id
                }));
        }

        /// <summary>
        /// Autor: rcontreras
        /// Fecha: 19-09-2015
        /// </summary>
        /// <returns></returns>
        [HttpPost]
        public ActionResult ListaOperarios(string __a, string __b)
        {
            return View(new NTecnico().NLista(
                new Request_Operario_Empresa_Local_Opcion()
                {
                    loc_id = Convert.ToInt32(__a),
                    tip_serv = Convert.ToInt32(__b),
                    emp_id = ((Sesion)Session["Session_Usuario_Acceso"]).empresa.emp_id
                }));
        }

        /// <summary>
        /// Autor: rcontreras
        /// Fecha: 19-09-2015
        /// </summary>
        /// <returns></returns>
        [HttpPost]
        public ActionResult JsonListaHistorico(string __a)
        {
            return new ContentResult
            {
                Content = MvcApplication._Serialize(
                        new NAsignaLecturaReLectura().NListaHistorico(__a)
                    , true),
                ContentType = "application/json"
            };
        }

        /// <summary>
        /// Autor: rcontreras
        /// Fecha: 19-09-2015
        /// </summary>
        /// <param name="__a"></param>
        /// <returns></returns>
        [HttpPost]
        public ActionResult DescargaHistorico(string __a)
        {
            Int32 _fila = 2;
            String _servidor;
            String _ruta;

            if (__a.Length == 0)
            {
                return View();
            }

            List<AsignaLecturaReLectura> _lista = MvcApplication._Deserialize<List<AsignaLecturaReLectura>>(__a);

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
                Excel.ExcelWorksheet oWs = oEx.Workbook.Worksheets.Add("Historico de Lecturas");

                oWs.Cells[1, 1].Value = "MEDIDOR";
                oWs.Cells[1, 2].Value = "FECHA DE LECTURA";
                oWs.Cells[1, 3].Value = "LECTURA";
                oWs.Cells[1, 4].Value = "OPERARIO";
                oWs.Cells[1, 5].Value = "OBSERVACION";
                oWs.Cells[1, 6].Value = "ESTADO";

                foreach (AsignaLecturaReLectura oBj in _lista)
                {
                    oWs.Cells[_fila, 1].Value = oBj.medidorLectura;
                    oWs.Cells[_fila, 2].Value = oBj.fecLectura;
                    oWs.Cells[_fila, 3].Value = (oBj.lectura);
                    oWs.Cells[_fila, 4].Value = (oBj.ope_nombre);
                    oWs.Cells[_fila, 5].Value = (oBj.obsLectura);
                    oWs.Cells[_fila, 6].Value = (oBj.Estado);

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
        /// Fecha: 19-09-2015
        /// </summary>
        /// <returns></returns>
        [HttpPost]
        public ActionResult ActualizaReasignaOperador(string __a, string __b, string __c, string __d)
        {
            Int32 respuesta = new NAsignaLecturaReLectura().NActualizaReasignaOperador(
                        new Request_actualiza_app_movil_lec_relec()
                        {
                            idLectura = __a,
                            fechActu = Convert.ToDateTime(__b),
                            idOperador = Convert.ToInt32(__c),
                            usuEdi = ((Sesion)Session["Session_Usuario_Acceso"]).usuario.usu_id,
                            idEmp = ((Sesion)Session["Session_Usuario_Acceso"]).empresa.emp_id,
                            opcion = Convert.ToInt32(__d)
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
