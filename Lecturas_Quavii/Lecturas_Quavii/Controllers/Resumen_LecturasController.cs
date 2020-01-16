using DSIGE.Modelo;
using DSIGE.Negocio;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using Excel = OfficeOpenXml;
using Style = OfficeOpenXml.Style;

namespace DSIGE.Web.Controllers
{
    public class Resumen_LecturasController : Controller
    {
        //
        // GET: /Resumen_Lecturas/

        public ActionResult Inicio()
        {
            return View();
        }


        public ActionResult ReporteProduccionDiario()
        {
            return View();
        }


        public ActionResult Inicio_new()
        {
            return View();
        }
        /// <summary>
        /// Autor: rcontreras
        /// Fecha: 19-11-2015
        /// </summary>
        /// <returns></returns>
        /// 

        public static string _Serialize(object value, bool ignore = false)
        {
            var SerializerSettings = new JsonSerializerSettings()
            {
                MaxDepth = Int32.MaxValue,
                NullValueHandling = (ignore == true ? NullValueHandling.Ignore : NullValueHandling.Include)
            };
            return JsonConvert.SerializeObject(value, Formatting.Indented, SerializerSettings);
        }


        [HttpPost]
        public string MostrandoFotos(int id_lectura, int idservicio)
        {
            object loDatos;
            try
            {
                NLectura obj_negocio = new NLectura();
                loDatos = obj_negocio.Capa_Negocio_Get_MostrarFotos(id_lectura, idservicio);
                return _Serialize(loDatos, true);
            }
            catch (Exception ex)
            {
                return _Serialize(ex.Message, true);
            }
        }


        [HttpPost]
        public string ListandoAgrupadoLecturas(int id_local, int id_servicio, string fechaini, string fechaFin)
        {
            object loDatos;
            try
            {
                NLectura obj_negocio = new NLectura();
                loDatos = obj_negocio.NListaLecturaAgrupado(id_local, id_servicio, fechaini, fechaFin);
                return _Serialize(loDatos, true);
            }
            catch (Exception ex)
            {
                return _Serialize(ex.Message, true);
            }
        }


        [HttpPost]
        public string ListadoReporteDiario(int cliente, string fechaini, string fechafin, int idServicio)
        {
            object loDatos;
            try
            {
                NLectura obj_negocio = new NLectura();
                loDatos = obj_negocio.Capa_Negocio_Get_ListadoReporteDiario(cliente, fechaini, fechafin, idServicio);
                return _Serialize(loDatos, true);
            }
            catch (Exception ex)
            {
                return _Serialize(ex.Message, true);
            }
        }



        [HttpPost]
        public string ListandoAgrupadoLecturas_new(int id_local, int id_servicio, string fechaini, string fechaFin, int id_supervisor, int id_operario_supervisor)
        {
            object loDatos;
            try
            {
                NLectura obj_negocio = new NLectura();
                loDatos = obj_negocio.NListaLecturaAgrupado_new(id_local, id_servicio, fechaini, fechaFin, id_supervisor, id_operario_supervisor);
                return _Serialize(loDatos, true);
            }
            catch (Exception ex)
            {
                return _Serialize(ex.Message, true);
            }
        }


        [HttpPost]
        public string ListandoDetalladoLecturasEstado(int id_local, int id_servicio, string fechaini, string fechaFin, int ipOperario, int categoriaLectura)
        {
            object loDatos;
            try
            {
                NLectura obj_negocio = new NLectura();
                loDatos = obj_negocio.NListaLecturaDetalladoEstado(id_local, id_servicio, fechaini, fechaFin, ipOperario, categoriaLectura);
                return _Serialize(loDatos, true);
            }
            catch (Exception ex)
            {
                return _Serialize(ex.Message, true);
            }
        }


        [HttpPost]
        public string ListandoDetalladoLecturas(int id_local, int id_servicio, string fechaini, string fechaFin, int ipOperario)
        {
            object loDatos;
            try
            {
                NLectura obj_negocio = new NLectura();
                loDatos = obj_negocio.NListaLecturaDetallado(id_local, id_servicio, fechaini, fechaFin, ipOperario);
                return _Serialize(loDatos, true);
            }
            catch (Exception ex)
            {
                return _Serialize(ex.Message, true);
            }
        }


        [HttpPost]
        public string MostrandoFotosLecturas(int id_lectura)
        {
            object loDatos;
            try
            {
                NLectura obj_negocio = new NLectura();
                loDatos = obj_negocio.Capa_Negocio_Get_MostrarFotosLecturas(id_lectura);
                return _Serialize(loDatos, true);
            }
            catch (Exception ex)
            {
                return _Serialize(ex.Message, true);
            }
        }


        [HttpPost]
        public ActionResult JsonListaLecturasResumen(string __a, string __c, string __d, string __e)
        {
            return new ContentResult
            {
                Content = MvcApplication._Serialize(
                        new NLectura().NListaLecturasResumen(
                            new Request_Lectura_Historico()
                            {
                                local = Convert.ToInt32(__a),
                                f_ini = Convert.ToString(__c),
                                f_fin = Convert.ToString(__d),
                                lista = Convert.ToString(__e)                                 
                            }
                        )
                    , true),
                ContentType = "application/json"
            };
        }


        /// <summary>
        /// Autor: rcontreras
        /// Fecha: 24-11-2015
        /// </summary>
        /// <param name="__a"></param>
        /// <returns></returns>
        [HttpPost]
        public ActionResult DescargaExcel(string __a)
        {
            Int32 _fila = 2;
            String _servidor;
            String _ruta;

            if (__a.Length == 0)
            {
                return View();
            }

            List<LecturaHistorico> _lista = MvcApplication._Deserialize<List<LecturaHistorico>>(__a);

            _servidor = String.Format("{0:ddMMyyyy_hhmmss}.xlsx", DateTime.Now);
            //_ruta = Path.Combine(Server.MapPath("/Lecturas/Temp"), _servidor);
            _ruta = Path.Combine(Server.MapPath("/Temp"), _servidor);

            FileInfo _file = new FileInfo(_ruta);
            if (_file.Exists)
            {
                _file.Delete();
                _file = new FileInfo(_ruta);
            }

            using (Excel.ExcelPackage oEx = new Excel.ExcelPackage(_file))
            {
                Excel.ExcelWorksheet oWs = oEx.Workbook.Worksheets.Add("Resumen de Lecturas");

                oWs.Cells[1, 1].Value = "OPERADOR";
                oWs.Cells[1, 2].Value = "TOTAL";
                oWs.Cells[1, 3].Value = "REALIZADOS";
                oWs.Cells[1, 4].Value = "CON FOTO";
                oWs.Cells[1, 5].Value = "PENDIENTES";
                oWs.Cells[1, 6].Value = "% AVANCE";
                oWs.Cells[1, 7].Value = "HORA INICIO DE TRABAJO";
                oWs.Cells[1, 8].Value = "HORA TERMINO DE TRABAJO";
                oWs.Cells[1, 9].Value = "HORAS TRABAJADAS";

                foreach (LecturaHistorico oBj in _lista)
                {
                    oWs.Cells[_fila, 1].Value = oBj.des_ope;
                    oWs.Cells[_fila, 2].Value = oBj.total;
                    oWs.Cells[_fila, 3].Value = (oBj.realizado);
                    oWs.Cells[_fila, 4].Value = (oBj.conFoto);
                    oWs.Cells[_fila, 5].Value = (oBj.pendiente);
                    oWs.Cells[_fila, 6].Value = (oBj.avance);
                    oWs.Cells[_fila, 7].Value = (oBj.f_ini);
                    oWs.Cells[_fila, 8].Value = (oBj.f_fin);
                    oWs.Cells[_fila, 9].Value = (oBj.horas);

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

                oEx.Save();
            }

            return new ContentResult
            {
                Content = "{ \"__a\": \"" + _servidor + "\" }",
                ContentType = "application/json"
            };
        }



        [HttpPost]
        public string Reporte_DetalleLecturas(int id_local, int id_servicio, string fechaini, string fechaFin, int ipOperario, int categoriaLectura)
        {

            string Res = "";
            string _servidor;
            string _ruta;
            string resultado = "";
            int _fila = 2;
            string FileRuta = "";
            string FileExcel = "";


            List<ResumenLecturas> loDatos = new List<ResumenLecturas>();

            try
            {
                NLectura obj_negocio = new NLectura();
                loDatos = obj_negocio.NListaLecturaDetalladoEstado(id_local, id_servicio, fechaini, fechaFin, ipOperario, categoriaLectura);

                if (loDatos.Count == 0)
                {
                    return _Serialize("0|No hay informacion para mostrar.", true);
                }

                _servidor = String.Format("{0:ddMMyyyy_hhmmss}.xlsx", DateTime.Now);
                FileRuta = System.Web.Hosting.HostingEnvironment.MapPath("~/Upload/Detalle_Resumen_Lecturas_" + _servidor);
                string rutaServer = ConfigurationManager.AppSettings["servidor-archivos"];


                FileExcel = rutaServer + "Detalle_Resumen_Lecturas_" + _servidor;
                FileInfo _file = new FileInfo(FileRuta);
                if (_file.Exists)
                {
                    _file.Delete();
                    _file = new FileInfo(FileRuta);
                }

                using (Excel.ExcelPackage oEx = new Excel.ExcelPackage(_file))
                {
                    Excel.ExcelWorksheet oWs = oEx.Workbook.Worksheets.Add("ControlProcesos");
                    oWs.Cells.Style.Font.SetFromFont(new Font("Tahoma", 9));

                    oWs.Cells[1, 1].Value = "Lecturista";
                    oWs.Cells[1, 2].Value = "Direccion";
                    oWs.Cells[1, 3].Value = "Suministro";
                    oWs.Cells[1, 4].Value = "Medidor";
                    oWs.Cells[1, 5].Value = "Zona";
                    oWs.Cells[1, 6].Value = "Lectura";
                    oWs.Cells[1, 7].Value = "Confirmacion Lectura";
                    oWs.Cells[1, 8].Value = "Observacion";
                    oWs.Cells[1, 9].Value = "Estado";


                    //marco detalle CABECERA
                    for (int i = 1; i <= 11; i++)
                    {
                        oWs.Cells[1, i].Style.Font.Size = 9;
                        oWs.Cells[1, i].Style.Font.Bold = true;
                        oWs.Cells[1, i].Style.Border.BorderAround(OfficeOpenXml.Style.ExcelBorderStyle.Thin);    //marco
                        oWs.Cells[1, i].Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Center;
                    }

                    foreach (ResumenLecturas obj in loDatos)
                    {
                        //marco detalle   
                        for (int i = 1; i <= 9; i++)
                        {
                            oWs.Cells[_fila, i].Style.Border.BorderAround(OfficeOpenXml.Style.ExcelBorderStyle.Thin);    //marco  
                        }

                        oWs.Cells[_fila, 1].Value = obj.id_Operario_Lectura;
                        oWs.Cells[_fila, 2].Value = obj.direccion_lectura;
                        oWs.Cells[_fila, 3].Value = obj.suministro_lectura;
                        oWs.Cells[_fila, 4].Value = obj.medidor_lectura;
                        oWs.Cells[_fila, 5].Value = obj.Zona_lectura;
                        oWs.Cells[_fila, 6].Value = obj.LecturaMovil_Lectura;
                        oWs.Cells[_fila, 7].Value = obj.confirmacion_Lectura;
                        oWs.Cells[_fila, 8].Value = obj.abreviatura_observacion;
                        oWs.Cells[_fila, 9].Value = obj.descripcion_estado;

                        _fila++;
                    }

                    oWs.Cells.Style.Font.Size = 8; //letra tamaño  
                    oWs.Row(1).Style.Font.Bold = true;
                    oWs.Row(1).Style.HorizontalAlignment = Style.ExcelHorizontalAlignment.Center;
                    oWs.Row(1).Style.VerticalAlignment = Style.ExcelVerticalAlignment.Center;

                    for (int i = 1; i <= 9; i++)
                    {
                        oWs.Column(i).AutoFit();
                    }

                    oEx.Save();
                }
                return _Serialize("1|" + FileExcel, true);
            }
            catch (Exception ex)
            {
                return _Serialize("0|" + ex.Message, true);
            }
        }



    }
}

