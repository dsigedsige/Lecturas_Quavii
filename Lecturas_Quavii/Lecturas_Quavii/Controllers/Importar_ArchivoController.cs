using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Web;
using System.Web.Mvc;

using Excel = OfficeOpenXml;
using Style = OfficeOpenXml.Style;

using DSIGE;
using DSIGE.Modelo;
using DSIGE.Negocio;
using System.Data.OleDb;
using System.Data;
using Newtonsoft.Json;

namespace DSIGE.Web.Controllers
{
    public class Importar_ArchivoController : Controller
    {
        //
        // GET: /Importar_Archivo/

        private const string _contraseña = "dsigev1.0";

        public ActionResult Inicio()
        {
            return View();
        }

        /// <summary>
        /// Autor: jlucero
        /// Fecha: 2015-07-02
        /// </summary>
        /// <param name="__a"></param>
        /// <returns></returns>
        [HttpPost]
        public ActionResult ImportaArchivo(HttpPostedFileBase __a, string __b)
        {
            int _fila = 0;
            int _cantidad = 0;
            bool _estado = true;
            string _valor = "";
            string _local = "";
            string _server = "";
            string _extension = "";
            int ser_id = 0;
            string ser_descripcion = "";
            DateTime _fecha_actual = DateTime.Now;
            DateTime _valida_fecha;
            Int32 _valida_entero;

            List<Lectura> oRp = new List<Lectura>();

            if (__a != null)
            {
                #region Inserta informacion

                _extension = __a.FileName.Split('.')[1];
                //_server = String.Format("{0:ddMMyyyy_hhmmss}." + _extension, DateTime.Now);
                _server = String.Format("{0:ddMMyyyy_hhmmss}.xlsx", _fecha_actual);
                _local = Path.Combine(Server.MapPath("/Lecturas/Temp/Importacion"), _server);
                //_local = Path.Combine(Server.MapPath("/Temp/Importacion"), _server);
                __a.SaveAs(_local);

                FileInfo oFile = new FileInfo(_local);
                using (Excel.ExcelPackage oEx = new Excel.ExcelPackage(oFile))
                {
                    foreach (Excel.ExcelWorksheet oWs in oEx.Workbook.Worksheets)
                    {
                        ser_descripcion = oWs.Name.ToUpper();
                        ser_id = (ser_descripcion == "LECTURA" ? 1 : (ser_descripcion == "RELECTURA" ? 2 : (ser_descripcion == "CORTE" ? 3 : (ser_descripcion == "RECONEXION" ? 4 : 0))));

                        #region Servicio lectura y/o relectura

                        if (ser_descripcion == "LECTURA" || ser_descripcion == "RELECTURA")
                        {
                            List<Lectura> _oRq = new List<Lectura>();

                            _fila = 2;
                            _estado = true;

                            do
                            {
                                _valor = oWs.Cells[_fila, 1].Text;
                                if (_valor.Length != 0)
                                {

                                    _oRq.Add(
                                            new Lectura()
                                            {

                                                lec_importacion_archivo = _server,
                                                //lec_numero_instalacion = oWs.Cells[_fila, 2].Text,
                                                //lec_medidor = oWs.Cells[_fila, 4].Text,
                                                //lec_direccion = oWs.Cells[_fila, 5].Text + " " + oWs.Cells[_fila, 6].Text,
                                                //lec_distrito = oWs.Cells[_fila, 7].Text,
                                                //lec_anterior = oWs.Cells[_fila, 9].Text,
                                                //lec_interlocutor = oWs.Cells[_fila, 13].Text,
                                                //lec_suministro = oWs.Cells[_fila, 14].Text,
                                                //lec_cliente_tipo = oWs.Cells[_fila, 15].Text,
                                                //lec_categoria = oWs.Cells[_fila, 16].Text,
                                                //lec_zona = oWs.Cells[_fila, 17].Text,
                                                //lec_medidor_marca = oWs.Cells[_fila, 18].Text,
                                                //lec_cliente_nombre = oWs.Cells[_fila, 19].Text,
                                                //ope_id = Convert.ToInt32(oWs.Cells[_fila, 20].Text.Length == 0 ? "0" : oWs.Cells[_fila, 20].Text),
                                                //lec_orden = Convert.ToInt32(oWs.Cells[_fila, 21].Text.Length == 0 ? "0" : oWs.Cells[_fila, 21].Text),

                                                lec_suministro = oWs.Cells[_fila, 1].Text,
                                                lec_direccion = oWs.Cells[_fila, 2].Text,
                                                lec_seccion = Convert.ToInt32(oWs.Cells[_fila, 3].Text),
                                                lec_zona = oWs.Cells[_fila, 4].Text,
                                                lec_correlativo = Convert.ToInt32(oWs.Cells[_fila, 5].Text),
                                                lec_medidor = oWs.Cells[_fila, 6].Text,
                                                lec_medidor_marca = oWs.Cells[_fila, 7].Text,
                                                lec_anterior = oWs.Cells[_fila, 8].Text,
                                                dniOpe = oWs.Cells[_fila, 9].Text,
                                                bloque = oWs.Cells[_fila, 10].Text,
                                                lec_orden = Convert.ToInt32(oWs.Cells[_fila, 11].Text),
                                                loc_id = Convert.ToInt32(__b),
                                                ser_id = ser_id,
                                                emp_id = ((Sesion)Session["Session_Usuario_Acceso"]).empresa.emp_id
                                            }
                                        );

                                    _fila++;
                                }
                                else
                                {
                                    _estado = false;
                                }
                            } while (_estado == true);

                            _cantidad = new NLectura().NBulkInfo(_oRq);

                            break;
                        }

                        #endregion

                        #region Servicio corte y/o reconexion

                        if (ser_descripcion == "CORTE" || ser_descripcion == "RECONEXION")
                        {
                            List<Corte> _oRq = new List<Corte>();

                            _fila = 2;
                            _estado = true;

                            do
                            {
                                _valor = oWs.Cells[_fila, 24].Text;
                                if (_valor.Length != 0)
                                {
                                    _oRq.Add(
                                            new Corte()
                                            {
                                                emp_id = ((Sesion)Session["Session_Usuario_Acceso"]).empresa.emp_id,
                                                ser_id = ser_id,
                                                cor_importacion_fecha = _fecha_actual,
                                                cor_importacion_archivo = _server,
                                                loc_id = Convert.ToInt32(__b),
                                                cor_aviso_clase = oWs.Cells[_fila, 1].Text,
                                                cor_aviso = oWs.Cells[_fila, 2].Text,
                                                cor_aviso_fecha = (DateTime.TryParse(oWs.Cells[_fila, 3].Text, out _valida_fecha) ? String.Format("{0:yyyy-MM-dd}", _valida_fecha) : ""), //Convert.ToDateTime(oWs.Cells[_fila, 3].Text)
                                                cor_documento_bloqueo = oWs.Cells[_fila, 4].Text,
                                                cor_suministro = oWs.Cells[_fila, 5].Text,
                                                cor_interlocutor = oWs.Cells[_fila, 6].Text,
                                                cor_clase_cuenta = oWs.Cells[_fila, 7].Text,
                                                cor_deuda = oWs.Cells[_fila, 8].Text,
                                                cor_cantidad = oWs.Cells[_fila, 9].Text,
                                                cor_instalacion_numero = oWs.Cells[_fila, 10].Text,
                                                cor_medidor = oWs.Cells[_fila, 11].Text,
                                                cor_direccion = oWs.Cells[_fila, 12].Text,
                                                cor_distrito = oWs.Cells[_fila, 13].Text,
                                                cor_seccion = oWs.Cells[_fila, 14].Text,
                                                ope_id = (Int32.TryParse(oWs.Cells[_fila, 24].Text, out _valida_entero) ? _valida_entero : 0),
                                                cor_orden = (Int32.TryParse(oWs.Cells[_fila, 25].Text, out _valida_entero) ? _valida_entero : 0)
                                            }
                                        );

                                    _fila++;
                                }
                                else
                                {
                                    _estado = false;
                                }
                            } while (_estado == true);

                            _cantidad = new NCorte().NBulkInfo(_oRq);

                            break;
                        }

                        #endregion
                    }
                }

                #endregion
            }

            return new ContentResult
            {
                Content = "{ \"__a\" : \"" + _server + "\", \"__b\" : " + _cantidad + ", \"__c\": { \"__a\": " + ser_id + ", \"__b\": \"" + ser_descripcion + "\" } }",
                ContentType = "application/json"
            };
        }


        /// <summary>
        /// Autor: jlucero
        /// Fecha: 2015-07-07
        /// </summary>
        /// <param name="__a"></param>
        /// <returns></returns>
        [HttpGet]
        public ActionResult ValidaInformacion(string __a, string __b) {
            /* Leyenda de servicios:
             * 1 -> Lectura
             * 2 -> Relectura
             * 3 -> Corte
             * 4 -> Reconexion
             * 
             * _v = lec_estado
             * _e = lec_suministro
             * _f = lec_medidor
             * _i = lec_direccion
             * _l = lec_cliente_nombre
             * clase = clase
             * _t = cor_orden
             */

            Int32 servicio = Convert.ToInt32(__b);

            List<Lectura> oLs = new List<Lectura>();

            if (servicio == 1 || servicio == 2) {
                oLs = new NLectura().NListaValidado(
                        new Request_Lectura_Validacion()
                        {
                            archivo = __a
                        }
                    );
            }else if(servicio == 3 || servicio == 4){
                foreach (Corte oBj in new NCorte().NListaValidado(
                        new Request_Corte_Validacion()
                        {
                            archivo = __a
                        }
                    ))
                {
                    oLs.Add(
                            new Lectura() { 
                                lec_estado = oBj.cor_estado,
                                lec_suministro = oBj.cor_suministro,
                                lec_medidor = oBj.cor_medidor,
                                lec_direccion = oBj.cor_direccion,
                                lec_cliente_nombre = oBj.cor_interlocutor,
                                clase = oBj.clase,
                                lec_orden = oBj.cor_orden
                            }
                        );
                }
            }

            return new ContentResult
            {
                Content = MvcApplication._Serialize(oLs),
                ContentType = "application/json"
            };
        }

        /// <summary>
        /// Autor: jlucero
        /// Fecha: 2015-07-08
        /// </summary>
        /// <param name="__a"></param>
        /// <param name="__b"></param>
        /// <param name="__c"></param>
        /// <param name="__d"></param>
        /// <returns></returns>
        [HttpPost]
        public ActionResult Inserta(string __a, string __b, string __c, string __d)
        {
            Int32 resultado = 0;

            if (Convert.ToInt32(__d) == 1 || Convert.ToInt32(__d) == 2)
            {
                resultado = new NLectura().NInsertaMasivo(
                    new Request_Lectura_Inserta_Masivo()
                    {
                        lec_archivo = __a,
                        lec_asignacion_fecha = __b,
                        lec_id = __c,
                        usu_id = ((Sesion)Session["Session_Usuario_Acceso"]).usuario.usu_id
                    }
                );
            }
            else if (Convert.ToInt32(__d) == 3 || Convert.ToInt32(__d) == 4)
            {
                resultado = new NCorte().NInsertaMasivo(
                    new Request_Corte_Inserta_Masivo()
                    {
                        cor_archivo = __a,
                        cor_asignacion_fecha = __b,
                        cor_id = __c,
                        usu_id = ((Sesion)Session["Session_Usuario_Acceso"]).usuario.usu_id
                    }
                );
            }

            return new ContentResult
            {
                Content = "{ \"_a\": " + resultado + " }",
                ContentType = "application/json"
            };
        }


        [HttpPost]
        public ActionResult Importar(HttpPostedFileBase file, int idlocal)
        {
            DateTime _fecha_actual = DateTime.Now;
            try
            {



                string extension = System.IO.Path.GetExtension(file.FileName);
                string nomExcel = ((Sesion)Session["Session_Usuario_Acceso"]).usuario.usu_id + extension;
                string fileLocation = Server.MapPath("~/Upload") + "\\" + nomExcel;

                file.SaveAs(fileLocation);



                DataTable dt = new DataTable();
                List<Lectura> listaLecturas = new List<Lectura>();
                Lectura objLectura;

                dt = new NLectura().ListaExcel(fileLocation);

                foreach (DataRow row in dt.Rows)
                {
                    objLectura = new Lectura();
                    objLectura.lec_importacion_archivo = nomExcel;
                    objLectura.lec_suministro =  row["CLIENTE"].ToString();
                    objLectura.lec_direccion = row["DIRECCION"].ToString();
                    objLectura.lec_seccion = Convert.ToInt32(row["Sc"].ToString());
                    objLectura.lec_zona = row["Zn"].ToString();
                    objLectura.lec_correlativo = Convert.ToInt32(row["CRR"].ToString());
                    objLectura.lec_medidor = row["MEDIDOR"].ToString();
                    objLectura.lec_medidor_marca = row["MARCA"].ToString();
                    objLectura.MesUno = row["Mes Uno"].ToString();
                    objLectura.MesDos = row["Mes Dos"].ToString();
                    objLectura.MesTres = row["Mes Tres"].ToString();
                    objLectura.MesCuatro = row["Mes Cuatro"].ToString();
                    objLectura.MesCinco = row["Mes Cinco"].ToString();
                    objLectura.dniOpe = row["DNI"].ToString();
                    objLectura.bloque = row["Bloque"].ToString();
                    objLectura.lec_orden = Convert.ToInt32(row["Orden"].ToString());
                    objLectura.loc_id = idlocal;
                    objLectura.id_UsuarioExpor = ((Sesion)Session["Session_Usuario_Acceso"]).usuario.usu_id;
                    listaLecturas.Add(objLectura);
                                                                                                                                                                                                                                                                                                
                }
                return new ContentResult
                {
                    Content = MvcApplication._Serialize(new NLectura().NImportarArchivoExcel(listaLecturas)),
                    ContentType = "application/json"
                };
            }
            catch (Exception e)
            {
                return new ContentResult
                {
                    Content = MvcApplication._Serialize(e.Message),
                    ContentType = "application/json"
                };
            }

        }


        

        [HttpPost]
        public string Lecturas_Detallado(string SC, string DNI)
        {
            object obj_list_datos;
            try
            {
               Cls_Negocio_Importacion_Lecturas Objeto_Negocio = new Cls_Negocio_Importacion_Lecturas();
               List<Cls_Entidad_Importacion_Lecturas> objeto_List_Importacion_Lecturas = new List<Cls_Entidad_Importacion_Lecturas>();
                
               obj_list_datos=Objeto_Negocio.Capa_Negocio_Listar_Temporal_Lecturas_Detallado(((Sesion)Session["Session_Usuario_Acceso"]).usuario.usu_id,SC, DNI ) ;

               return _Serialize(obj_list_datos, true);

            }
            catch (Exception ex)
            {
                return _Serialize(ex.Message,true);
            }
                       

        }

               

        //[HttpPost]
        //public string JsonInsertarExel(HttpPostedFileBase file, int idlocal, string idfechaAsignacion)
        //{

        //    List<Cls_Entidad_Importacion_Lecturas> objeto_List_Importacion_Lecturas = new List<Cls_Entidad_Importacion_Lecturas>();
        //    DateTime _fecha_actual = DateTime.Now;

        //    object loDatos;
    
        //     try
        //    {

        //        string extension = System.IO.Path.GetExtension(file.FileName);
        //        string nomExcel = ((Sesion)Session["Session_Usuario_Acceso"]).usuario.usu_id + extension;
        //        string fileLocation = Server.MapPath("~/Upload") + "\\" + nomExcel;
        //        file.SaveAs(fileLocation);

                    

        //        NLectura obj_Lectura_Importacion_Bl = new NLectura();

        //        Cls_Negocio_Importacion_Lecturas Objeto_Negocio = new Cls_Negocio_Importacion_Lecturas();
        //        Objeto_Negocio.Capa_Negocio_ListaExcel(fileLocation, ((Sesion)Session["Session_Usuario_Acceso"]).usuario.usu_id, idlocal, idfechaAsignacion);

        //       loDatos = Objeto_Negocio.Capa_Negocio_Listar_Temporal_Lecturas_Agrupado(((Sesion)Session["Session_Usuario_Acceso"]).usuario.usu_id);
        //       return _Serialize(loDatos, true);
                
                 

        //      ////  Objeto_Negocio.Capa_Negocio_Eliminar_Tabla_Temporal(((Sesion)Session["Session_Usuario_Acceso"]).usuario.usu_id);  
                
        //      // // dt_detalle = obj_Lectura_Importacion_Bl.ListaExcel(fileLocation,);                
        //      //  //foreach (DataRow row in dt_detalle.Rows)
        //      //  //{
        //      //  //    Cls_Entidad_Importacion_Lecturas Entidad = new Cls_Entidad_Importacion_Lecturas();

        //      //  //    Entidad.CLIENTE = row[0].ToString();
        //      //  //    Entidad.DIRECCION = row[1].ToString();
        //      //  //    Entidad.Sc = row[2].ToString();
        //      //  //    Entidad.Zn = row[3].ToString();
        //      //  //    Entidad.CRR = row[4].ToString();
        //      //  //    Entidad.MEDIDOR = row[5].ToString();
        //      //  //    Entidad.MARCA = row[6].ToString();
        //      //  //    Entidad.MesUno = row[7].ToString();
        //      //  //    Entidad.MesDos = row[8].ToString();
        //      //  //    Entidad.MesTres = row[9].ToString();
        //      //  //    Entidad.MesCuatro = row[10].ToString();
        //      //  //    Entidad.MesCinco = row[11].ToString();
        //      //  //    Entidad.DNI = row[12].ToString();
        //      //  //    Entidad.Nombre_Operador = "";
        //      //  //    Entidad.Bloque = row[13].ToString();
        //      //  //    Entidad.Orden = Convert.ToInt32(row[14].ToString());

        //      //  //    Entidad.loc_id = idlocal;
        //      //  //    Entidad.id_UsuarioExpor = ((Sesion)Session["Session_Usuario_Acceso"]).usuario.usu_id;

        //      //  //    Entidad.Promedio = 0;
        //      //  //    Entidad.Minimo = 0;
        //      //  //    Entidad.Maximo = 0;
        //      //  //    Entidad.Lect_Minima = 0;
        //      //  //    Entidad.Lect_Maxima = 0;
        //      //  //    Entidad.Ultimas_Lect_3 =0;

        //      //  //    objeto_List_Importacion_Lecturas.Add(Entidad);

        //      //  //    //Guardando los datos en la table temporal

        //      //  // Objeto_Negocio.Capa_Negocio_ImportarArchivoExcel(Entidad);
        //      //  //}
                
               
        //    }
        //    catch (Exception ex)
        //    {
        //        return _Serialize(ex.Message, true);
        //    }

       
            
        //}



        [HttpPost]
        public string JsonInsertarExel(HttpPostedFileBase file, int idlocal, string idfechaAsignacion)
        {

            List<Cls_Entidad_Importacion_Lecturas> objeto_List_Importacion_Lecturas = new List<Cls_Entidad_Importacion_Lecturas>();
            DateTime _fecha_actual = DateTime.Now;

            object loDatos;

            try
            {

                string extension = System.IO.Path.GetExtension(file.FileName);
                string nomExcel = ((Sesion)Session["Session_Usuario_Acceso"]).usuario.usu_id + extension;
                string fileLocation = Server.MapPath("~/Upload") + "\\" + nomExcel;
                file.SaveAs(fileLocation);



                NLectura obj_Lectura_Importacion_Bl = new NLectura();

                Cls_Negocio_Importacion_Lecturas Objeto_Negocio = new Cls_Negocio_Importacion_Lecturas();
                Objeto_Negocio.Capa_Negocio_ListaExcel(fileLocation, ((Sesion)Session["Session_Usuario_Acceso"]).usuario.usu_id, idlocal, idfechaAsignacion);

                loDatos = Objeto_Negocio.Capa_Negocio_Listar_Temporal_Lecturas_Agrupado(((Sesion)Session["Session_Usuario_Acceso"]).usuario.usu_id);
                return _Serialize(loDatos, true);

            }
            catch (Exception ex)
            {
                return _Serialize(ex.Message, true);
            }



        }


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
        public string ListadoServicios()
        {

            object loDatos;
            try
            {
                Cls_Negocio_Importacion_Lecturas objeto_negocio = new Cls_Negocio_Importacion_Lecturas();

                loDatos = objeto_negocio.Capa_Negocio_PermisoListUsuarioServicio(((Sesion)Session["Session_Usuario_Acceso"]).usuario.usu_id); 
                

            }
            catch (Exception ex)
            {
                throw ex;
            }

            return _Serialize(loDatos, true);
        }



        [HttpPost]
        public string Eliminando_Tabla_Temporal()
        {

            bool loDatos;
            try
            {

                var usuario = ((Sesion)Session["Session_Usuario_Acceso"]).usuario.usu_id;

                Cls_Negocio_Importacion_Lecturas obj_negocio = new Cls_Negocio_Importacion_Lecturas();
                loDatos = obj_negocio.Capa_Negocio_Eliminar_Tabla_Temporal(usuario); 

            }
            catch (Exception ex)
            {
                throw ex;
            }
            return _Serialize(loDatos, true);

        }



        [HttpPost]
        public ActionResult JsonListarTablaTemp()
        {
            return new ContentResult
            {
                Content = MvcApplication._Serialize(new NLectura().NListarTablaTem(((Sesion)Session["Session_Usuario_Acceso"]).usuario.usu_id)),
                ContentType = "application/json"
            };
        }



        [HttpPost]
        public string JsonRegistrarTemp_Lecturas(string fechaAsignacion, int id_servicio, string nombre_archivo)
        {

            bool loDatos;
            try
            {
                var usuario = ((Sesion)Session["Session_Usuario_Acceso"]).usuario.usu_id;

               Cls_Negocio_Importacion_Lecturas obj_negocio = new Cls_Negocio_Importacion_Lecturas();
               loDatos = obj_negocio.Capa_Negocio_Guardar_Informacion(fechaAsignacion, id_servicio, nombre_archivo, usuario);

            }
            catch (Exception ex)
            {
                throw ex;
            }

            return _Serialize(loDatos, true);

        }
        

        //[HttpGet]
        //public ActionResult JsonRegistrarTemp_Lecturas(string fechaAsignacion, int id_servicio)
        //{
        //    try
        //    {
        //        return new ContentResult
        //        {
        //            Content = MvcApplication._Serialize(new NLectura().Guardar_TableTemp_Lecturas(((Sesion)Session["Session_Usuario_Acceso"]).usuario.usu_id, fechaAsignacion)),
        //            ContentType = "application/json"
        //        };
        //    }
        //    catch (Exception)
        //    {
        //        throw;
        //    }

        //}

        ///// <summary>
        ///// Autor: jlucero
        ///// Fecha: 2015-07-09
        ///// </summary>
        ///// <param name="__a"></param>
        ///// <param name="__b"></param>
        ///// <returns></returns>
        //[HttpPost]
        //public ActionResult Elimina(string __a, string __b)
        //{
        //    Int32 resultado = 0;

        //    if (Convert.ToInt32(__b) == 1 || Convert.ToInt32(__b) == 2)
        //    {
        //        resultado = new NLectura().NEliminaMasivo(
        //            new Request_Elimina_Masivo()
        //            {
        //                archivo = __a,
        //                opcion = Convert.ToInt32(__b)
        //            }
        //        );
        //    }

        //    return new ContentResult
        //    {
        //        Content = "{ \"_a\": " + resultado + " }",
        //        ContentType = "application/json"
        //    };
        //}
    }
}