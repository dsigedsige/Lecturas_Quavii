using DSIGE.Modelo;
using DSIGE.Negocio;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading;
using System.Web;
using System.Web.Mvc;

namespace Lecturas_Quavii.Controllers
{
    public class ImportarController : Controller
    {
        //
        // GET: /Importar/

        public ActionResult importar_index()
        {
            return View();
        }


        /// <summary>
        /// SERIALIZACION DE JSONPROPERTY d 
        /// </summary>
        /// <param name="value"></param>
        /// <param name="ignore"></param>
        /// <returns></returns>
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
        public string get_marcaMedidor()
        {
            object loDatos;
            try
            {
                Cls_Negocio_Importacion_Lecturas objeto_negocio = new Cls_Negocio_Importacion_Lecturas();
                loDatos = objeto_negocio.Capa_Negocio_get_marcaMedidor();

            }
            catch (Exception ex)
            {
                throw ex;
            }
            return _Serialize(loDatos, true);
        }

        [HttpPost]
        public string get_buscarCodigoEmr(string codigo, string fechaCarga)
        {
            object loDatos;
            try
            {
                Cls_Negocio_Importacion_Lecturas objeto_negocio = new Cls_Negocio_Importacion_Lecturas();
                loDatos = objeto_negocio.Capa_Negocio_buscarCodigoEMr(codigo, fechaCarga);
            }
            catch (Exception ex)
            {
                throw ex;
            }
            return _Serialize(loDatos, true);
        }


        [HttpPost]
        public string save_GrandesClientesfile(HttpPostedFileBase file, int Id_GrandeCliente, string CodigoEMR, int  id_marcaMedidor,  string fechaCarga)
        {
             List<CorteTemporalCorte> oCortes = new List<CorteTemporalCorte>();
            DateTime _fecha_actual = DateTime.Now;
            string fileLocation = "";
            try
            {
                object loDatos = null;
                int resFile = 0;

                var nombrefile = System.IO.Path.GetFileName(file.FileName);
                string extension = System.IO.Path.GetExtension(file.FileName);
                string correlativo = String.Format("{0:ddMMyyyy_hhmmss}", DateTime.Now);
                string nombreUrl = "";


                if ((nombrefile).ToUpper().Contains("D"))
                {
                    nombreUrl = CodigoEMR + " D-" + correlativo + extension;
                }
                else {
                    nombreUrl = CodigoEMR + "-" + correlativo + extension;
                }


                Cls_Negocio_Importacion_Lecturas objeto_negocio = new Cls_Negocio_Importacion_Lecturas();
                resFile = objeto_negocio.Capa_Negocio_grabarGrandesClienteFile(Id_GrandeCliente, CodigoEMR, file.FileName, nombreUrl, id_marcaMedidor, fechaCarga, ((Sesion)Session["Session_Usuario_Acceso"]).usuario.usu_id);
                
                if (resFile > 0)
                {
                    fileLocation = System.Web.Hosting.HostingEnvironment.MapPath("~/Files_GrandesClientes/" + resFile);
                    file.SaveAs(fileLocation);

                    if (!System.IO.File.Exists(fileLocation))
                    {
                        loDatos = "No se pudo guardar el archivo : ";
                    }
                    else {
                        loDatos = "OK";                    
                    }
                }

                Thread.Sleep(1000);
                ////_rutaServer = ConfigurationManager.AppSettings["servidor-archivos"];
                return _Serialize(loDatos, true);

            }
            catch (Exception ex)
            {
                return _Serialize(ex.Message, true);
            }
        }




    }
}
