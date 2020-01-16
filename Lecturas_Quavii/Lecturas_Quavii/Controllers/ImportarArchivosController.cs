using DSIGE.Modelo;
using DSIGE.Negocio;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Web;
using System.Web.Mvc;

namespace DSIGE.Web.Controllers
{
    public class ImportarArchivosController : Controller
    {
        //
        // GET: /ImportarArchivos/

        public ActionResult ImportarArchivos()
        {
            return View();
        }
        
        [HttpPost]
        public ActionResult saveFilesToServidor(HttpPostedFileBase file, string local, string servicio, string tipodoc, string contrato,   string fechaAsignacion)
        { 
            DateTime _fecha_actual = DateTime.Now;
            object loDatos = null;
            string nomFile = contrato + "_"+ String.Format("{0:ddMMyyyy_hhmmss}", DateTime.Now);
            string extension =  Path.GetExtension(file.FileName);
            string fileLocation = Server.MapPath("~/Upload") + "\\" + nomFile + extension;
            file.SaveAs(fileLocation);

            ImportarArchivo_BL obj_negocio = new ImportarArchivo_BL();
            string obj = obj_negocio.Capa_Negocio_GuardarArchivos(((Sesion)Session["Session_Usuario_Acceso"]).empresa.emp_id  , ((Sesion)Session["Session_Usuario_Acceso"]).usuario.usu_id, local, servicio, tipodoc, contrato, fechaAsignacion, file.FileName, nomFile + extension );

            return new ContentResult
            {
                Content = MvcApplication._Serialize(obj),
                ContentType = "application/json"
            };
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

        public string ListandoLocales()
        {
            object loDatos;
            try
            {
                ImportarArchivo_BL obj_negocio = new ImportarArchivo_BL();
                loDatos = obj_negocio.Capa_Negocio_Get_Listado_Locales();

                return _Serialize(loDatos, true);
            }
            catch (Exception ex)
            {
                return _Serialize(ex.Message, true);
            }
        }

        public string ListandoServicios()
        {
            object loDatos;
            try
            {
                ImportarArchivo_BL obj_negocio = new ImportarArchivo_BL();
                loDatos = obj_negocio.Capa_Negocio_Get_Listado_Servicios();

                return _Serialize(loDatos, true);
            }
            catch (Exception ex)
            {
                return _Serialize(ex.Message, true);
            }
        }

        public string ListandoTipoDocumento()
        {
            object loDatos;
            try
            {
                ImportarArchivo_BL obj_negocio = new ImportarArchivo_BL();
                loDatos = obj_negocio.Capa_Negocio_Get_TipoDocumentos();

                return _Serialize(loDatos, true);
            }
            catch (Exception ex)
            {
                return _Serialize(ex.Message, true);
            }
        }



    }
}
