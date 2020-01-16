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
    public class VisorArchivosController : Controller
    {
        //
        // GET: /VisorArchivos/

        public ActionResult VisorArchivos()
        {
            return View();
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
                VisorArchivo_BL obj_negocio = new VisorArchivo_BL();
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
                VisorArchivo_BL obj_negocio = new VisorArchivo_BL();
                loDatos = obj_negocio.Capa_Negocio_Get_Listado_Servicios();

                return _Serialize(loDatos, true);
            }
            catch (Exception ex)
            {
                return _Serialize(ex.Message, true);
            }
        }

        public string ListandoDetallesDocumentos(string id_local, string id_servicio, string fecha_ini, string fecha_fin, string contrato)
        {
            object loDatos;
            try
            {
                VisorArchivo_BL obj_negocio = new VisorArchivo_BL();
                loDatos = obj_negocio.Capa_Negocio_Get_ListandoDetallesDocumentos(((Sesion)Session["Session_Usuario_Acceso"]).empresa.emp_id, id_local, id_servicio, fecha_ini, fecha_fin, contrato);

                return _Serialize(loDatos, true);
            }
            catch (Exception ex)
            {
                return _Serialize(ex.Message, true);
            }
        }

        public string DescargarArchivos(string id_Archivo)
        {
            object loDatos;
            try
            {
                VisorArchivo_BL obj_negocio = new VisorArchivo_BL();
                loDatos = obj_negocio.Capa_Negocio_Get_RutaArchivos(id_Archivo);

                return _Serialize(loDatos, true);
            }
            catch (Exception ex)
            {
                return _Serialize(ex.Message, true);
            }
        }

    }
}
