using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using DSIGE;
using DSIGE.Modelo;
using DSIGE.Negocio;
using System.Data.OleDb;
using System.Data;
using Newtonsoft.Json;


namespace DSIGE.Web.Controllers
{
    public class VisorNovedadesController : Controller
    {
        //
        // GET: /VisorNovedades/

        public ActionResult VisorNovedades()
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

        public string ListandoOperarios()
        {
            object loDatos;
            try
            {
                VisorNovedades_BL obj_negocio = new VisorNovedades_BL();
                loDatos = obj_negocio.Capa_Negocio_Get_Listado_Operarios();

                return _Serialize(loDatos, true);
            }
            catch (Exception ex)
            {
                return _Serialize(ex.Message, true);
            }
        }
 
        public string ListandoDetallesDocumentos(string id_operario, string fecha_ini, string fecha_fin, string contrato)
        {
            object loDatos;
            try
            {
                VisorNovedades_BL obj_negocio = new VisorNovedades_BL();
                loDatos = obj_negocio.Capa_Negocio_Get_ListandoDetallesDocumentos(id_operario, fecha_ini, fecha_fin, contrato);

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
