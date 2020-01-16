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
    public class GeneracionActasController : Controller
    {
        //
        // GET: /Resumen_Lecturas/

        public ActionResult GeneracionActas_index()
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


        [HttpPost]
        public string Servicios()
        {
            object loDatos;
            try
            {
                GeneracionActas_BL obj_negocio = new GeneracionActas_BL();
                loDatos = obj_negocio.Capa_Negocio_Servicio();
                return _Serialize(loDatos, true);
            }
            catch (Exception ex)
            {
                return _Serialize(ex.Message, true);
            }
        }


        [HttpPost]
        public string Mostrando_informacion_general(int local , int servicio, int operario, string fecha )
        {
            object loDatos;
            try
            {
                GeneracionActas_BL obj_negocio = new GeneracionActas_BL();
                loDatos = obj_negocio.Capa_Negocio_Mostrando_informacion_general(local,servicio, operario, fecha);
                return _Serialize(loDatos, true);
            }
            catch (Exception ex)
            {
                return _Serialize(ex.Message, true);
            }
        }


        [HttpPost]
        public string mostrandoActa(int idCorte, int idServicio)
        {
            object loDatos;
            try
            {
                GeneracionActas_BL obj_negocio = new GeneracionActas_BL();
                loDatos = obj_negocio.Capa_Negocio_MostrandoActa(idCorte, idServicio);
                return _Serialize(loDatos, true);
            }
            catch (Exception ex)
            {
                return _Serialize(ex.Message, true);
            }
        }





        [HttpPost]
        public string Mostrando_informacion_Inspecciones(int servicio, int operario, string fecha ,int tipoReporte)
        {
            object loDatos;
            try
            {
                GeneracionActas_BL obj_negocio = new GeneracionActas_BL();
                loDatos = obj_negocio.Capa_Negocio_Mostrando_informacion_Inspecciones(servicio, operario, fecha, tipoReporte);
                return _Serialize(loDatos, true);
            }
            catch (Exception ex)
            {
                return _Serialize(ex.Message, true);
            }
        }



        [HttpPost]
        public string get_generacionPdf_inspecciones(int id_inspeccion)
        {
            object loDatos;
            try
            {
                GeneracionActas_BL obj_negocio = new GeneracionActas_BL();
                loDatos = obj_negocio.Capa_negocio_get_generacionPdf_inspecciones(id_inspeccion);
                return _Serialize(loDatos, true);
            }
            catch (Exception ex)
            {
                return _Serialize(ex.Message, true);
            }
        }




    }
}

